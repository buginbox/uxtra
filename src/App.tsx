import { useEffect, useMemo, useState, type FormEvent } from 'react'
import './App.css'
import {
  addGroupMember,
  fetchGroupMembers,
  fetchOwnedGroups,
  normalizeBearerToken,
  removeGroupMember,
  renameGroup,
  searchDirectoryObjects,
  type DirectoryCandidate,
  type DirectoryMember,
  type GroupSummary,
} from './lib/graph'

type Notice = {
  kind: 'error' | 'success' | 'info'
  text: string
}

const TOKEN_STORAGE_KEY = 'entra-group-manager-token'

function App() {
  const [tokenDraft, setTokenDraft] = useState('')
  const [accessToken, setAccessToken] = useState('')
  const [groups, setGroups] = useState<GroupSummary[]>([])
  const [selectedGroupId, setSelectedGroupId] = useState('')
  const [members, setMembers] = useState<DirectoryMember[]>([])
  const [searchTerm, setSearchTerm] = useState('')
  const [searchResults, setSearchResults] = useState<DirectoryCandidate[]>([])
  const [searchResultTotal, setSearchResultTotal] = useState(0)
  const [notice, setNotice] = useState<Notice | null>(null)
  const [loadingGroups, setLoadingGroups] = useState(false)
  const [loadingMembers, setLoadingMembers] = useState(false)
  const [searchingUsers, setSearchingUsers] = useState(false)
  const [pendingAddId, setPendingAddId] = useState<string | null>(null)
  const [pendingRemoveId, setPendingRemoveId] = useState<string | null>(null)
  const [renameModalOpen, setRenameModalOpen] = useState(false)
  const [renameDraft, setRenameDraft] = useState('')
  const [descriptionDraft, setDescriptionDraft] = useState('')
  const [renamingGroup, setRenamingGroup] = useState(false)
  const [tokenPanelCollapsed, setTokenPanelCollapsed] = useState(false)

  useEffect(() => {
    const savedToken = window.localStorage.getItem(TOKEN_STORAGE_KEY) ?? ''
    setTokenDraft(savedToken)
    setAccessToken(normalizeBearerToken(savedToken))
  }, [])

  useEffect(() => {
    const normalizedToken = normalizeBearerToken(accessToken)
    const normalizedSearch = searchTerm.trim()

    if (normalizedToken.length < 2 || normalizedSearch.length < 2 || !selectedGroupId) {
      setSearchingUsers(false)
      setSearchResults([])
      setSearchResultTotal(0)
      return
    }

    let cancelled = false
    const timeoutId = window.setTimeout(async () => {
      setSearchingUsers(true)

      try {
        const result = await searchDirectoryObjects(
          normalizedToken,
          normalizedSearch,
          selectedGroupId,
          25,
        )

        if (!cancelled) {
          setSearchResults(result.items)
          setSearchResultTotal(result.total)
        }
      } catch (error) {
        if (!cancelled) {
          setNotice({
            kind: 'error',
            text:
              error instanceof Error
                ? error.message
                : 'Unexpected error while talking to Microsoft Graph.',
          })
          setSearchResults([])
          setSearchResultTotal(0)
        }
      } finally {
        if (!cancelled) {
          setSearchingUsers(false)
        }
      }
    }, 300)

    return () => {
      cancelled = true
      window.clearTimeout(timeoutId)
    }
  }, [searchTerm, selectedGroupId, accessToken])

  const selectedGroup = useMemo(
    () => groups.find((group) => group.id === selectedGroupId) ?? null,
    [groups, selectedGroupId],
  )
  const currentGroupName = selectedGroup?.displayName?.trim() ?? ''
  const currentGroupDescription = selectedGroup?.description?.trim() ?? ''
  const nextGroupName = renameDraft.trim()
  const nextGroupDescription = descriptionDraft.trim()
  const selectedGroupIsMicrosoft365 = selectedGroup?.groupTypes?.includes('Unified') ?? false
  const hasLoadedGroups = groups.length > 0

  async function handleTokenSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    const normalized = normalizeBearerToken(tokenDraft)
    setAccessToken(normalized)
    window.localStorage.setItem(TOKEN_STORAGE_KEY, tokenDraft)
    await loadGroups(normalized)
  }

  async function loadGroups(token = accessToken) {
    const normalized = normalizeBearerToken(token)
    if (!normalized) {
      setNotice({ kind: 'error', text: 'Paste a delegated Microsoft Graph bearer token first.' })
      return
    }

    setLoadingGroups(true)
    setNotice({ kind: 'info', text: 'Loading owned groups from Microsoft Graph…' })

    try {
      const nextGroups = await fetchOwnedGroups(normalized)
      setGroups(nextGroups)

      if (!nextGroups.length) {
        setSelectedGroupId('')
        setMembers([])
        setRenameModalOpen(false)
        setTokenPanelCollapsed(false)
        setNotice({
          kind: 'info',
          text: 'No owned groups returned for this token.',
        })
        return
      }

      const nextGroupId = nextGroups.some((group) => group.id === selectedGroupId)
        ? selectedGroupId
        : nextGroups[0].id

      setSelectedGroupId(nextGroupId)
      setTokenPanelCollapsed(true)
      setNotice({
        kind: 'success',
        text: `Loaded ${nextGroups.length} group${nextGroups.length === 1 ? '' : 's'}.`,
      })
      await loadMembers(nextGroupId, normalized)
    } catch (error) {
      setGroups([])
      setSelectedGroupId('')
      setMembers([])
      setRenameModalOpen(false)
      setTokenPanelCollapsed(false)
      setNotice({
        kind: 'error',
        text:
          error instanceof Error
            ? error.message
            : 'Unexpected error while talking to Microsoft Graph.',
      })
    } finally {
      setLoadingGroups(false)
    }
  }

  async function loadMembers(groupId: string, token = accessToken) {
    const normalized = normalizeBearerToken(token)
    if (!groupId || !normalized) {
      return
    }

    setLoadingMembers(true)

    try {
      const nextMembers = await fetchGroupMembers(normalized, groupId)
      setMembers(nextMembers)
    } catch (error) {
      setMembers([])
      setNotice({
        kind: 'error',
        text:
          error instanceof Error
            ? error.message
            : 'Unexpected error while talking to Microsoft Graph.',
      })
    } finally {
      setLoadingMembers(false)
    }
  }

  async function handleGroupChange(nextGroupId: string) {
    setSelectedGroupId(nextGroupId)
    setSearchTerm('')
    setSearchResults([])
    setSearchResultTotal(0)
    setRenameModalOpen(false)
    await loadMembers(nextGroupId)
  }

  async function handleAddMember(user: DirectoryCandidate) {
    const isSecurityGroup = user['@odata.type'] === '#microsoft.graph.group'

    if (!selectedGroupId || !accessToken) {
      return
    }

    if (isSecurityGroup && selectedGroupIsMicrosoft365) {
      setNotice({
        kind: 'info',
        text: 'Security groups cannot be added inside a Microsoft 365 group.',
      })
      return
    }

    setPendingAddId(user.id)

    try {
      await addGroupMember(accessToken, selectedGroupId, user.id)
      setMembers((currentMembers) =>
        [...currentMembers, user].sort((left, right) => {
          const a =
            left.displayName?.toLocaleLowerCase() ?? left.userPrincipalName?.toLocaleLowerCase() ?? ''
          const b =
            right.displayName?.toLocaleLowerCase() ?? right.userPrincipalName?.toLocaleLowerCase() ?? ''
          return a.localeCompare(b)
        }),
      )
      setSearchResults((currentResults) => currentResults.filter((candidate) => candidate.id !== user.id))
      setSearchResultTotal((currentTotal) => Math.max(currentTotal - 1, 0))
      setSearchTerm('')
      setNotice({
        kind: 'success',
        text: `${user.displayName ?? user.userPrincipalName ?? user.mail ?? user.id} added to ${selectedGroup?.displayName ?? 'the group'}.`,
      })
    } catch (error) {
      setNotice({
        kind: 'error',
        text:
          error instanceof Error
            ? error.message
            : 'Unexpected error while talking to Microsoft Graph.',
      })
    } finally {
      setPendingAddId(null)
    }
  }

  async function handleRemoveMember(member: DirectoryMember) {
    if (!selectedGroupId || !accessToken) {
      return
    }

    setPendingRemoveId(member.id)

    try {
      await removeGroupMember(accessToken, selectedGroupId, member.id)
      setMembers((currentMembers) => currentMembers.filter((entry) => entry.id !== member.id))
      setNotice({
        kind: 'success',
        text: `${member.displayName ?? member.userPrincipalName ?? member.mail ?? member.id} removed from ${selectedGroup?.displayName ?? 'the group'}.`,
      })
    } catch (error) {
      setNotice({
        kind: 'error',
        text:
          error instanceof Error
            ? error.message
            : 'Unexpected error while talking to Microsoft Graph.',
      })
    } finally {
      setPendingRemoveId(null)
    }
  }

  function openRenameModal() {
    if (!selectedGroup) {
      return
    }

    setRenameDraft(selectedGroup.displayName ?? '')
    setDescriptionDraft(selectedGroup.description ?? '')
    setRenameModalOpen(true)
  }

  function closeRenameModal() {
    if (renamingGroup) {
      return
    }

    setRenameModalOpen(false)
  }

  async function handleRenameGroup(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!selectedGroup || !accessToken) {
      return
    }

    const trimmedName = renameDraft.trim()
    if (!trimmedName) {
      setNotice({ kind: 'error', text: 'Group name cannot be empty.' })
      return
    }

    const nextDescription = descriptionDraft.trim()
    const currentName = selectedGroup.displayName?.trim() ?? ''
    const currentDescription = selectedGroup.description?.trim() ?? ''
    if (trimmedName === currentName && nextDescription === currentDescription) {
      setRenameModalOpen(false)
      return
    }

    if (
      !window.confirm(
        `Save changes to "${currentName || 'selected group'}"?\n\nNew name: ${trimmedName}\nNew description: ${nextDescription || '(empty)'}`,
      )
    ) {
      return
    }

    setRenamingGroup(true)

    try {
      await renameGroup(accessToken, selectedGroup.id, trimmedName, nextDescription)
      setGroups((currentGroups) =>
        currentGroups.map((group) =>
          group.id === selectedGroup.id
            ? { ...group, displayName: trimmedName, description: nextDescription }
            : group,
        ),
      )
      setRenameModalOpen(false)
      setNotice({
        kind: 'success',
        text: `Group updated: ${trimmedName}.`,
      })
    } catch (error) {
      setNotice({
        kind: 'error',
        text:
          error instanceof Error
            ? error.message
            : 'Unexpected error while talking to Microsoft Graph.',
      })
    } finally {
      setRenamingGroup(false)
    }
  }

  return (
    <>
      <main className="app-shell">
        <section className={`panel token-panel ${tokenPanelCollapsed ? 'collapsed' : ''}`}>
          <div className="panel-heading token-panel-heading">
            <div>
              <p className="eyebrow">Manual token</p>
              <h1>Entra ID group membership editor</h1>
              <p className="lede">
                Paste a delegated Microsoft Graph bearer token, load the groups you own, then add
                or remove members from one place.
              </p>
              <p className="token-helper">
                Need a token? Open{' '}
                <a
                  href="https://developer.microsoft.com/en-us/graph/graph-explorer"
                  target="_blank"
                  rel="noreferrer"
                >
                  Microsoft Graph Explorer
                </a>{' '}
                to sign in to the right tenant with the right account, then open the{' '}
                <strong>Access token</strong> tab and paste that bearer token here.
              </p>
              {tokenPanelCollapsed ? (
                <p className="token-summary">
                  Token ready. {groups.length} loaded group{groups.length === 1 ? '' : 's'}.
                </p>
              ) : null}
            </div>
            <div className="token-panel-actions">
              {hasLoadedGroups ? (
                <button
                  type="button"
                  className="ghost-button"
                  onClick={() => setTokenPanelCollapsed((collapsed) => !collapsed)}
                >
                  {tokenPanelCollapsed ? 'Show token panel' : 'Hide token panel'}
                </button>
              ) : null}
            </div>
          </div>

          {!tokenPanelCollapsed ? (
            <>
              <form className="token-form" onSubmit={handleTokenSubmit}>
                <label className="field" htmlFor="access-token">
                  <span>Bearer access token</span>
                  <textarea
                    id="access-token"
                    value={tokenDraft}
                    onChange={(event) => setTokenDraft(event.target.value)}
                    placeholder="Paste Bearer eyJ0eXAiOiJKV1QiLCJhbGciOi..."
                    spellCheck={false}
                    rows={3}
                  />
                </label>

                <div className="token-actions">
                  <button type="submit" disabled={!tokenDraft.trim() || loadingGroups}>
                    {loadingGroups ? 'Loading groups…' : 'Load owned groups'}
                  </button>
                  <button
                    type="button"
                    className="ghost-button"
                    onClick={() => {
                      setTokenDraft('')
                      setAccessToken('')
                      setGroups([])
                      setSelectedGroupId('')
                      setMembers([])
                      setSearchTerm('')
                      setSearchResults([])
                      setSearchResultTotal(0)
                      setNotice(null)
                      setRenameModalOpen(false)
                      setTokenPanelCollapsed(false)
                      window.localStorage.removeItem(TOKEN_STORAGE_KEY)
                    }}
                  >
                    Clear token
                  </button>
                </div>
              </form>

              <div className="tips-grid">
                <article className="tip-card">
                  <h2>Delegated permissions</h2>
                  <p>Recommended scopes: Group.ReadWrite.All and User.Read.All.</p>
                </article>
                <article className="tip-card">
                  <h2>Graph calls used</h2>
                  <ul>
                    <li>GET /me/ownedObjects/microsoft.graph.group</li>
                    <li>PATCH /groups/{'{id}'}</li>
                    <li>GET /groups/{'{id}'}/members</li>
                    <li>GET /users?$filter=startswith(...)</li>
                    <li>GET /groups?$filter=securityEnabled eq true and startswith(...)</li>
                    <li>POST and DELETE /groups/{'{id}'}/members/$ref</li>
                  </ul>
                </article>
              </div>
            </>
          ) : null}

          {notice ? <p className={`notice ${notice.kind}`}>{notice.text}</p> : null}
        </section>

        <section className="workspace-grid">
          <section className="panel group-panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">Step 1</p>
                <h2>Select a group</h2>
              </div>
              <div className="panel-actions">
                <span className="count-pill">{groups.length} loaded</span>
                <button
                  type="button"
                  className="ghost-button icon-button"
                  onClick={openRenameModal}
                  disabled={!selectedGroup}
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path
                      d="M4 20h4l10.5-10.5a2.12 2.12 0 1 0-3-3L5 17v3Zm2.5-2.5 9-9 1.5 1.5-9 9H6.5v-1.5Z"
                      fill="currentColor"
                    />
                  </svg>
                  <span>Edit</span>
                </button>
              </div>
            </div>

            <label className="field" htmlFor="group-picker">
              <span>Owned groups</span>
              <select
                id="group-picker"
                value={selectedGroupId}
                onChange={(event) => void handleGroupChange(event.target.value)}
                disabled={!groups.length || loadingGroups}
              >
                {!groups.length ? <option value="">Load groups first</option> : null}
                {groups.map((group) => (
                  <option key={group.id} value={group.id}>
                    {group.displayName ?? '(Unnamed group)'}
                  </option>
                ))}
              </select>
            </label>

            {selectedGroup ? (
              <dl className="group-meta">
                <div>
                  <dt>Type</dt>
                  <dd>{describeGroupType(selectedGroup)}</dd>
                </div>
                <div>
                  <dt>Mail</dt>
                  <dd>{selectedGroup.mail || '—'}</dd>
                </div>
                <div>
                  <dt>Description</dt>
                  <dd>{selectedGroup.description || '—'}</dd>
                </div>
              </dl>
            ) : (
              <p className="empty-state">Choose a token, then load owned groups.</p>
            )}
          </section>

          <section className="panel picker-panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">Step 2</p>
                <h2>Add a member</h2>
              </div>
              <span className="count-pill">
                {searchResults.length} shown / {searchResultTotal} total
              </span>
            </div>

            <label className="field" htmlFor="member-search">
              <span>People picker</span>
              <input
                id="member-search"
                type="search"
                value={searchTerm}
                onChange={(event) => setSearchTerm(event.target.value)}
                placeholder="Search by first name, last name, email or security group"
                disabled={!selectedGroupId || !accessToken}
              />
            </label>

            <p className="helper-text">
              Enter at least two characters. Results come from Microsoft Graph users and security
              groups. Nested security groups are disabled when the selected target is a Microsoft
              365 group.
            </p>

            <div className="result-panel">
              <div className="result-list">
                {!selectedGroupId ? (
                  <p className="empty-state">Select a group to start adding users or security groups.</p>
                ) : null}
                {selectedGroupId && searchTerm.trim().length < 2 ? (
                  <p className="empty-state">Type a name, email or group name to see suggestions.</p>
                ) : null}
                {searchingUsers ? <p className="empty-state">Searching directory…</p> : null}
                {!searchingUsers &&
                  searchResults.map((user) => {
                    const alreadyMember = members.some((member) => member.id === user.id)
                    const isSecurityGroup = user['@odata.type'] === '#microsoft.graph.group'
                    const nestedGroupBlocked = isSecurityGroup && selectedGroupIsMicrosoft365

                    return (
                      <article key={user.id} className="person-row">
                        <div>
                          <strong>{user.displayName ?? user.userPrincipalName ?? user.mail ?? user.id}</strong>
                          <p>{describePerson(user)}</p>
                        </div>
                        <button
                          type="button"
                          onClick={() => void handleAddMember(user)}
                          disabled={alreadyMember || nestedGroupBlocked || pendingAddId === user.id}
                          aria-label={`Add ${isSecurityGroup ? 'security group' : 'member'} ${user.displayName ?? user.id}`}
                          title={
                            nestedGroupBlocked
                              ? 'Security groups cannot be added inside Microsoft 365 groups.'
                              : undefined
                          }
                        >
                          {alreadyMember
                            ? 'Already in group'
                            : nestedGroupBlocked
                              ? 'Not supported'
                              : pendingAddId === user.id
                                ? 'Adding…'
                                : 'Add'}
                        </button>
                      </article>
                    )
                  })}
                {selectedGroupId &&
                searchTerm.trim().length >= 2 &&
                !searchingUsers &&
                searchResultTotal === 0 ? (
                  <p className="empty-state">No matching users or security groups returned for this search.</p>
                ) : null}
              </div>
            </div>
          </section>
        </section>

        <section className="panel members-panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">Step 3</p>
              <h2>Current members</h2>
            </div>
            <div className="member-actions">
              <span className="count-pill">
                {members.length} member{members.length === 1 ? '' : 's'}
              </span>
              <button
                type="button"
                className="ghost-button"
                onClick={() => void loadMembers(selectedGroupId)}
                disabled={!selectedGroupId || loadingMembers}
              >
                {loadingMembers ? 'Refreshing…' : 'Refresh members'}
              </button>
            </div>
          </div>

          {!selectedGroupId ? (
            <p className="empty-state">No group selected yet.</p>
          ) : loadingMembers ? (
            <p className="empty-state">Loading current membership…</p>
          ) : !members.length ? (
            <p className="empty-state">This group currently has no members returned by Graph.</p>
          ) : (
            <div className="member-list">
              {members.map((member) => (
                <article key={member.id} className="person-row">
                  <div>
                    <strong>{member.displayName ?? member.userPrincipalName ?? member.mail ?? member.id}</strong>
                    <p>{describePerson(member)}</p>
                  </div>
                  <button
                    type="button"
                    className="danger-button"
                    onClick={() => void handleRemoveMember(member)}
                    disabled={pendingRemoveId === member.id}
                  >
                    {pendingRemoveId === member.id ? 'Removing…' : 'Remove'}
                  </button>
                </article>
              ))}
            </div>
          )}
        </section>
      </main>

      {renameModalOpen && selectedGroup ? (
        <div className="modal-backdrop" onClick={closeRenameModal}>
          <div
            className="modal-card"
            role="dialog"
            aria-modal="true"
            aria-labelledby="rename-group-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="modal-header">
              <div>
                <p className="eyebrow">Edit group</p>
                <h2 id="rename-group-title">Update selected group</h2>
              </div>
              <button
                type="button"
                className="ghost-button modal-close-button"
                onClick={closeRenameModal}
                disabled={renamingGroup}
                aria-label="Close rename dialog"
              >
                ×
              </button>
            </div>

            <form className="modal-form" onSubmit={handleRenameGroup}>
              <label className="field" htmlFor="group-name">
                <span>Display name</span>
                <input
                  id="group-name"
                  type="text"
                  value={renameDraft}
                  onChange={(event) => setRenameDraft(event.target.value)}
                  placeholder="Enter a new group name"
                  autoFocus
                />
              </label>

              <label className="field" htmlFor="group-description">
                <span>Description</span>
                <textarea
                  id="group-description"
                  value={descriptionDraft}
                  onChange={(event) => setDescriptionDraft(event.target.value)}
                  placeholder="Add or update the group description"
                  rows={4}
                />
              </label>

              <p className="helper-text">
                Description can be left empty. Saving will ask for confirmation before the PATCH
                request is sent to Microsoft Graph.
              </p>

              <div className="modal-actions">
                <button
                  type="button"
                  className="ghost-button"
                  onClick={closeRenameModal}
                  disabled={renamingGroup}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={
                    renamingGroup ||
                    !nextGroupName ||
                    (nextGroupName === currentGroupName &&
                      nextGroupDescription === currentGroupDescription)
                  }
                >
                  {renamingGroup ? 'Saving…' : 'Save'}
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
    </>
  )
}

function describeGroupType(group: GroupSummary): string {
  if (group.groupTypes?.includes('Unified')) {
    return 'Microsoft 365'
  }

  if (group.securityEnabled) {
    return 'Security'
  }

  if (group.mailEnabled) {
    return 'Distribution'
  }

  return 'Group'
}

function describePerson(
  person: Pick<
    DirectoryMember,
    'givenName' | 'surname' | 'mail' | 'userPrincipalName' | '@odata.type'
  > & { description?: string | null },
): string {
  if (person['@odata.type'] === '#microsoft.graph.group') {
    return ['Security group', person.description ?? person.mail].filter(Boolean).join(' · ')
  }

  const identity = [person.givenName, person.surname, person.mail ?? person.userPrincipalName]
    .filter(Boolean)
    .join(' · ')

  if (identity) {
    return identity
  }

  return person['@odata.type']?.replace('#microsoft.graph.', '') ?? 'Directory object'
}

export default App
