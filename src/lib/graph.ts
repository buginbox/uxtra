export type GroupSummary = {
  id: string
  displayName?: string | null
  description?: string | null
  mail?: string | null
  groupTypes?: string[]
  mailEnabled?: boolean
  securityEnabled?: boolean
}

export type DirectoryMember = {
  id: string
  displayName?: string | null
  givenName?: string | null
  surname?: string | null
  mail?: string | null
  userPrincipalName?: string | null
  '@odata.type'?: string
}

export type UserCandidate = {
  id: string
  displayName?: string | null
  givenName?: string | null
  surname?: string | null
  mail?: string | null
  userPrincipalName?: string | null
}

type GraphCollection<T> = {
  value: T[]
  '@odata.nextLink'?: string
}

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'

export function normalizeBearerToken(input: string): string {
  return input.trim().replace(/^Bearer\s+/i, '')
}

export function buildUsersFilter(term: string): string {
  const trimmed = term.trim()
  const escaped = trimmed.replace(/'/g, "''")
  const filters = [
    `startswith(displayName,'${escaped}')`,
    `startswith(givenName,'${escaped}')`,
    `startswith(surname,'${escaped}')`,
    `startswith(mail,'${escaped}')`,
    `startswith(userPrincipalName,'${escaped}')`,
  ]

  const nameParts = trimmed.split(/\s+/).filter(Boolean)
  if (nameParts.length >= 2) {
    const firstName = nameParts[0].replace(/'/g, "''")
    const remainingName = nameParts.slice(1).join(' ').replace(/'/g, "''")
    filters.push(
      `(startswith(givenName,'${firstName}') and startswith(surname,'${remainingName}'))`,
    )
  }

  return filters.join(' or ')
}

export function buildDirectoryRef(objectId: string): { '@odata.id': string } {
  return {
    '@odata.id': `${GRAPH_ROOT}/directoryObjects/${objectId}`,
  }
}

export async function fetchOwnedGroups(token: string): Promise<GroupSummary[]> {
  const groups = await fetchAllPages<GroupSummary>(
    '/me/ownedObjects/microsoft.graph.group?$select=id,displayName,description,mail,groupTypes,mailEnabled,securityEnabled',
    token,
  )

  return [...groups].sort((left, right) => {
    const a = left.displayName?.toLocaleLowerCase() ?? ''
    const b = right.displayName?.toLocaleLowerCase() ?? ''
    return a.localeCompare(b)
  })
}

export async function fetchGroupMembers(
  token: string,
  groupId: string,
): Promise<DirectoryMember[]> {
  const members = await fetchAllPages<DirectoryMember>(
    `/groups/${groupId}/members?$select=id,displayName,givenName,surname,mail,userPrincipalName`,
    token,
  )

  return [...members].sort((left, right) => {
    const a = left.displayName?.toLocaleLowerCase() ?? ''
    const b = right.displayName?.toLocaleLowerCase() ?? ''
    return a.localeCompare(b)
  })
}

export async function searchDirectoryUsers(
  token: string,
  term: string,
  limit = 8,
): Promise<UserCandidate[]> {
  const params = new URLSearchParams({
    '$select': 'id,displayName,givenName,surname,mail,userPrincipalName',
    '$top': String(limit),
    '$filter': buildUsersFilter(term),
  })

  const response = await graphJson<GraphCollection<UserCandidate>>(
    `/users?${params.toString()}`,
    token,
  )

  return response.value
}

export async function renameGroup(
  token: string,
  groupId: string,
  displayName: string,
  description: string,
): Promise<void> {
  await graphRequest(`/groups/${groupId}`, token, {
    method: 'PATCH',
    body: JSON.stringify({ displayName, description }),
  })
}

export async function addGroupMember(
  token: string,
  groupId: string,
  userId: string,
): Promise<void> {
  await graphRequest(`/groups/${groupId}/members/$ref`, token, {
    method: 'POST',
    body: JSON.stringify(buildDirectoryRef(userId)),
  })
}

export async function removeGroupMember(
  token: string,
  groupId: string,
  memberId: string,
): Promise<void> {
  await graphRequest(`/groups/${groupId}/members/${memberId}/$ref`, token, {
    method: 'DELETE',
  })
}

export async function graphJson<T>(
  pathOrUrl: string,
  token: string,
  init?: RequestInit,
): Promise<T> {
  const response = await graphRequest(pathOrUrl, token, init)
  return (await response.json()) as T
}

export async function graphRequest(
  pathOrUrl: string,
  token: string,
  init: RequestInit = {},
): Promise<Response> {
  const response = await fetch(
    pathOrUrl.startsWith('http') ? pathOrUrl : `${GRAPH_ROOT}${pathOrUrl}`,
    {
      ...init,
      headers: {
        Accept: 'application/json',
        Authorization: `Bearer ${token}`,
        ...(init.body ? { 'Content-Type': 'application/json' } : {}),
        ...init.headers,
      },
    },
  )
  if (!response.ok) {
    throw new Error(await extractGraphErrorMessage(response))
  }

  return response
}

export async function extractGraphErrorMessage(response: Response): Promise<string> {
  try {
    const payload = (await response.json()) as {
      error?: { message?: string; code?: string }
    }

    if (payload.error?.message) {
      return `${response.status} ${response.statusText}: ${payload.error.message}`
    }
  } catch {
    // Ignore JSON parsing failures and fall back to response text.
  }

  const fallback = await response.text().catch(() => '')
  if (fallback.trim()) {
    return `${response.status} ${response.statusText}: ${fallback.trim()}`
  }

  return `${response.status} ${response.statusText}`
}

async function fetchAllPages<T>(pathOrUrl: string, token: string): Promise<T[]> {
  let nextUrl: string | undefined = pathOrUrl
  const items: T[] = []

  while (nextUrl) {
    const payload: GraphCollection<T> = await graphJson<GraphCollection<T>>(nextUrl, token)
    items.push(...payload.value)
    nextUrl = payload['@odata.nextLink']
  }

  return items
}
