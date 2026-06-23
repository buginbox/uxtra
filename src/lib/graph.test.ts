import { describe, expect, it } from 'vitest'

import {
  buildDirectoryRef,
  buildUsersFilter,
  extractGraphErrorMessage,
  normalizeBearerToken,
} from './graph'

describe('normalizeBearerToken', () => {
  it('removes the bearer prefix and trims whitespace', () => {
    expect(normalizeBearerToken('  Bearer abc.def  ')).toBe('abc.def')
  })
})

describe('buildUsersFilter', () => {
  it('targets name and email fields and escapes apostrophes', () => {
    expect(buildUsersFilter("O'Hara")).toBe(
      "startswith(displayName,'O''Hara') or startswith(givenName,'O''Hara') or startswith(surname,'O''Hara') or startswith(mail,'O''Hara') or startswith(userPrincipalName,'O''Hara')",
    )
  })

  it('supports first-name then last-name searches', () => {
    expect(buildUsersFilter('Ada Lovelace')).toBe(
      "startswith(displayName,'Ada Lovelace') or startswith(givenName,'Ada Lovelace') or startswith(surname,'Ada Lovelace') or startswith(mail,'Ada Lovelace') or startswith(userPrincipalName,'Ada Lovelace') or (startswith(givenName,'Ada') and startswith(surname,'Lovelace'))",
    )
  })
})

describe('buildDirectoryRef', () => {
  it('builds the Graph reference payload used for member inserts', () => {
    expect(buildDirectoryRef('user-123')).toEqual({
      '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/user-123',
    })
  })
})

describe('extractGraphErrorMessage', () => {
  it('prefers the Graph error payload message', async () => {
    const response = new Response(
      JSON.stringify({ error: { message: 'Insufficient privileges to complete the operation.' } }),
      {
        status: 403,
        statusText: 'Forbidden',
        headers: { 'Content-Type': 'application/json' },
      },
    )

    await expect(extractGraphErrorMessage(response)).resolves.toBe(
      '403 Forbidden: Insufficient privileges to complete the operation.',
    )
  })
})
