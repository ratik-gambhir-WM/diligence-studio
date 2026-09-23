import { describe, expect, it } from 'vitest'

import { parseMicrosoftSupport } from './microsoftSupport'

describe('parseMicrosoftSupport', () => {
  it('enables Microsoft support only for the exact true value', () => {
    expect(parseMicrosoftSupport('true')).toBe(true)
    expect(parseMicrosoftSupport(' TRUE ')).toBe(true)
    expect(parseMicrosoftSupport('1')).toBe(false)
    expect(parseMicrosoftSupport('yes')).toBe(false)
    expect(parseMicrosoftSupport('false')).toBe(false)
    expect(parseMicrosoftSupport(undefined)).toBe(false)
  })
})
