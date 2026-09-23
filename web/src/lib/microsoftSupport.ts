export function parseMicrosoftSupport(value: string | undefined) {
  return value?.trim().toLowerCase() === 'true'
}

export const MICROSOFT_SUPPORT = parseMicrosoftSupport(import.meta.env.VITE_MICROSOFT_SUPPORT)
