// In-memory settings store
let settings: Record<string, any> = {
  theme: 'light',
  notifications: true,
  autoSave: true,
  language: 'en',
}

export function getAllSettings(): Record<string, any> {
  return { ...settings }
}

export function getSetting(key: string): any {
  return settings[key]
}

export function updateSetting(key: string, value: any): void {
  settings[key] = value
}

export function updateSettings(updates: Record<string, any>): void {
  settings = { ...settings, ...updates }
}

export function toggleSetting(key: string): void {
  settings[key] = !settings[key]
}

export function resetAllSettings(): void {
  settings = {
    theme: 'light',
    notifications: true,
    autoSave: true,
    language: 'en',
  }
}

export function exportSettings(): string {
  return JSON.stringify(settings, null, 2)
}

export function importSettings(data: string): void {
  try {
    const imported = JSON.parse(data)
    settings = { ...settings, ...imported }
  } catch (error) {
    console.error('Failed to import settings:', error)
    throw new Error('Invalid settings format')
  }
}

