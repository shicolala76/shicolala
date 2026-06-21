import { useState, useEffect } from 'react'

export default function App() {
  const [settings, setSettings] = useState<Record<string, any>>({})
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    fetch('/api/settings')
      .then(res => res.json())
      .then(data => {
        setSettings(data)
        setLoading(false)
      })
      .catch(err => {
        console.error('Failed to load settings:', err)
        setLoading(false)
      })
  }, [])

  if (loading) return <div>Loading...</div>

  return (
    <div style={{ padding: '20px', fontFamily: 'sans-serif' }}>
      <h1>Settings Manager</h1>
      <pre>{JSON.stringify(settings, null, 2)}</pre>
    </div>
  )
}

