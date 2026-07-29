import { useState, useEffect } from 'react'
import { FileBadge, Database, Layers, Sun, Moon } from 'lucide-react'
import './index.css'
import DataTab from './tabs/DataTab'
import GenerateTab from './tabs/GenerateTab'

type Tab = 'data' | 'generate'

export default function App() {
  const [tab, setTab] = useState<Tab>('data')
  // Shared flag: has generation been run at least once?
  const [, setHasManifest] = useState(false)
  const [isDark, setIsDark] = useState(() => localStorage.getItem('theme') === 'dark')

  useEffect(() => {
    if (isDark) {
      document.documentElement.classList.add('dark')
      localStorage.setItem('theme', 'dark')
    } else {
      document.documentElement.classList.remove('dark')
      localStorage.setItem('theme', 'light')
    }
  }, [isDark])

  return (
    <div className="container">
      {/* Header */}
      <header className="app-header">
        <FileBadge size={32} color="var(--primary)" strokeWidth={2.5} />
        <h1>Certificate Automation</h1>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 16 }}>
          <span className="badge-success">v1.0</span>
          <button 
            onClick={() => setIsDark(!isDark)}
            className="btn btn-ghost btn-sm"
            style={{ padding: 8, borderRadius: '50%' }}
            title="Toggle Theme"
          >
            {isDark ? <Sun size={20} /> : <Moon size={20} />}
          </button>
        </div>
      </header>

      {/* Tab bar */}
      <nav className="tab-list">
        <button
          className={`tab-btn ${tab === 'data' ? 'active' : ''}`}
          onClick={() => setTab('data')}
        >
          <Database size={18} /> Data & Template
        </button>
        <button
          className={`tab-btn ${tab === 'generate' ? 'active' : ''}`}
          onClick={() => setTab('generate')}
        >
          <Layers size={18} /> Generate & Send
        </button>
      </nav>

      {/* Tab content */}
      <main className="tab-content">
        {tab === 'data'     && <DataTab />}
        {tab === 'generate' && <GenerateTab onManifestReady={() => setHasManifest(true)} />}
      </main>
    </div>
  )
}
