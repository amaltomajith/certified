import { useState, useRef, useEffect } from 'react'
import { Settings, Play, RefreshCw, Mail, FileDown, Loader2, Send } from 'lucide-react'
import { streamPost, getManifest, pdfPreviewUrl, downloadPdf, type ManifestSummary } from '../api'

type Props = { onManifestReady: () => void }

function logClass(line: string): string {
  if (line.startsWith('ERROR')) return 'log-error'
  if (line.startsWith('WARNING') || line.startsWith('WARN')) return 'log-warn'
  if (line === '__DONE__') return 'log-done'
  if (line.includes('successfully') || line.includes('All ')) return 'log-ok'
  return 'log-info'
}

export default function GenerateTab({ onManifestReady }: Props) {
  const [logs, setLogs] = useState<string[]>([])
  const [running, setRunning] = useState(false)
  const [runningType, setRunningType] = useState<'generate' | 'send' | null>(null)
  const [done, setDone] = useState(false)
  const [error, setError] = useState('')
  const [manifest, setManifest] = useState<ManifestSummary | null>(null)
  const [sendingPoc, setSendingPoc] = useState<string | null>(null)
  const [sentPocs, setSentPocs] = useState<Record<string, boolean>>({})
  const logRef = useRef<HTMLDivElement>(null)

  // Auto-scroll log pane
  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight
  }, [logs])

  const handleGenerate = async () => {
    setRunning(true); setRunningType('generate'); setDone(false); setError(''); setLogs([]); setManifest(null)

    try {
      await streamPost('/generate', {}, (msg) => {
        if (msg === '__DONE__') {
          setDone(true)
        } else {
          setLogs(prev => [...prev, msg])
        }
      })
    } catch (e: any) {
      setError(e.message)
    } finally {
      setRunning(false)
    }

    // Fetch manifest summary after generation
    try {
      const m = await getManifest()
      if (m.total > 0) {
        setManifest(m)
        onManifestReady()
      }
    } catch { /* manifest fetch is non-critical */ }
  }

  const handleSend = async (pocEmail?: string) => {
    setRunning(true); setRunningType('send'); setDone(false); setError(''); setLogs([]);
    if (pocEmail) {
      setSendingPoc(pocEmail)
    }
    try {
      await streamPost('/send', { poc_email: pocEmail }, (msg) => {
        if (msg === '__DONE__') {
          setDone(true)
        } else if (msg.startsWith('__RESULTS__:')) {
          try {
            const raw = msg.slice(12)
            const results = JSON.parse(raw)
            if (Array.isArray(results)) {
              const newSent: Record<string, boolean> = {}
              results.forEach((r: any) => {
                if (r.status === 'sent' && r.sent_to) {
                  newSent[r.sent_to] = true
                }
              })
              setSentPocs(prev => ({ ...prev, ...newSent }))
            }
          } catch (e) {
            console.error('Failed to parse send results:', e)
          }
        } else {
          setLogs(prev => [...prev, msg])
          // Instantly parse live progress line for visual updates
          if (msg.startsWith('Sent →') || msg.startsWith('Sent \u2192')) {
            const match = msg.match(/(?:Sent →|Sent \u2192)\s+([^\s]+)/)
            if (match && match[1]) {
              const email = match[1].trim()
              setSentPocs(prev => ({ ...prev, [email]: true }))
            }
          }
        }
      })
    } catch (e: any) {
      setError(e.message)
    } finally {
      setRunning(false)
      setSendingPoc(null)
    }
  }

  const hasError = logs.some(l => l.startsWith('ERROR'))

  return (
    <div>
      <div className="card">
        <div className="card-header"><Settings size={18} /> Generate Certificates</div>
        <div className="card-body">
          <p className="text-muted" style={{ marginBottom: 16 }}>
            Reads your Excel roster and template, fills in the data, and converts each one to PDF via LibreOffice.
          </p>

          <button
            className={`btn ${hasError && runningType === 'generate' ? 'btn-danger' : 'btn-primary'}`}
            onClick={handleGenerate}
            disabled={running}
            style={{ marginBottom: 16 }}
          >
            {running && runningType === 'generate' ? <><Loader2 size={16} className="animate-spin" /> Generating…</> : done && !hasError && runningType === 'generate' ? <><RefreshCw size={16} /> Regenerate</> : <><Play size={16} fill="currentColor" /> Generate Certificates</>}
          </button>

          {error && <div className="alert alert-danger">{error}</div>}

          {/* Live log pane */}
          {(logs.length > 0 || running) && (
            <div className="log-pane" ref={logRef}>
              {logs.length === 0 && running && (
                <span className="log-info">Starting…</span>
              )}
              {logs.map((line, i) => (
                <div key={i} className={logClass(line)}>{line}</div>
              ))}
              {running && <span className="log-info">▌</span>}
              {done && !hasError && <div className="log-done">── Done ──</div>}
            </div>
          )}
        </div>
      </div>

      {/* Summary card */}
      {manifest && manifest.total > 0 && (
        <div className="card">
          <div className="card-header">
            📋 Generation Summary
            <span className="badge-success">{manifest.total} certificates</span>
          </div>
          <div className="card-body">
            {manifest.groups.map((g, i) => (
              <div key={i} className="card" style={{ marginBottom: 16, border: '1px solid var(--border)' }}>
                <div className="card-header" style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                  <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 600, display: 'flex', alignItems: 'center', gap: 6 }}>
                    <Mail size={16} className="text-muted" /> {g.poc_email}
                  </span>
                  <span className="badge-success">{g.student_count} cert{g.student_count !== 1 ? 's' : ''}</span>
                  
                   <button 
                    className={`btn btn-sm ${sentPocs[g.poc_email] ? 'btn-ghost-primary' : 'btn-primary'}`}
                    style={{ 
                      marginLeft: 'auto', 
                      display: 'flex', 
                      alignItems: 'center', 
                      gap: 6,
                      ...(running && sendingPoc !== g.poc_email ? { opacity: 1, cursor: 'not-allowed' } : {})
                    }}
                    onClick={() => handleSend(g.poc_email)}
                    disabled={running}
                  >
                    {running && runningType === 'send' && sendingPoc === g.poc_email ? (
                      <>
                        <Loader2 size={14} className="animate-spin" />
                        Sending…
                      </>
                    ) : sentPocs[g.poc_email] ? (
                      <>
                        <RefreshCw size={14} />
                        Resend
                      </>
                    ) : (
                      <>
                        <Send size={14} />
                        Send to this POC
                      </>
                    )}
                  </button>
                </div>
                <div className="card-body" style={{ padding: '8px 16px' }}>
                  <table className="data-table" style={{ fontSize: 13 }}>
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Student Name</th>
                        <th>School</th>
                        <th>Event</th>
                        <th>Download PDF</th>
                      </tr>
                    </thead>
                    <tbody>
                      {g.students.map((s, j) => {
                        const filename = g.pdf_files?.[j];
                        const detail = g.student_details?.[j];
                        return (
                          <tr key={j}>
                            <td style={{ color: 'var(--muted)', width: 32 }}>{j + 1}</td>
                            <td><strong>{s}</strong></td>
                            <td>{detail?.school || <span style={{ color: 'var(--muted)' }}>—</span>}</td>
                            <td>{detail?.event_name || <span style={{ color: 'var(--muted)' }}>—</span>}</td>
                            <td>
                              {filename ? (
                                <button
                                  className="btn btn-ghost btn-sm"
                                  style={{
                                    display: 'inline-flex', alignItems: 'center', gap: 4,
                                    color: 'var(--primary)', fontSize: '12px', fontWeight: 500,
                                    padding: '2px 8px',
                                  }}
                                  onClick={() => downloadPdf(filename)}
                                  title={`Download ${filename}`}
                                >
                                  <FileDown size={14} /> Download
                                </button>
                              ) : <span style={{ color: 'var(--muted)' }}>—</span>}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}

            {/* Inline PDF preview */}
            {manifest.first_pdf && (
              <div>
                <div className="section-title" style={{ marginBottom: 8 }}>
                  Sample PDF Preview — {manifest.first_pdf}
                  <button
                    className="btn btn-ghost btn-sm"
                    style={{ marginLeft: 12 }}
                    onClick={() => downloadPdf(manifest.first_pdf!)}
                  >
                    <FileDown size={14} /> Download
                  </button>
                </div>
                <div style={{ width: '100%', borderRadius: 6, border: '1px solid var(--border)', overflow: 'hidden', backgroundColor: '#f3f4f6', display: 'flex', justifyContent: 'center', padding: '20px 0' }}>
                  <img
                    src={pdfPreviewUrl(manifest.first_pdf)}
                    alt="PDF Preview"
                    style={{ maxWidth: '90%', height: 'auto', boxShadow: '0 4px 12px rgba(0,0,0,0.15)', borderRadius: 4 }}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}
