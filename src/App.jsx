import { useState, useEffect, useRef } from 'react'
import UUIDMapperAgent from './UUIDMapperAgent'

const PW = btoa('PAR1')
const SESSION_KEY = 'par_ddm_auth'

function PasswordGate({ onUnlock }) {
  const [value, setValue] = useState('')
  const [error, setError]  = useState('')
  const inputRef = useRef(null)

  useEffect(() => { inputRef.current?.focus() }, [])

  const check = () => {
    if (btoa(value.trim()) === PW) {
      sessionStorage.setItem(SESSION_KEY, PW)
      onUnlock()
    } else {
      setError('Incorrect password. Please try again.')
      setValue('')
      inputRef.current?.focus()
      setTimeout(() => setError(''), 3000)
    }
  }

  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 999,
      background: '#f4f5f9', display: 'flex',
      alignItems: 'center', justifyContent: 'center',
      fontFamily: "'Inter', system-ui, sans-serif",
    }}>
      <div style={{
        background: '#fff', border: '1px solid #d9dae8', borderRadius: 14,
        padding: '44px 40px', width: '100%', maxWidth: 360, textAlign: 'center',
        boxShadow: '0 4px 32px rgba(0,0,0,.08)',
      }}>
        <img
          src="https://partech.com/wp-content/uploads/2025/11/par-logo.svg"
          alt="PAR"
          style={{ height: 30, marginBottom: 22 }}
        />
        <div style={{ height: 2, background: 'linear-gradient(90deg,#f97316,#7c3aed)', borderRadius: 2, marginBottom: 24 }} />
        <h2 style={{ fontSize: 15, fontWeight: 800, color: '#0f0e2a', margin: '0 0 6px' }}>
          DDM File Builder
        </h2>
        <p style={{ fontSize: 12, color: '#6b7280', margin: '0 0 22px', lineHeight: 1.6 }}>
          This tool is for PAR Retail internal use.<br />
          Enter the access password to continue.
        </p>
        <input
          ref={inputRef}
          type="password"
          value={value}
          onChange={e => setValue(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') check() }}
          placeholder="Password"
          autoComplete="off"
          style={{
            width: '100%', padding: '10px 13px', boxSizing: 'border-box',
            border: '1px solid #d9dae8', borderRadius: 7, fontSize: 14,
            fontFamily: 'inherit', background: '#f4f5f9', color: '#0f0e2a',
            outline: 'none', textAlign: 'center', letterSpacing: '0.12em',
            marginBottom: 10,
          }}
        />
        <button
          onClick={check}
          style={{
            width: '100%', padding: '10px 0', border: 'none', borderRadius: 7,
            background: 'linear-gradient(135deg,#f97316,#7c3aed)', color: '#fff',
            fontSize: 13, fontWeight: 700, fontFamily: 'inherit', cursor: 'pointer',
          }}
        >
          Access Tool →
        </button>
        {error && (
          <div style={{ marginTop: 10, fontSize: 12, color: '#dc2626' }}>{error}</div>
        )}
      </div>
    </div>
  )
}

export default function App() {
  const [unlocked, setUnlocked] = useState(
    () => sessionStorage.getItem(SESSION_KEY) === PW
  )

  if (!unlocked) return <PasswordGate onUnlock={() => setUnlocked(true)} />

  return (
    <div className="App">
      <UUIDMapperAgent />
    </div>
  )
}
