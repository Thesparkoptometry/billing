import React, { useState } from 'react';
import { signIn } from './auth';

export default function Login({ onLogin }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handle = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    try {
      const cred = await signIn(email.trim(), password);
      onLogin(cred.user);
    } catch (err) {
      setError('Invalid email or password.');
    }
    setLoading(false);
  };

  const inp = {
    width: '100%', padding: '11px 14px', border: '1.5px solid #E2E8F0',
    borderRadius: 8, fontSize: 14, fontFamily: "'DM Sans', sans-serif",
    outline: 'none', boxSizing: 'border-box',
  };

  return (
    <div style={{ minHeight: '100vh', background: 'linear-gradient(135deg, #112640 0%, #1B3A5C 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
      <div style={{ background: 'white', borderRadius: 20, padding: '40px 36px', maxWidth: 400, width: '100%', boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }}>
        {/* Logo */}
        <div style={{ textAlign: 'center', marginBottom: 28 }}>
          <svg width="52" height="52" viewBox="0 0 48 48" fill="none" style={{ marginBottom: 12 }}>
            <circle cx="24" cy="24" r="24" fill="#1B3A5C" />
            <path d="M12 24c0-6.627 5.373-12 12-12s12 5.373 12 12" stroke="#2E7D8C" strokeWidth="2.5" strokeLinecap="round" />
            <circle cx="24" cy="24" r="5" fill="#2E7D8C" />
            <circle cx="24" cy="24" r="2" fill="white" />
          </svg>
          <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 22, color: '#1B3A5C', marginBottom: 4 }}>The Spark Billing</h1>
          <p style={{ color: '#94A3B8', fontSize: 13 }}>Sign in to your account</p>
        </div>

        <form onSubmit={handle}>
          <div style={{ marginBottom: 14 }}>
            <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#64748B', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 6 }}>Email</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} style={inp} placeholder="your@email.com" required autoFocus />
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#64748B', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 6 }}>Password</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={inp} placeholder="••••••••" required />
          </div>

          {error && <div style={{ background: '#FEE2E2', color: '#DC2626', borderRadius: 8, padding: '10px 14px', fontSize: 13, marginBottom: 16 }}>{error}</div>}

          <button type="submit" disabled={loading} style={{ width: '100%', padding: '12px', background: loading ? '#94A3B8' : '#1B3A5C', color: 'white', border: 'none', borderRadius: 10, fontSize: 15, fontWeight: 700, cursor: loading ? 'default' : 'pointer', fontFamily: "'DM Sans', sans-serif" }}>
            {loading ? 'Signing in…' : 'Sign In'}
          </button>
        </form>
      </div>
    </div>
  );
}
