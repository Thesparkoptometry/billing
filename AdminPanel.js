import React, { useState, useEffect } from 'react';
import { db, auth } from './firebase';
import { createUserWithEmailAndPassword, sendPasswordResetEmail } from 'firebase/auth';
import { loadAllUsers, saveUserProfile, deleteUser } from './auth';

const ROLES = [
  { value: 'doctor', label: 'Doctor — sees only their own log' },
  { value: 'biller', label: 'Biller — sees billing sheet + all logs' },
  { value: 'master', label: 'Master — full access (admin)' },
];

export default function AdminPanel({ currentUser }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ name: '', email: '', password: '', role: 'doctor', doctorId: '' });
  const [creating, setCreating] = useState(false);
  const [msg, setMsg] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    loadAllUsers().then(u => { setUsers(u); setLoading(false); });
  }, []);

  const createUser = async () => {
    if (!form.name || !form.email || !form.password || !form.role) return;
    setCreating(true);
    setError('');
    try {
      // Create Firebase Auth account
      const cred = await createUserWithEmailAndPassword(auth, form.email, form.password);
      // Save profile to Firestore
      const profile = {
        uid: cred.user.uid,
        name: form.name,
        email: form.email,
        role: form.role,
        doctorId: form.doctorId || form.name.split(' ').pop().toLowerCase(),
        createdAt: new Date().toISOString(),
      };
      await saveUserProfile(cred.user.uid, profile);
      setUsers(prev => [...prev, profile]);
      setMsg(`✓ Created account for ${form.name}`);
      setForm({ name: '', email: '', password: '', role: 'doctor', doctorId: '' });
      setShowForm(false);
      setTimeout(() => setMsg(''), 3000);
    } catch (e) {
      setError(e.message);
    }
    setCreating(false);
  };

  const removeUser = async (uid, name) => {
    if (!window.confirm(`Remove ${name}? They will no longer be able to sign in.`)) return;
    await deleteUser(uid);
    setUsers(prev => prev.filter(u => u.uid !== uid));
  };

  const resetPassword = async (email, name) => {
    try {
      await sendPasswordResetEmail(auth, email);
      setMsg(`✓ Password reset email sent to ${name}`);
      setTimeout(() => setMsg(''), 3000);
    } catch (e) {
      setError('Could not send reset email: ' + e.message);
    }
  };

  const inp = { padding: '8px 10px', border: '1.5px solid #E2E8F0', borderRadius: 7, fontSize: 13, fontFamily: "'DM Sans',sans-serif", outline: 'none', width: '100%', boxSizing: 'border-box' };
  const lbl = (label, children) => (
    <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <span style={{ fontSize: 11, fontWeight: 700, color: '#64748B', textTransform: 'uppercase', letterSpacing: '0.05em' }}>{label}</span>
      {children}
    </label>
  );

  return (
    <div style={{ maxWidth: 700, margin: '0 auto', padding: '0 20px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
        <div>
          <h2 style={{ fontFamily: "'DM Serif Display',serif", fontSize: 20, color: '#1B3A5C' }}>⚙ User Management</h2>
          <p style={{ fontSize: 12, color: '#94A3B8', marginTop: 2 }}>Add, remove, or reset passwords for billing app users.</p>
        </div>
        <button onClick={() => setShowForm(v => !v)} style={{ background: '#1B3A5C', color: 'white', border: 'none', borderRadius: 8, padding: '9px 18px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
          + Add User
        </button>
      </div>

      {msg && <div style={{ background: '#D1FAE5', border: '1px solid #6EE7B7', borderRadius: 8, padding: '10px 16px', fontSize: 13, color: '#065F46', marginBottom: 14 }}>{msg}</div>}
      {error && <div style={{ background: '#FEE2E2', border: '1px solid #FCA5A5', borderRadius: 8, padding: '10px 16px', fontSize: 13, color: '#DC2626', marginBottom: 14 }}>{error}</div>}

      {showForm && (
        <div style={{ background: 'white', borderRadius: 14, padding: '20px 22px', marginBottom: 20, border: '1px solid #E2E8F0', boxShadow: '0 2px 8px rgba(0,0,0,0.06)' }}>
          <h3 style={{ fontFamily: "'DM Serif Display',serif", fontSize: 16, color: '#1B3A5C', marginBottom: 16 }}>New User</h3>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 }}>
            {lbl('Full Name', <input value={form.name} onChange={e => setForm(f => ({...f, name: e.target.value}))} style={inp} placeholder="Dr. Jane Smith" />)}
            {lbl('Email (for login)', <input type="email" value={form.email} onChange={e => setForm(f => ({...f, email: e.target.value}))} style={inp} placeholder="jane@thesparkoptometry.app" />)}
            {lbl('Temporary Password', <input type="text" value={form.password} onChange={e => setForm(f => ({...f, password: e.target.value}))} style={inp} placeholder="They will change this" />)}
            {lbl('Role', 
              <select value={form.role} onChange={e => setForm(f => ({...f, role: e.target.value}))} style={{ ...inp, cursor: 'pointer' }}>
                {ROLES.map(r => <option key={r.value} value={r.value}>{r.label}</option>)}
              </select>
            )}
            {lbl('Doctor ID (for log matching)', <input value={form.doctorId} onChange={e => setForm(f => ({...f, doctorId: e.target.value}))} style={inp} placeholder="e.g. fan, kaneta, pan" />)}
          </div>
          <p style={{ fontSize: 11, color: '#94A3B8', marginBottom: 14 }}>💡 Doctor ID should match the name used in the production spreadsheet (e.g. "fan", "kaneta"). This is how their log entries are linked.</p>
          <div style={{ display: 'flex', gap: 10 }}>
            <button onClick={createUser} disabled={creating || !form.name || !form.email || !form.password}
              style={{ background: creating ? '#94A3B8' : '#10B981', color: 'white', border: 'none', borderRadius: 8, padding: '9px 20px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
              {creating ? 'Creating…' : 'Create Account'}
            </button>
            <button onClick={() => setShowForm(false)} style={{ background: '#F1F5F9', color: '#64748B', border: 'none', borderRadius: 8, padding: '9px 14px', cursor: 'pointer', fontSize: 13 }}>Cancel</button>
          </div>
        </div>
      )}

      {/* User list */}
      {loading ? <p style={{ color: '#94A3B8' }}>Loading users…</p> : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {users.map(u => (
            <div key={u.uid} style={{ background: 'white', borderRadius: 10, padding: '14px 18px', border: '1px solid #E2E8F0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                  <span style={{ fontWeight: 700, fontSize: 14, color: '#1B3A5C' }}>{u.name}</span>
                  <span style={{ fontSize: 10, fontWeight: 700, borderRadius: 4, padding: '2px 8px', background: u.role === 'master' ? '#1B3A5C' : u.role === 'biller' ? '#2E7D8C' : '#F1F5F9', color: u.role === 'master' ? 'white' : u.role === 'biller' ? 'white' : '#64748B' }}>
                    {u.role}
                  </span>
                  {u.doctorId && <span style={{ fontSize: 11, color: '#94A3B8' }}>ID: {u.doctorId}</span>}
                </div>
                <p style={{ fontSize: 12, color: '#94A3B8', marginTop: 2 }}>{u.email}</p>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={() => resetPassword(u.email, u.name)}
                  style={{ background: '#EFF6FF', border: 'none', borderRadius: 6, padding: '6px 12px', cursor: 'pointer', color: '#1D4ED8', fontSize: 11, fontWeight: 600 }}>
                  Reset Password
                </button>
                {u.uid !== currentUser?.uid && (
                  <button onClick={() => removeUser(u.uid, u.name)}
                    style={{ background: '#FEE2E2', border: 'none', borderRadius: 6, padding: '6px 10px', cursor: 'pointer', color: '#DC2626', fontSize: 11 }}>
                    Remove
                  </button>
                )}
              </div>
            </div>
          ))}
          {users.length === 0 && <p style={{ color: '#94A3B8', fontSize: 13 }}>No users added yet. Add yourself first, then your team.</p>}
        </div>
      )}

      <div style={{ marginTop: 24, padding: '16px 18px', background: '#F8FAFC', borderRadius: 10, border: '1px solid #E2E8F0' }}>
        <p style={{ fontSize: 12, fontWeight: 700, color: '#64748B', marginBottom: 6 }}>SETUP INSTRUCTIONS</p>
        <ol style={{ fontSize: 12, color: '#64748B', lineHeight: 1.8, paddingLeft: 18, margin: 0 }}>
          <li>Create accounts for each doctor using their preferred email</li>
          <li>Give them their temporary password and tell them to change it after first login</li>
          <li>Doctor ID must match the name in the production spreadsheet (lowercase)</li>
          <li>Doctors see only their own log. Dr. Kha (biller) sees everything except Admin.</li>
          <li>Use "Reset Password" to send a password reset email if someone forgets theirs</li>
        </ol>
      </div>
    </div>
  );
}
