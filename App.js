import React, { useState, useEffect } from 'react';
import { onAuthStateChanged, signOut } from 'firebase/auth';
import { auth } from './firebase';
import { loadUserProfile, MASTER_EMAIL } from './auth';
import Login from './Login';
import DoctorLog from './DoctorLog';
import BillingSheet from './BillingSheet';
import AdminPanel from './AdminPanel';
import './index.css';

export default function App() {
  const [user, setUser] = useState(null);
  const [profile, setProfile] = useState(null);
  const [loading, setLoading] = useState(true);
  const [page, setPage] = useState('main');

  useEffect(() => {
    return onAuthStateChanged(auth, async (u) => {
      if (u) {
        setUser(u);
        // Load profile from Firestore
        const p = await loadUserProfile(u.uid);
        if (p) {
          setProfile(p);
        } else if (u.email === MASTER_EMAIL) {
          setProfile({ name: 'Christine', role: 'master', uid: u.uid, email: u.email });
        } else {
          // Unknown user — sign out
          await signOut(auth);
          setUser(null);
          setProfile(null);
        }
      } else {
        setUser(null);
        setProfile(null);
      }
      setLoading(false);
    });
  }, []);

  const handleLogin = (u) => setUser(u);
  const handleSignOut = () => signOut(auth);

  if (loading) return <LoadingScreen />;
  if (!user || !profile) return <Login onLogin={handleLogin} />;

  const isMaster = profile.role === 'master' || user.email === MASTER_EMAIL;
  const isBiller = profile.role === 'biller' || isMaster;

  return (
    <div style={{ minHeight: '100vh', background: '#F7F9FB', fontFamily: "'DM Sans', sans-serif" }}>
      {/* Nav bar */}
      <header style={{ background: 'linear-gradient(90deg, #112640 0%, #1B3A5C 100%)', padding: '0 24px', display: 'flex', alignItems: 'center', gap: 0, height: 52, boxShadow: '0 2px 8px rgba(0,0,0,0.15)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginRight: 32 }}>
          <svg width="28" height="28" viewBox="0 0 48 48" fill="none">
            <circle cx="24" cy="24" r="24" fill="rgba(46,125,140,0.3)" />
            <circle cx="24" cy="24" r="5" fill="#2E7D8C" />
            <circle cx="24" cy="24" r="2" fill="white" />
          </svg>
          <span style={{ fontFamily: "'DM Serif Display',serif", fontSize: 15, color: 'white' }}>The Spark Billing</span>
        </div>

        {/* Nav tabs */}
        <div style={{ display: 'flex', gap: 2, flex: 1 }}>
          {isBiller && (
            <NavBtn active={page === 'main'} onClick={() => setPage('main')}>📋 Billing Sheet</NavBtn>
          )}
          {!isBiller && (
            <NavBtn active={page === 'main'} onClick={() => setPage('main')}>📝 My Log</NavBtn>
          )}
          {isBiller && (
            <NavBtn active={page === 'logs'} onClick={() => setPage('logs')}>👩‍⚕️ Doctor Logs</NavBtn>
          )}
          {isMaster && (
            <NavBtn active={page === 'admin'} onClick={() => setPage('admin')}>⚙ Admin</NavBtn>
          )}
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <span style={{ fontSize: 12, color: 'rgba(255,255,255,0.5)' }}>{profile.name}</span>
          <button onClick={handleSignOut} style={{ background: 'none', border: '1px solid rgba(255,255,255,0.2)', borderRadius: 6, padding: '5px 12px', color: 'rgba(255,255,255,0.6)', cursor: 'pointer', fontSize: 12 }}>Sign out</button>
        </div>
      </header>

      {/* Page content */}
      <main style={{ padding: '20px 0' }}>
        {/* Billing sheet — billers and master */}
        {isBiller && page === 'main' && (
          <BillingSheet user={user} profile={profile} isMaster={isMaster} />
        )}

        {/* Doctor log — doctors only on main, billers on logs tab */}
        {!isBiller && page === 'main' && (
          <DoctorLog user={user} profile={profile} />
        )}

        {/* Doctor logs browser — biller/master only */}
        {isBiller && page === 'logs' && (
          <DoctorLogsView />
        )}

        {/* Admin panel — master only */}
        {isMaster && page === 'admin' && (
          <AdminPanel currentUser={user} />
        )}
      </main>
    </div>
  );
}

function NavBtn({ active, onClick, children }) {
  return (
    <button onClick={onClick} style={{ background: active ? 'rgba(46,125,140,0.3)' : 'transparent', color: active ? 'white' : 'rgba(255,255,255,0.55)', border: 'none', borderBottom: active ? '2px solid #2E7D8C' : '2px solid transparent', padding: '14px 18px', cursor: 'pointer', fontSize: 13, fontWeight: active ? 600 : 400, fontFamily: "'DM Sans',sans-serif", transition: 'all 0.15s' }}>
      {children}
    </button>
  );
}

// ── Doctor Logs Browser (for billers) ──────────────────────────────────────
function DoctorLogsView() {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filterDate, setFilterDate] = useState('');
  const [filterLoc, setFilterLoc] = useState('');
  const [filterDoctor, setFilterDoctor] = useState('');
  const { collection, getDocs, query, orderBy } = require('firebase/firestore');
  const { db } = require('./firebase');
  const { LOCATIONS, LOC_FULL, LOC_COLORS, SERVICE_FIELDS, fmt$, fmtDate } = require('./utils');

  useEffect(() => {
    getDocs(collection(db, 'doctorLogs')).then(snap => {
      setLogs(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      setLoading(false);
    });
  }, []);

  const filtered = logs.filter(l => {
    if (filterDate && l.date !== filterDate) return false;
    if (filterLoc && l.location !== filterLoc) return false;
    if (filterDoctor && l.doctorId !== filterDoctor) return false;
    return true;
  }).sort((a, b) => b.date?.localeCompare(a.date));

  const allDoctors = [...new Set(logs.map(l => l.doctorId))].filter(Boolean).sort();
  const inp = { padding: '7px 10px', border: '1.5px solid #E2E8F0', borderRadius: 7, fontSize: 12, fontFamily: "'DM Sans',sans-serif", outline: 'none', background: 'white' };

  if (loading) return <div style={{ padding: 40, textAlign: 'center', color: '#94A3B8' }}>Loading logs…</div>;

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto', padding: '0 20px' }}>
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ fontFamily: "'DM Serif Display',serif", fontSize: 20, color: '#1B3A5C', marginBottom: 12 }}>Doctor Logs</h2>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <input type="date" value={filterDate} onChange={e => setFilterDate(e.target.value)} style={inp} />
          <select value={filterLoc} onChange={e => setFilterLoc(e.target.value)} style={inp}>
            <option value="">All Locations</option>
            {LOCATIONS.map(l => <option key={l} value={l}>{l} — {LOC_FULL[l]}</option>)}
          </select>
          <select value={filterDoctor} onChange={e => setFilterDoctor(e.target.value)} style={inp}>
            <option value="">All Doctors</option>
            {allDoctors.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          {(filterDate || filterLoc || filterDoctor) && <button onClick={() => { setFilterDate(''); setFilterLoc(''); setFilterDoctor(''); }} style={{ background: 'none', border: 'none', color: '#94A3B8', cursor: 'pointer', fontSize: 12 }}>Clear</button>}
          <span style={{ fontSize: 11, color: '#94A3B8' }}>{filtered.length} entries</span>
        </div>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {filtered.map(l => (
          <div key={l.id} style={{ background: 'white', borderRadius: 8, padding: '10px 14px', border: '1px solid #E2E8F0', display: 'flex', gap: 12, alignItems: 'flex-start' }}>
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 4, flexWrap: 'wrap' }}>
                <span style={{ fontWeight: 700, fontSize: 13, color: '#1B3A5C' }}>{l.patientName}</span>
                <span style={{ fontSize: 11, color: '#64748B', fontWeight: 600 }}>{l.doctorId}</span>
                {l.location && <span style={{ background: LOC_COLORS[l.location] + '25', color: LOC_COLORS[l.location], borderRadius: 4, padding: '1px 6px', fontSize: 10, fontWeight: 700 }}>{l.location}</span>}
                <span style={{ fontSize: 11, color: '#94A3B8' }}>{fmtDate(l.date)}</span>
                {l.payor1 && <span style={{ fontSize: 10, background: '#F1F5F9', borderRadius: 3, padding: '1px 6px', color: '#475569' }}>{l.payor1}</span>}
              </div>
              <div style={{ display: 'flex', gap: 5, flexWrap: 'wrap', fontSize: 11 }}>
                {SERVICE_FIELDS.filter(f => l[f.key] && parseFloat(l[f.key]) > 0).map(f => (
                  <span key={f.key} style={{ background: f.key === 'myopia' ? '#F5F3FF' : '#F8FAFC', color: f.key === 'myopia' ? '#8B5CF6' : '#64748B', borderRadius: 3, padding: '1px 6px' }}>
                    {f.label}: {fmt$(l[f.key])}
                  </span>
                ))}
                {l.cash && <span style={{ background: '#F0FDF4', color: '#16A34A', borderRadius: 3, padding: '1px 6px' }}>Cash: {fmt$(l.cash)}</span>}
                {l.claimNumber && <span style={{ background: '#FFF7ED', color: '#C2410C', borderRadius: 3, padding: '1px 6px' }}>Claim: {l.claimNumber}</span>}
                {l.notes && <span style={{ color: '#94A3B8', fontStyle: 'italic' }}>"{l.notes}"</span>}
              </div>
            </div>
          </div>
        ))}
        {filtered.length === 0 && <div style={{ textAlign: 'center', padding: 32, color: '#94A3B8' }}>No doctor log entries for this filter.</div>}
      </div>
    </div>
  );
}

function LoadingScreen() {
  return (
    <div style={{ minHeight: '100vh', background: 'linear-gradient(135deg, #112640 0%, #1B3A5C 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ textAlign: 'center', color: 'white' }}>
        <div style={{ width: 36, height: 36, border: '3px solid rgba(255,255,255,0.2)', borderTopColor: '#2E7D8C', borderRadius: '50%', animation: 'spin 0.8s linear infinite', margin: '0 auto 14px' }} />
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        <p style={{ opacity: 0.6, fontSize: 13 }}>Loading…</p>
      </div>
    </div>
  );
}
