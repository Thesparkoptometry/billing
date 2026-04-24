import React, { useState, useEffect, useMemo } from 'react';
import { db } from './firebase';
import { collection, getDocs, setDoc, doc, deleteDoc, query, where, orderBy } from 'firebase/firestore';
import { SERVICE_FIELDS, LOCATIONS, LOC_FULL, LOC_COLORS, PAYORS, uid, today, fmtDate, fmtDay, fmt$, emptyEntry } from './utils';
import { changePassword } from './auth';

// ── Entry Row Form ─────────────────────────────────────────────────────────
function EntryForm({ entry, onSave, onCancel, isNew }) {
  const [form, setForm] = useState({ ...entry });
  const [saving, setSaving] = useState(false);

  const set = (key, val) => setForm(f => ({ ...f, [key]: val }));
  const inp = { padding: '6px 8px', border: '1.5px solid #E2E8F0', borderRadius: 6, fontSize: 12, fontFamily: "'DM Sans',sans-serif", outline: 'none', width: '100%', boxSizing: 'border-box' };
  const sel = { ...inp, background: 'white', cursor: 'pointer' };

  const total = SERVICE_FIELDS.reduce((s, f) => s + (parseFloat(form[f.key]) || 0), 0);
  const nonMyopiaTotal = SERVICE_FIELDS.filter(f => f.key !== 'myopia').reduce((s, f) => s + (parseFloat(form[f.key]) || 0), 0);

  const save = async () => {
    if (!form.patientName || !form.date || !form.location) return;
    setSaving(true);
    const updated = { ...form, updatedAt: new Date().toISOString() };
    await setDoc(doc(db, 'doctorLogs', form.id), updated);
    onSave(updated);
    setSaving(false);
  };

  return (
    <div style={{ background: isNew ? '#F0F9FF' : '#FFFBEB', border: `1px solid ${isNew ? '#BAE6FD' : '#FDE68A'}`, borderRadius: 10, padding: '14px 16px', marginBottom: 10 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '140px 100px 1fr 120px', gap: 8, marginBottom: 10 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Date</span>
          <input type="date" value={form.date} onChange={e => set('date', e.target.value)} style={inp} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Location</span>
          <select value={form.location} onChange={e => set('location', e.target.value)} style={sel}>
            <option value="">—</option>
            {LOCATIONS.map(l => <option key={l} value={l}>{l} — {LOC_FULL[l]}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Patient Name</span>
          <input value={form.patientName} onChange={e => set('patientName', e.target.value)} style={inp} placeholder="Last, First" />
        </label>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Visit Total</span>
          <div style={{ padding: '6px 8px', background: 'white', border: '1px solid #E2E8F0', borderRadius: 6, fontSize: 13, fontWeight: 700, color: '#1B3A5C' }}>
            {fmt$(total)} {form.myopia ? <span style={{ fontSize: 10, color: '#8B5CF6', fontWeight: 400 }}>(+{fmt$(form.myopia)} myopia)</span> : null}
          </div>
        </div>
      </div>

      {/* Services */}
      <div style={{ marginBottom: 10 }}>
        <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Services & Amounts</span>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(90px, 1fr))', gap: 6, marginTop: 6 }}>
          {SERVICE_FIELDS.map(f => (
            <label key={f.key} style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <span style={{ fontSize: 9, fontWeight: 600, color: f.key === 'myopia' ? '#8B5CF6' : '#475569', textTransform: 'uppercase' }} title={f.hint}>{f.label} {f.key === 'myopia' ? '⚡' : ''}</span>
              <input type="number" step="0.01" value={form[f.key]} onChange={e => set(f.key, e.target.value)} style={{ ...inp, textAlign: 'right', fontSize: 12 }} placeholder="0" />
            </label>
          ))}
        </div>
        {form.myopia > 0 && <p style={{ fontSize: 10, color: '#8B5CF6', marginTop: 4 }}>⚡ Myopia amount excluded from avg/patient calculations</p>}
      </div>

      {/* Insurance & payment */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 100px 100px', gap: 8, marginBottom: 10 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Payor 1</span>
          <select value={form.payor1} onChange={e => set('payor1', e.target.value)} style={sel}>
            <option value="">—</option>
            {PAYORS.map(p => <option key={p}>{p}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Payor 2</span>
          <select value={form.payor2} onChange={e => set('payor2', e.target.value)} style={sel}>
            <option value="">—</option>
            {PAYORS.map(p => <option key={p}>{p}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Cash Paid</span>
          <input type="number" step="0.01" value={form.cash} onChange={e => set('cash', e.target.value)} style={{ ...inp, textAlign: 'right' }} placeholder="0" />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Ins Copay</span>
          <input type="number" step="0.01" value={form.ins} onChange={e => set('ins', e.target.value)} style={{ ...inp, textAlign: 'right' }} placeholder="0" />
        </label>
      </div>

      {/* Claim + Notes */}
      <div style={{ display: 'grid', gridTemplateColumns: '180px 1fr', gap: 8, marginBottom: 12 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Claim # <span style={{ fontWeight: 400, textTransform: 'none' }}>(optional)</span></span>
          <input value={form.claimNumber} onChange={e => set('claimNumber', e.target.value)} style={inp} placeholder="e.g. 12345678" />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Notes</span>
          <input value={form.notes} onChange={e => set('notes', e.target.value)} style={inp} placeholder="Any notes about this visit..." />
        </label>
      </div>

      <div style={{ display: 'flex', gap: 8 }}>
        <button onClick={save} disabled={saving || !form.patientName || !form.location}
          style={{ background: saving ? '#94A3B8' : '#1B3A5C', color: 'white', border: 'none', borderRadius: 7, padding: '8px 20px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
          {saving ? 'Saving…' : isNew ? '+ Add Patient' : 'Save Changes'}
        </button>
        {onCancel && <button onClick={onCancel} style={{ background: '#F1F5F9', color: '#64748B', border: 'none', borderRadius: 7, padding: '8px 14px', cursor: 'pointer', fontSize: 13, fontFamily: "'DM Sans',sans-serif" }}>Cancel</button>}
      </div>
    </div>
  );
}

// ── Doctor Log View ────────────────────────────────────────────────────────
export default function DoctorLog({ user, profile }) {
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filterDate, setFilterDate] = useState(today());
  const [filterLoc, setFilterLoc] = useState('');
  const [editId, setEditId] = useState(null);
  const [showNew, setShowNew] = useState(false);
  const [showChangePw, setShowChangePw] = useState(false);
  const [newPw, setNewPw] = useState('');
  const [pwMsg, setPwMsg] = useState('');

  useEffect(() => {
    async function load() {
      try {
        const q = query(collection(db, 'doctorLogs'), where('doctorId', '==', profile.doctorId || profile.uid));
        const snap = await getDocs(q);
        setEntries(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      } catch (e) { console.error(e); }
      setLoading(false);
    }
    load();
  }, [profile]);

  const filtered = useMemo(() => {
    return entries.filter(e => {
      if (filterDate && e.date !== filterDate) return false;
      if (filterLoc && e.location !== filterLoc) return false;
      return true;
    }).sort((a, b) => a.patientName.localeCompare(b.patientName));
  }, [entries, filterDate, filterLoc]);

  const handleSave = (entry) => {
    setEntries(prev => {
      const idx = prev.findIndex(e => e.id === entry.id);
      if (idx >= 0) { const n = [...prev]; n[idx] = entry; return n; }
      return [...prev, entry];
    });
    setShowNew(false);
    setEditId(null);
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this entry?')) return;
    await deleteDoc(doc(db, 'doctorLogs', id));
    setEntries(prev => prev.filter(e => e.id !== id));
  };

  const handleChangePw = async () => {
    if (!newPw || newPw.length < 6) { setPwMsg('Password must be at least 6 characters.'); return; }
    try {
      await changePassword(newPw);
      setPwMsg('Password changed successfully!');
      setNewPw('');
      setTimeout(() => { setPwMsg(''); setShowChangePw(false); }, 2000);
    } catch (e) { setPwMsg('Error: ' + e.message); }
  };

  const inp = { padding: '7px 10px', border: '1.5px solid #E2E8F0', borderRadius: 7, fontSize: 13, fontFamily: "'DM Sans',sans-serif", outline: 'none', background: 'white' };

  if (loading) return <div style={{ padding: 40, textAlign: 'center', color: '#94A3B8' }}>Loading your log…</div>;

  return (
    <div style={{ maxWidth: 960, margin: '0 auto', padding: '24px 20px' }}>
      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20 }}>
        <div>
          <h1 style={{ fontFamily: "'DM Serif Display',serif", fontSize: 26, color: '#1B3A5C' }}>My Patient Log</h1>
          <p style={{ color: '#94A3B8', fontSize: 13, marginTop: 2 }}>{profile.name} · {entries.length} total entries</p>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button onClick={() => setShowChangePw(v => !v)} style={{ background: '#F1F5F9', color: '#64748B', border: 'none', borderRadius: 7, padding: '8px 14px', cursor: 'pointer', fontSize: 12, fontFamily: "'DM Sans',sans-serif" }}>
            🔑 Change Password
          </button>
          <button onClick={() => setShowNew(true)} style={{ background: '#1B3A5C', color: 'white', border: 'none', borderRadius: 8, padding: '9px 18px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
            + Add Patient
          </button>
        </div>
      </div>

      {/* Change password */}
      {showChangePw && (
        <div style={{ background: 'white', borderRadius: 10, padding: '14px 18px', marginBottom: 16, border: '1px solid #E2E8F0', display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <input type="password" value={newPw} onChange={e => setNewPw(e.target.value)} placeholder="New password (min 6 chars)" style={{ ...inp, width: 240 }} />
          <button onClick={handleChangePw} style={{ background: '#10B981', color: 'white', border: 'none', borderRadius: 7, padding: '8px 16px', cursor: 'pointer', fontSize: 13, fontWeight: 600 }}>Update Password</button>
          <button onClick={() => setShowChangePw(false)} style={{ background: '#F1F5F9', color: '#64748B', border: 'none', borderRadius: 7, padding: '8px 12px', cursor: 'pointer', fontSize: 12 }}>Cancel</button>
          {pwMsg && <span style={{ fontSize: 12, color: pwMsg.includes('Error') ? '#DC2626' : '#10B981', fontWeight: 600 }}>{pwMsg}</span>}
        </div>
      )}

      {/* Filters */}
      <div style={{ display: 'flex', gap: 10, marginBottom: 16, alignItems: 'center', flexWrap: 'wrap' }}>
        <input type="date" value={filterDate} onChange={e => setFilterDate(e.target.value)} style={inp} />
        <select value={filterLoc} onChange={e => setFilterLoc(e.target.value)} style={inp}>
          <option value="">All Locations</option>
          {LOCATIONS.map(l => <option key={l} value={l}>{l} — {LOC_FULL[l]}</option>)}
        </select>
        {(filterDate || filterLoc) && <button onClick={() => { setFilterDate(''); setFilterLoc(''); }} style={{ background: 'none', border: 'none', color: '#94A3B8', cursor: 'pointer', fontSize: 12 }}>Clear filters</button>}
        <span style={{ fontSize: 12, color: '#94A3B8' }}>{filtered.length} patient{filtered.length !== 1 ? 's' : ''} shown</span>
      </div>

      {/* New entry form */}
      {showNew && (
        <EntryForm
          entry={emptyEntry(profile.doctorId || profile.uid, profile.name)}
          onSave={handleSave}
          onCancel={() => setShowNew(false)}
          isNew
        />
      )}

      {/* Entries */}
      {filtered.length === 0 && !showNew ? (
        <div style={{ textAlign: 'center', padding: '40px', color: '#94A3B8' }}>
          <p style={{ fontSize: 15 }}>No entries for this filter.</p>
          <p style={{ fontSize: 13, marginTop: 4 }}>Click "+ Add Patient" to log a visit.</p>
        </div>
      ) : (
        <div>
          {filtered.map(entry => (
            editId === entry.id ? (
              <EntryForm key={entry.id} entry={entry} onSave={handleSave} onCancel={() => setEditId(null)} isNew={false} />
            ) : (
              <div key={entry.id} style={{ background: 'white', borderRadius: 10, padding: '12px 16px', marginBottom: 8, border: '1px solid #E2E8F0', display: 'flex', gap: 14, alignItems: 'flex-start' }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 4, flexWrap: 'wrap' }}>
                    <span style={{ fontWeight: 700, fontSize: 14, color: '#1B3A5C' }}>{entry.patientName}</span>
                    {entry.location && <span style={{ background: LOC_COLORS[entry.location] + '20', color: LOC_COLORS[entry.location], borderRadius: 4, padding: '1px 7px', fontSize: 11, fontWeight: 700 }}>{entry.location}</span>}
                    <span style={{ fontSize: 11, color: '#94A3B8' }}>{fmtDate(entry.date)}</span>
                    {entry.payor1 && <span style={{ fontSize: 11, color: '#64748B', background: '#F1F5F9', borderRadius: 4, padding: '1px 6px' }}>{entry.payor1}</span>}
                    {entry.payor2 && <span style={{ fontSize: 11, color: '#64748B', background: '#F1F5F9', borderRadius: 4, padding: '1px 6px' }}>{entry.payor2}</span>}
                  </div>
                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', fontSize: 11, color: '#64748B' }}>
                    {SERVICE_FIELDS.filter(f => entry[f.key] && parseFloat(entry[f.key]) > 0).map(f => (
                      <span key={f.key} style={{ background: f.key === 'myopia' ? '#F5F3FF' : '#F8FAFC', borderRadius: 4, padding: '2px 7px', color: f.key === 'myopia' ? '#8B5CF6' : '#475569' }}>
                        {f.label}: {fmt$(entry[f.key])}
                      </span>
                    ))}
                    {entry.cash && <span style={{ background: '#F0FDF4', borderRadius: 4, padding: '2px 7px', color: '#16A34A' }}>Cash: {fmt$(entry.cash)}</span>}
                    {entry.ins && <span style={{ background: '#EFF6FF', borderRadius: 4, padding: '2px 7px', color: '#1D4ED8' }}>Ins copay: {fmt$(entry.ins)}</span>}
                    {entry.claimNumber && <span style={{ background: '#FFF7ED', borderRadius: 4, padding: '2px 7px', color: '#C2410C' }}>Claim: {entry.claimNumber}</span>}
                    {entry.notes && <span style={{ color: '#94A3B8', fontStyle: 'italic' }}>"{entry.notes}"</span>}
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                  <div style={{ textAlign: 'right', marginRight: 8 }}>
                    <p style={{ fontSize: 13, fontWeight: 700, color: '#1B3A5C' }}>
                      {fmt$(SERVICE_FIELDS.reduce((s, f) => s + (parseFloat(entry[f.key]) || 0), 0))}
                    </p>
                    <p style={{ fontSize: 10, color: '#94A3B8' }}>total</p>
                  </div>
                  <button onClick={() => setEditId(entry.id)} style={{ background: '#EFF6FF', border: 'none', borderRadius: 5, padding: '5px 10px', cursor: 'pointer', color: '#1D4ED8', fontSize: 11 }}>Edit</button>
                  <button onClick={() => handleDelete(entry.id)} style={{ background: '#FEE2E2', border: 'none', borderRadius: 5, padding: '5px 8px', cursor: 'pointer', color: '#DC2626', fontSize: 11 }}>✕</button>
                </div>
              </div>
            )
          ))}
        </div>
      )}
    </div>
  );
}
