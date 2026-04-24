import React, { useState, useEffect, useMemo, useRef } from 'react';
import { db } from './firebase';
import { collection, getDocs, setDoc, doc, deleteDoc, query, orderBy } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import { SERVICE_FIELDS, LOCATIONS, LOC_FULL, LOC_COLORS, PAYORS, STATUS_OPTIONS, uid, today, fmtDate, fmtDay, fmt$, emptyEntry, getWeek } from './utils';

const DOCTORS = ['Kha', 'Pan', 'Fan', 'Luong', 'Kaneta', 'Ghag', 'Zhang', 'So', 'Burger', 'Cheng', 'Duong', 'Miranda', 'Pham', 'Yang'];

// ── Biller Entry Form ──────────────────────────────────────────────────────
function BillerEntryForm({ entry, onSave, onCancel, isNew, doctorLogs }) {
  const [form, setForm] = useState({ ...entry });
  const [saving, setSaving] = useState(false);
  const [showRef, setShowRef] = useState(false);

  const set = (key, val) => setForm(f => ({ ...f, [key]: val }));
  const inp = { padding: '6px 8px', border: '1.5px solid #E2E8F0', borderRadius: 6, fontSize: 12, fontFamily: "'DM Sans',sans-serif", outline: 'none', width: '100%', boxSizing: 'border-box' };
  const sel = { ...inp, background: 'white', cursor: 'pointer' };

  // Find matching doctor log entries for reference
  const matchingLogs = useMemo(() => {
    if (!doctorLogs || !form.date) return [];
    return doctorLogs.filter(l =>
      l.date === form.date &&
      (!form.location || l.location === form.location) &&
      (!form.doctorId || l.doctorId === form.doctorId?.toLowerCase())
    );
  }, [doctorLogs, form.date, form.location, form.doctorId]);

  const total = SERVICE_FIELDS.reduce((s, f) => s + (parseFloat(form[f.key]) || 0), 0);

  const save = async () => {
    if (!form.patientName || !form.date) return;
    setSaving(true);
    const updated = { ...form, updatedAt: new Date().toISOString() };
    await setDoc(doc(db, 'billingEntries', form.id), updated);
    onSave(updated);
    setSaving(false);
  };

  return (
    <div style={{ background: isNew ? '#F0F9FF' : '#FFFBEB', border: `1px solid ${isNew ? '#BAE6FD' : '#FDE68A'}`, borderRadius: 10, padding: '16px', marginBottom: 12 }}>
      {/* Row 1: key fields */}
      <div style={{ display: 'grid', gridTemplateColumns: '130px 90px 100px 1fr 1fr 1fr', gap: 8, marginBottom: 10 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Date</span>
          <input type="date" value={form.date} onChange={e => set('date', e.target.value)} style={inp} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Location</span>
          <select value={form.location} onChange={e => set('location', e.target.value)} style={sel}>
            <option value="">—</option>
            {LOCATIONS.map(l => <option key={l} value={l}>{l}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Doctor</span>
          <select value={form.doctorId} onChange={e => set('doctorId', e.target.value)} style={sel}>
            <option value="">—</option>
            {DOCTORS.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Patient Name</span>
          <input value={form.patientName} onChange={e => set('patientName', e.target.value)} style={inp} placeholder="Last, First" />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Status</span>
          <select value={form.status} onChange={e => set('status', e.target.value)} style={{ ...sel, color: form.status === 'flagged' ? '#DC2626' : form.status === 'completed' ? '#059669' : '#64748B' }}>
            {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s.charAt(0).toUpperCase() + s.slice(1)}</option>)}
          </select>
        </label>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Total</span>
          <div style={{ padding: '6px 8px', background: 'white', border: '1px solid #E2E8F0', borderRadius: 6, fontSize: 13, fontWeight: 700, color: '#1B3A5C' }}>{fmt$(total)}</div>
        </div>
      </div>

      {/* Services */}
      <div style={{ marginBottom: 10 }}>
        <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Services & Amounts</span>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(85px, 1fr))', gap: 6, marginTop: 6 }}>
          {SERVICE_FIELDS.map(f => (
            <label key={f.key} style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              <span style={{ fontSize: 9, fontWeight: 600, color: f.key === 'myopia' ? '#8B5CF6' : '#475569', textTransform: 'uppercase' }}>{f.label}</span>
              <input type="number" step="0.01" value={form[f.key]} onChange={e => set(f.key, e.target.value)} style={{ ...inp, textAlign: 'right' }} placeholder="0" />
            </label>
          ))}
        </div>
      </div>

      {/* Insurance & payment */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 90px 90px 90px 90px', gap: 8, marginBottom: 10 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Payor 1</span>
          <select value={form.payor1} onChange={e => set('payor1', e.target.value)} style={sel}>
            <option value="">—</option>
            {PAYORS.map(p => <option key={p}>{p}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Payor 2</span>
          <select value={form.payor2} onChange={e => set('payor2', e.target.value)} style={sel}>
            <option value="">—</option>
            {PAYORS.map(p => <option key={p}>{p}</option>)}
          </select>
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Cash</span>
          <input type="number" step="0.01" value={form.cash} onChange={e => set('cash', e.target.value)} style={{ ...inp, textAlign: 'right' }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#64748B', textTransform: 'uppercase' }}>Ins Copay</span>
          <input type="number" step="0.01" value={form.ins} onChange={e => set('ins', e.target.value)} style={{ ...inp, textAlign: 'right' }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#2E7D8C', textTransform: 'uppercase' }}>Ins Paid 1</span>
          <input type="number" step="0.01" value={form.insurancePaid1} onChange={e => set('insurancePaid1', e.target.value)} style={{ ...inp, textAlign: 'right', borderColor: form.insurancePaid1 ? '#2E7D8C' : '#E2E8F0' }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#2E7D8C', textTransform: 'uppercase' }}>Ins Paid 2</span>
          <input type="number" step="0.01" value={form.insurancePaid2} onChange={e => set('insurancePaid2', e.target.value)} style={{ ...inp, textAlign: 'right', borderColor: form.insurancePaid2 ? '#2E7D8C' : '#E2E8F0' }} />
        </label>
      </div>

      {/* Loss + ATTN */}
      <div style={{ display: 'grid', gridTemplateColumns: '120px 120px 1fr', gap: 8, marginBottom: 12 }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#EF4444', textTransform: 'uppercase' }}>Payment Error Loss</span>
          <input type="number" step="0.01" value={form.paymentErrorLoss} onChange={e => set('paymentErrorLoss', e.target.value)} style={{ ...inp, textAlign: 'right', borderColor: form.paymentErrorLoss ? '#EF4444' : '#E2E8F0' }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#EF4444', textTransform: 'uppercase' }}>Ins Nonpayment Loss</span>
          <input type="number" step="0.01" value={form.insuranceNonpaymentLoss} onChange={e => set('insuranceNonpaymentLoss', e.target.value)} style={{ ...inp, textAlign: 'right', borderColor: form.insuranceNonpaymentLoss ? '#EF4444' : '#E2E8F0' }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 9, fontWeight: 700, color: '#8B5CF6', textTransform: 'uppercase' }}>ATTN (for Christine)</span>
          <input value={form.attn} onChange={e => set('attn', e.target.value)} style={{ ...inp, borderColor: form.attn ? '#8B5CF6' : '#E2E8F0' }} placeholder="Flag for Christine's attention..." />
        </label>
      </div>

      {/* Doctor log reference */}
      {matchingLogs.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          <button onClick={() => setShowRef(v => !v)} style={{ background: '#F5F3FF', border: '1px solid #DDD6FE', borderRadius: 6, padding: '5px 12px', cursor: 'pointer', fontSize: 11, color: '#6D28D9', fontWeight: 600 }}>
            📋 {showRef ? 'Hide' : 'Show'} doctor log reference ({matchingLogs.length} entries)
          </button>
          {showRef && (
            <div style={{ marginTop: 8, background: '#F5F3FF', borderRadius: 8, padding: '10px 14px', border: '1px solid #DDD6FE' }}>
              <p style={{ fontSize: 10, fontWeight: 700, color: '#6D28D9', textTransform: 'uppercase', marginBottom: 8 }}>Doctor Log — {form.date} {form.location}</p>
              {matchingLogs.map(l => (
                <div key={l.id} style={{ fontSize: 11, color: '#374151', marginBottom: 5, padding: '5px 8px', background: 'white', borderRadius: 5 }}>
                  <strong>{l.patientName}</strong> · {l.payor1 || '—'} ·{' '}
                  {SERVICE_FIELDS.filter(f => l[f.key] > 0).map(f => `${f.label}: ${fmt$(l[f.key])}`).join(', ')}
                  {l.notes && <span style={{ color: '#94A3B8' }}> · "{l.notes}"</span>}
                  {l.claimNumber && <span style={{ color: '#C2410C' }}> · Claim: {l.claimNumber}</span>}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      <div style={{ display: 'flex', gap: 8 }}>
        <button onClick={save} disabled={saving || !form.patientName || !form.date}
          style={{ background: saving ? '#94A3B8' : '#1B3A5C', color: 'white', border: 'none', borderRadius: 7, padding: '8px 20px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
          {saving ? 'Saving…' : isNew ? '+ Add Entry' : 'Save Changes'}
        </button>
        {onCancel && <button onClick={onCancel} style={{ background: '#F1F5F9', color: '#64748B', border: 'none', borderRadius: 7, padding: '8px 14px', cursor: 'pointer', fontSize: 13, fontFamily: "'DM Sans',sans-serif" }}>Cancel</button>}
      </div>
    </div>
  );
}

// ── Main Billing Sheet ─────────────────────────────────────────────────────
export default function BillingSheet({ user, profile, isMaster }) {
  const [entries, setEntries] = useState([]);
  const [doctorLogs, setDoctorLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filterDate, setFilterDate] = useState(today());
  const [filterLoc, setFilterLoc] = useState('');
  const [filterDoctor, setFilterDoctor] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterMonth, setFilterMonth] = useState('');
  const [editId, setEditId] = useState(null);
  const [showNew, setShowNew] = useState(false);
  const [searchName, setSearchName] = useState('');

  useEffect(() => {
    async function load() {
      try {
        const [snap, logSnap] = await Promise.all([
          getDocs(query(collection(db, 'billingEntries'), orderBy('date', 'desc'))),
          getDocs(collection(db, 'doctorLogs')),
        ]);
        setEntries(snap.docs.map(d => ({ id: d.id, ...d.data() })));
        setDoctorLogs(logSnap.docs.map(d => ({ id: d.id, ...d.data() })));
      } catch (e) { console.error(e); }
      setLoading(false);
    }
    load();
  }, []);

  const filtered = useMemo(() => {
    return entries.filter(e => {
      if (filterDate && e.date !== filterDate) return false;
      if (filterMonth && !e.date?.startsWith(filterMonth)) return false;
      if (filterLoc && e.location !== filterLoc) return false;
      if (filterDoctor && e.doctorId !== filterDoctor) return false;
      if (filterStatus && e.status !== filterStatus) return false;
      if (searchName && !e.patientName?.toLowerCase().includes(searchName.toLowerCase())) return false;
      return true;
    });
  }, [entries, filterDate, filterMonth, filterLoc, filterDoctor, filterStatus, searchName]);

  const flaggedCount = entries.filter(e => e.status === 'flagged').length;
  const attnCount = entries.filter(e => e.attn).length;

  const handleSave = (entry) => {
    setEntries(prev => {
      const idx = prev.findIndex(e => e.id === entry.id);
      if (idx >= 0) { const n = [...prev]; n[idx] = entry; return n; }
      return [entry, ...prev];
    });
    setShowNew(false);
    setEditId(null);
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this entry?')) return;
    await deleteDoc(doc(db, 'billingEntries', id));
    setEntries(prev => prev.filter(e => e.id !== id));
  };

  const exportExcel = () => {
    // Build rows matching 2026 tab format
    const headers = ['Location','Day','Month','Week','Date','Status','Doctor','Gross','Patient Name','Routine','CL','Optos','DFE','OV','OCT','Topo','Other','Myopia','LASIK','Materials','Paid','Ins','Total','Payor','Payor 2','Cash','Insurance Paid','Payment Error Loss','Insurance Nonpayment Loss','ATTN'];
    const rows = entries.map(e => {
      const d = new Date((e.date || today()) + 'T00:00:00');
      const gross = SERVICE_FIELDS.filter(f => f.key !== 'myopia').reduce((s, f) => s + (parseFloat(e[f.key]) || 0), 0);
      const total = SERVICE_FIELDS.reduce((s, f) => s + (parseFloat(e[f.key]) || 0), 0);
      const insPaid = (parseFloat(e.insurancePaid1) || 0) + (parseFloat(e.insurancePaid2) || 0);
      return [
        e.location, d.toLocaleDateString('en-US',{weekday:'long'}),
        d.getMonth()+1, getWeek(e.date), e.date, e.status, e.doctorId, gross, e.patientName,
        e.routine||'', e.cl||'', e.optos||'', e.dfe||'', e.ov||'', e.oct||'',
        e.topo||'', e.other||'', e.myopia||'', e.lasik||'', e.materials||'',
        e.cash||'', e.ins||'', total, e.payor1||'', e.payor2||'',
        (parseFloat(e.cash)||0)+(parseFloat(e.ins)||0),
        insPaid > 0 ? insPaid : '',
        e.paymentErrorLoss||'', e.insuranceNonpaymentLoss||'', e.attn||''
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '2026');
    XLSX.writeFile(wb, `spark_billing_${today()}.xlsx`);
  };

  const inp = { padding: '7px 10px', border: '1.5px solid #E2E8F0', borderRadius: 7, fontSize: 12, fontFamily: "'DM Sans',sans-serif", outline: 'none', background: 'white' };

  if (loading) return <div style={{ padding: 40, textAlign: 'center', color: '#94A3B8' }}>Loading billing sheet…</div>;

  return (
    <div style={{ maxWidth: 1100, margin: '0 auto', padding: '24px 20px' }}>
      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 16, flexWrap: 'wrap', gap: 12 }}>
        <div>
          <h1 style={{ fontFamily: "'DM Serif Display',serif", fontSize: 24, color: '#1B3A5C' }}>Billing Sheet</h1>
          <p style={{ color: '#94A3B8', fontSize: 12, marginTop: 2 }}>{entries.length} total entries</p>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {flaggedCount > 0 && (
            <div style={{ background: '#FEF3C7', border: '1px solid #FDE68A', borderRadius: 7, padding: '7px 12px', fontSize: 12, color: '#92400E', fontWeight: 600 }}>
              ⚠ {flaggedCount} flagged
            </div>
          )}
          {attnCount > 0 && (
            <div style={{ background: '#F5F3FF', border: '1px solid #DDD6FE', borderRadius: 7, padding: '7px 12px', fontSize: 12, color: '#6D28D9', fontWeight: 600 }}>
              📋 {attnCount} need attention
            </div>
          )}
          <button onClick={exportExcel} style={{ background: '#10B981', color: 'white', border: 'none', borderRadius: 7, padding: '8px 16px', cursor: 'pointer', fontSize: 12, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
            📥 Export Excel
          </button>
          <button onClick={() => setShowNew(true)} style={{ background: '#1B3A5C', color: 'white', border: 'none', borderRadius: 8, padding: '9px 16px', cursor: 'pointer', fontSize: 13, fontWeight: 600, fontFamily: "'DM Sans',sans-serif" }}>
            + Add Entry
          </button>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center', background: 'white', padding: '12px 16px', borderRadius: 10, border: '1px solid #E2E8F0' }}>
        <input placeholder="Search patient..." value={searchName} onChange={e => { setSearchName(e.target.value); setFilterDate(''); setFilterMonth(''); }} style={{ ...inp, width: 180 }} />
        <input type="date" value={filterDate} onChange={e => { setFilterDate(e.target.value); setFilterMonth(''); setSearchName(''); }} style={inp} />
        <input type="month" value={filterMonth} onChange={e => { setFilterMonth(e.target.value); setFilterDate(''); setSearchName(''); }} style={inp} />
        <select value={filterLoc} onChange={e => setFilterLoc(e.target.value)} style={inp}>
          <option value="">All Locations</option>
          {LOCATIONS.map(l => <option key={l} value={l}>{l} — {LOC_FULL[l]}</option>)}
        </select>
        <select value={filterDoctor} onChange={e => setFilterDoctor(e.target.value)} style={inp}>
          <option value="">All Doctors</option>
          {DOCTORS.map(d => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={inp}>
          <option value="">All Status</option>
          {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s.charAt(0).toUpperCase() + s.slice(1)}</option>)}
        </select>
        {(filterDate || filterMonth || filterLoc || filterDoctor || filterStatus || searchName) && (
          <button onClick={() => { setFilterDate(today()); setFilterMonth(''); setFilterLoc(''); setFilterDoctor(''); setFilterStatus(''); setSearchName(''); }}
            style={{ background: 'none', border: 'none', color: '#94A3B8', cursor: 'pointer', fontSize: 12 }}>Clear</button>
        )}
        <span style={{ fontSize: 11, color: '#94A3B8', marginLeft: 4 }}>{filtered.length} entries</span>
      </div>

      {/* New entry form */}
      {showNew && (
        <BillerEntryForm
          entry={emptyEntry('', '')}
          onSave={handleSave}
          onCancel={() => setShowNew(false)}
          isNew
          doctorLogs={doctorLogs}
        />
      )}

      {/* Entries table */}
      {filtered.length === 0 && !showNew ? (
        <div style={{ textAlign: 'center', padding: 40, color: '#94A3B8' }}>
          <p>No entries for this filter.</p>
        </div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ borderBottom: '2px solid #F0F4F8', background: '#F8FAFC' }}>
                {['Date','Loc','Dr.','Patient','Services','Payor','Cash','Ins','Ins Paid','Status','ATTN',''].map(h => (
                  <th key={h} style={{ padding: '8px 8px', textAlign: 'left', fontSize: 9, fontWeight: 700, color: '#94A3B8', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((e, i) => {
                const gross = SERVICE_FIELDS.filter(f => f.key !== 'myopia').reduce((s, f) => s + (parseFloat(e[f.key]) || 0), 0);
                const insPaid = (parseFloat(e.insurancePaid1)||0) + (parseFloat(e.insurancePaid2)||0);
                return editId === e.id ? (
                  <tr key={e.id}>
                    <td colSpan={12} style={{ padding: '8px 0' }}>
                      <BillerEntryForm entry={e} onSave={handleSave} onCancel={() => setEditId(null)} isNew={false} doctorLogs={doctorLogs} />
                    </td>
                  </tr>
                ) : (
                  <tr key={e.id} style={{ background: e.status === 'flagged' ? '#FFF7ED' : e.attn ? '#F5F3FF' : i % 2 === 0 ? 'white' : '#FAFBFC', borderBottom: '1px solid #F0F4F8' }}>
                    <td style={{ padding: '7px 8px', whiteSpace: 'nowrap', color: '#64748B' }}>{fmtDate(e.date)}</td>
                    <td style={{ padding: '7px 8px' }}>
                      {e.location && <span style={{ background: LOC_COLORS[e.location] + '25', color: LOC_COLORS[e.location], borderRadius: 4, padding: '1px 6px', fontSize: 10, fontWeight: 700 }}>{e.location}</span>}
                    </td>
                    <td style={{ padding: '7px 8px', color: '#64748B' }}>{e.doctorId}</td>
                    <td style={{ padding: '7px 8px', fontWeight: 600, maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{e.patientName}</td>
                    <td style={{ padding: '7px 8px' }}>
                      <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                        {SERVICE_FIELDS.filter(f => e[f.key] && parseFloat(e[f.key]) > 0).map(f => (
                          <span key={f.key} style={{ background: f.key === 'myopia' ? '#F5F3FF' : '#F1F5F9', color: f.key === 'myopia' ? '#8B5CF6' : '#475569', borderRadius: 3, padding: '1px 5px', fontSize: 10 }}>
                            {f.label} {fmt$(e[f.key])}
                          </span>
                        ))}
                        {gross > 0 && <span style={{ background: '#F0FDF4', color: '#16A34A', borderRadius: 3, padding: '1px 5px', fontSize: 10, fontWeight: 700 }}>{fmt$(gross)}</span>}
                      </div>
                    </td>
                    <td style={{ padding: '7px 8px', fontSize: 11, color: '#64748B' }}>
                      {e.payor1 && <span style={{ display: 'block' }}>{e.payor1}</span>}
                      {e.payor2 && <span style={{ display: 'block', color: '#94A3B8' }}>{e.payor2}</span>}
                    </td>
                    <td style={{ padding: '7px 8px', textAlign: 'right', color: '#16A34A' }}>{e.cash ? fmt$(e.cash) : '—'}</td>
                    <td style={{ padding: '7px 8px', textAlign: 'right', color: '#1D4ED8' }}>{e.ins ? fmt$(e.ins) : '—'}</td>
                    <td style={{ padding: '7px 8px', textAlign: 'right', color: insPaid > 0 ? '#059669' : '#94A3B8' }}>
                      {insPaid > 0 ? fmt$(insPaid) : <span style={{ fontSize: 10 }}>pending</span>}
                    </td>
                    <td style={{ padding: '7px 8px' }}>
                      <span style={{ fontSize: 10, fontWeight: 600, borderRadius: 4, padding: '2px 7px', background: e.status === 'completed' ? '#D1FAE5' : e.status === 'flagged' ? '#FEF3C7' : '#F1F5F9', color: e.status === 'completed' ? '#065F46' : e.status === 'flagged' ? '#92400E' : '#64748B' }}>
                        {e.status}
                      </span>
                    </td>
                    <td style={{ padding: '7px 8px', maxWidth: 140, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {e.attn && <span style={{ fontSize: 11, color: '#6D28D9' }} title={e.attn}>📋 {e.attn.slice(0, 30)}{e.attn.length > 30 ? '…' : ''}</span>}
                      {(e.paymentErrorLoss > 0 || e.insuranceNonpaymentLoss > 0) && (
                        <span style={{ fontSize: 10, color: '#DC2626', display: 'block' }}>⚠ Loss: {fmt$((parseFloat(e.paymentErrorLoss)||0)+(parseFloat(e.insuranceNonpaymentLoss)||0))}</span>
                      )}
                    </td>
                    <td style={{ padding: '7px 8px' }}>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button onClick={() => setEditId(e.id)} style={{ background: '#EFF6FF', border: 'none', borderRadius: 4, padding: '4px 8px', cursor: 'pointer', color: '#1D4ED8', fontSize: 10 }}>Edit</button>
                        <button onClick={() => handleDelete(e.id)} style={{ background: '#FEE2E2', border: 'none', borderRadius: 4, padding: '4px 6px', cursor: 'pointer', color: '#DC2626', fontSize: 10 }}>✕</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
