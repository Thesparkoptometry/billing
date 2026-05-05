import React, { useState, useEffect, useRef } from 'react';
import { db, auth } from './firebase';
import { createUserWithEmailAndPassword, sendPasswordResetEmail } from 'firebase/auth';
import { doc, setDoc, getDoc, writeBatch, collection } from 'firebase/firestore';
import { loadAllUsers, saveUserProfile, deleteUser } from './auth';
import { LOCATIONS } from './utils';
import * as XLSX from 'xlsx';

const DEFAULT_FEES = { SC:89, F:89, SV:89, WC:95 };
const DEFAULT_OTHER_ITEMS = ['Contact Lens Training','No-Show Fee','Oasis Materials','Materials','Topo','Axial Length','Pachymetry','Refraction Only','Contact Lens Materials','LASIK','Glaucoma Check','MNCL Reimbursement','Other'];
const DEFAULT_DISCOUNT_TYPES = ['Discount Plan','Employee Discount','Friends & Family','AAA','AARP','Professional Courtesy','Hardship','Aetna Discount Plan','Target Team Member','Insurance Adjustment','Other'];
const DEFAULT_DOCTORS = ['Kha','Pan','Fan','Kaneta','Yang','Ghag','Zhang','So','Burger','Duong','Cheng','Luong'];

const ROLES = [
  { value:'doctor', label:'Doctor — sees only their own log' },
  { value:'biller', label:'Biller — sees billing sheet + all logs' },
  { value:'master', label:'Master — full access (admin)' },
];

// ── Payor normalization (matches billing app) ────────────────────────────────
const PAYOR_MAP = {
  'EM':'EM','EYEMED':'EM','AETNA':'EM','CIGNA':'EM','CIG':'EM',
  'VSP':'VSP','VPS':'VSP','METLIFE':'VSP','MET LIFE':'VSP','GUARDIAN':'VSP',
  'UHC':'UHC','UNITED':'UHC','SPECTERA':'UHC','SPEC':'UHC','SPECTRA':'UHC',
  'SUPERIOR':'SUP','SUP':'SUP','SUOERIOR':'SUP','SUPERIO':'SUP','MES':'SUP',
  'DAVIS':'DV','DAVIS VISION':'DV','DV':'DV',
  'COLONIAL':'COL','COLORIAL':'COL','COLONAIL':'COL','COL':'COL',
  'SELF':'Self','CASH':'Self','OOP':'Self','SELF PAY':'Self',
  'FEP':'FEP','NVA':'NVA','VBA':'VBA','AVESIS':'EM','HERITAGE':'EM',
  'PREMERA':'UHC','PREMIER':'UHC','WELL CARE':'UHC',
  'SPECTERA':'UHC','ALWAYSVISION':'VSP','ALWAYS VISION':'VSP',
};
function normPayor(p) {
  if (!p) return '';
  const up = String(p).trim().toUpperCase();
  if (PAYOR_MAP[up]) return PAYOR_MAP[up];
  // partial
  for (const [k,v] of Object.entries(PAYOR_MAP)) {
    if (up.startsWith(k) && k.length >= 3) return v;
  }
  return String(p).trim();
}

// ── Excel row parser ────────────────────────────────────────────────────────
function parseSheetRows(ws, sheetYear) {
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  // Find header row
  let hi = rows.findIndex(r => r && r[0] === 'Location' && r[6] === 'Doctor');
  if (hi === -1) hi = 0;
  const headers = rows[hi].map(h => h ? String(h).trim() : '');
  const c = name => headers.indexOf(name);

  // 2024 has patient in col 9, others in col 8
  const iPatient = sheetYear === 2024 ? 9 : c('Patient Name') !== -1 ? c('Patient Name') : 8;
  const iLoc=0, iDate=4, iDoctor=6;
  const iRoutine = sheetYear === 2024 ? 10 : c('Routine');
  const iCL      = sheetYear === 2024 ? 11 : c('CL');
  const iOptos   = sheetYear === 2024 ? 12 : c('Optos');
  const iDFE     = sheetYear === 2024 ? 13 : c('DFE');
  const iOV      = sheetYear === 2024 ? 14 : c('OV');
  const iOCT     = sheetYear === 2024 ? 15 : c('OCT');
  const iTopo    = sheetYear === 2024 ? 16 : c('Topo');
  const iOther   = sheetYear === 2024 ? 17 : c('Other');
  const iMyopia  = sheetYear === 2024 ? 18 : c('Myopia');
  const iLASIK   = sheetYear === 2024 ? 19 : c('LASIK');
  const iMat     = c('Materials') !== -1 ? c('Materials') : -1;
  const iPaid    = sheetYear === 2024 ? 20 : c('Paid');
  const iIns     = sheetYear === 2024 ? 21 : c('Ins');
  const iPayor   = sheetYear === 2024 ? 23 : c('Payor');
  const iPayor2  = sheetYear === 2024 ? 24 : c('Payor 2');
  const iPayErr  = sheetYear === 2024 ? 27 : c('Payment Error Loss');
  const iInsNon  = sheetYear === 2024 ? 28 : c('Insurance Nonpayment Loss');
  const iInsPaid = sheetYear === 2024 ? 26 : c('Insurance Paid');
  const iATTN    = sheetYear === 2024 ? 29 : c('ATTN');

  const entries = [];
  const doctors = new Set();

  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const loc = r[iLoc] ? String(r[iLoc]).trim() : null;
    if (!loc || !['SC','F','WC','SV'].includes(loc)) continue;
    const doctor = r[iDoctor] ? String(r[iDoctor]).trim() : null;
    if (!doctor || doctor === 'CLOSED') continue;
    const patient = r[iPatient] ? String(r[iPatient]).trim() : null;
    if (!patient || patient === 'Patient Name') continue;
    // skip month header rows
    if (patient.toUpperCase() === patient && patient.length < 15 && !patient.includes(',')) continue;

    let dateStr = null;
    const dateVal = r[iDate];
    if (dateVal instanceof Date) {
      dateStr = dateVal.toISOString().slice(0,10);
    } else if (typeof dateVal === 'number') {
      try {
        const d = XLSX.SSF.parse_date_code(dateVal);
        if (d) dateStr = `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
      } catch {}
    } else if (typeof dateVal === 'string' && dateVal) {
      const parsed = new Date(dateVal);
      if (!isNaN(parsed)) dateStr = parsed.toISOString().slice(0,10);
    }
    if (!dateStr || !dateStr.startsWith('20')) continue;

    doctors.add(doctor);

    // Service flags → store as dollar amounts where we have fee data, else as '1' if rendered
    const myopiaAmt = parseFloat(r[iMyopia]) || 0;
    const lasikAmt  = parseFloat(r[iLASIK])  || 0;
    const matsAmt   = iMat >= 0 ? parseFloat(r[iMat]) || 0 : 0;

    // For boolean service columns (1 = rendered), store as '1' so billing app shows them
    const sv = (col) => {
      if (col < 0) return '';
      const v = r[col];
      if (v === 1 || v === '1' || v === true) return '1';
      if (parseFloat(v) > 0) return String(parseFloat(v));
      return '';
    };

    const entry = {
      date: dateStr,
      location: loc,
      doctorId: doctor,
      patientName: patient,
      exam:   sv(iRoutine),
      cl:     sv(iCL),
      optos:  sv(iOptos),
      dfe:    sv(iDFE),
      ov:     sv(iOV),
      oct:    sv(iOCT),
      topo:   sv(iTopo),
      other:  sv(iOther),
      myopia: myopiaAmt > 0 ? String(myopiaAmt) : '',
      lasik:  lasikAmt  > 0 ? String(lasikAmt)  : '',
      materials: matsAmt > 0 ? String(matsAmt) : '',
      ptPaid: String(parseFloat(r[iPaid]) || 0) ,
      ins:    String(parseFloat(r[iIns])  || 0),
      payor1: normPayor(r[iPayor]),
      payor2: normPayor(r[iPayor2]),
      insurancePaid1: r[iInsPaid] ? String(parseFloat(r[iInsPaid]) || 0) : '',
      paymentErrorLoss: r[iPayErr] ? String(parseFloat(r[iPayErr]) || 0) : '',
      insuranceNonpaymentLoss: r[iInsNon] ? String(parseFloat(r[iInsNon]) || 0) : '',
      attn: r[iATTN] ? String(r[iATTN]).trim() : '',
      insPaidState: r[iInsPaid] && parseFloat(r[iInsPaid]) > 0 ? 'confirmed' : 'pending',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      _imported: true,
    };
    entries.push(entry);
  }

  return { entries, doctors: [...doctors] };
}

// Generate a stable ID from date+location+patient (for dedup)
function makeEntryId(e) {
  const safe = (s) => String(s||'').replace(/[^a-zA-Z0-9]/g,'').toLowerCase().slice(0,20);
  return `${e.date}_${safe(e.location)}_${safe(e.patientName)}_${safe(e.doctorId)}`;
}

// ── Import Tab ────────────────────────────────────────────────────────────────
function ImportTab() {
  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [importing, setImporting] = useState(false);
  const [progress, setProgress] = useState('');
  const [done, setDone] = useState(null);
  const [error, setError] = useState('');
  const fileRef = useRef();

  const handleFile = async (f) => {
    if (!f) return;
    setFile(f);
    setPreview(null);
    setDone(null);
    setError('');
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type:'array', cellDates:true });
      const results = {};
      let totalEntries = 0;
      const allDoctors = new Set();
      for (const year of [2024, 2025, 2026]) {
        const sheetName = String(year);
        if (!wb.SheetNames.includes(sheetName)) continue;
        const { entries, doctors } = parseSheetRows(wb.Sheets[sheetName], year);
        results[year] = entries;
        totalEntries += entries.length;
        doctors.forEach(d => allDoctors.add(d));
      }
      setPreview({ results, totalEntries, doctors: [...allDoctors].sort(), wb });
    } catch(e) {
      setError('Could not read file: ' + e.message);
    }
  };

  const runImport = async () => {
    if (!preview) return;
    setImporting(true);
    setProgress('Starting...');
    setError('');
    let written = 0, skipped = 0;
    try {
      for (const [year, entries] of Object.entries(preview.results)) {
        setProgress(`Importing ${year}... (${entries.length} rows)`);
        // Write in batches of 400 (Firestore limit is 500)
        const BATCH_SIZE = 100;
        for (let i = 0; i < entries.length; i += BATCH_SIZE) {
          const batch = writeBatch(db);
          const chunk = entries.slice(i, i + BATCH_SIZE);
          for (const entry of chunk) {
            const id = makeEntryId(entry);
            const ref = doc(collection(db, 'billingEntries'), id);
            batch.set(ref, entry); // overwrites if exists
            written++;
          }
          await batch.commit();
          setProgress(`Importing ${year}... ${Math.min(i + BATCH_SIZE, entries.length)} / ${entries.length}`);
          await new Promise(r => setTimeout(r, 300));
        }
      }
      setDone({ written, skipped });
      setProgress('');
    } catch(e) {
      setError('Import failed: ' + e.message);
    }
    setImporting(false);
  };

  const inp = { padding:'7px 10px', border:'1.5px solid #E2E8F0', borderRadius:7, fontSize:13, outline:'none', width:'100%', boxSizing:'border-box' };

  return (
    <div style={{background:'white',borderRadius:12,padding:20,border:'1px solid #E2E8F0'}}>
      <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:6}}>Import from Excel</h3>
      <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>
        Upload your production spreadsheet to import 2024, 2025, and 2026 data into the billing app.
        Existing entries with the same patient + date + location will be overwritten.
      </p>

      {error&&<div style={{background:'#FEE2E2',border:'1px solid #FCA5A5',borderRadius:8,padding:'10px 14px',fontSize:13,color:'#DC2626',marginBottom:14}}>{error}</div>}

      {done&&(
        <div style={{background:'#D1FAE5',border:'1px solid #6EE7B7',borderRadius:8,padding:'14px 18px',fontSize:13,marginBottom:14}}>
          <p style={{fontWeight:700,color:'#065F46',marginBottom:4}}>✓ Import complete</p>
          <p style={{color:'#065F46'}}>{done.written.toLocaleString()} entries written to Firestore.</p>
          <p style={{color:'#64748B',fontSize:12,marginTop:6}}>Refresh the billing app to see imported data.</p>
        </div>
      )}

      {/* File picker */}
      <div style={{border:'2px dashed #E2E8F0',borderRadius:10,padding:'24px',textAlign:'center',marginBottom:16,cursor:'pointer',background:'#FAFBFC'}}
        onClick={()=>fileRef.current.click()}>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>handleFile(e.target.files[0])}/>
        <p style={{fontSize:14,color:'#64748B',marginBottom:4}}>{file?`📄 ${file.name}`:'📂 Click to select Excel file'}</p>
        <p style={{fontSize:11,color:'#94A3B8'}}>Fun_Stuff_2025_2026.xlsx or similar</p>
      </div>

      {/* Preview */}
      {preview&&!done&&(
        <div style={{background:'#F8FAFC',borderRadius:10,padding:'16px',marginBottom:16,border:'1px solid #E2E8F0'}}>
          <p style={{fontWeight:700,color:'#1B3A5C',marginBottom:10,fontSize:14}}>Preview</p>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(160px,1fr))',gap:10,marginBottom:12}}>
            {Object.entries(preview.results).map(([year,entries])=>(
              <div key={year} style={{background:'white',borderRadius:8,padding:'10px 14px',border:'1px solid #E2E8F0',textAlign:'center'}}>
                <p style={{fontSize:11,color:'#94A3B8',marginBottom:2}}>{year}</p>
                <p style={{fontSize:18,fontWeight:800,color:'#1B3A5C'}}>{entries.length.toLocaleString()}</p>
                <p style={{fontSize:10,color:'#94A3B8'}}>entries</p>
              </div>
            ))}
            <div style={{background:'#EFF6FF',borderRadius:8,padding:'10px 14px',border:'1px solid #BFDBFE',textAlign:'center'}}>
              <p style={{fontSize:11,color:'#94A3B8',marginBottom:2}}>Total</p>
              <p style={{fontSize:18,fontWeight:800,color:'#1D4ED8'}}>{preview.totalEntries.toLocaleString()}</p>
              <p style={{fontSize:10,color:'#94A3B8'}}>entries</p>
            </div>
          </div>
          <div style={{marginBottom:12}}>
            <p style={{fontSize:11,fontWeight:700,color:'#64748B',marginBottom:6,textTransform:'uppercase'}}>Doctors found ({preview.doctors.length})</p>
            <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
              {preview.doctors.map(d=>(
                <span key={d} style={{background:'#F1F5F9',borderRadius:4,padding:'2px 8px',fontSize:11,color:'#475569'}}>{d}</span>
              ))}
            </div>
          </div>
          <p style={{fontSize:11,color:'#F59E0B',marginBottom:12}}>
            ⚠ This will write {preview.totalEntries.toLocaleString()} entries to Firestore. This may take a minute or two.
          </p>
          {progress&&<p style={{fontSize:12,color:'#1D4ED8',marginBottom:10,fontWeight:600}}>{progress}</p>}
          <button onClick={runImport} disabled={importing}
            style={{background:importing?'#94A3B8':'#10B981',color:'white',border:'none',borderRadius:7,padding:'10px 24px',cursor:importing?'not-allowed':'pointer',fontSize:14,fontWeight:700}}>
            {importing?'Importing...':'Confirm Import'}
          </button>
        </div>
      )}
    </div>
  );
}

// ── List editor ───────────────────────────────────────────────────────────────
function ListEditor({items, setItems, newVal, setNewVal, onSave, placeholder}) {
  return (
    <div>
      <div style={{display:'flex',gap:8,marginBottom:10}}>
        <input value={newVal} onChange={e=>setNewVal(e.target.value)} placeholder={placeholder}
          onKeyDown={e=>{if(e.key==='Enter'&&newVal.trim()){if(!items.includes(newVal.trim()))setItems(prev=>[...prev,newVal.trim()]);setNewVal('');}}}
          style={{flex:1,padding:'7px 10px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:13,outline:'none'}}/>
        <button onClick={()=>{if(newVal.trim()&&!items.includes(newVal.trim())){setItems(prev=>[...prev,newVal.trim()]);setNewVal('');}}}
          style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'7px 16px',cursor:'pointer',fontSize:13,fontWeight:600}}>+ Add</button>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:4,maxHeight:280,overflow:'auto',marginBottom:12}}>
        {items.map((item,i)=>(
          <div key={i} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'6px 12px',background:'#F8FAFC',borderRadius:7,border:'1px solid #E2E8F0'}}>
            <span style={{fontSize:13}}>{item}</span>
            <button onClick={()=>setItems(prev=>prev.filter((_,j)=>j!==i))}
              style={{background:'#FEE2E2',border:'none',borderRadius:4,padding:'2px 8px',cursor:'pointer',color:'#DC2626',fontSize:11}}>Remove</button>
          </div>
        ))}
      </div>
      <button onClick={onSave} style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'8px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>Save</button>
    </div>
  );
}


// ── Auto-Assign Missing Doctor IDs ───────────────────────────────────────────
function AutoAssignTab() {
  const [running, setRunning] = useState(false);
  const [preview, setPreview] = useState(null);
  const [done, setDone] = useState(null);
  const [error, setError] = useState('');

  const analyze = async () => {
    setRunning(true); setError(''); setPreview(null); setDone(null);
    try {
      const { getDocs, collection, query, orderBy } = await import('firebase/firestore');
      const snap = await getDocs(query(collection(db, 'billingEntries'), orderBy('date','desc')));
      const entries = snap.docs.map(d => ({id: d.id, ...d.data()}));

      // Group by date + location
      const groups = {};
      for (const e of entries) {
        const key = `${e.date}|${e.location}`;
        if (!groups[key]) groups[key] = [];
        groups[key].push(e);
      }

      // For each group, find the majority doctor
      const toAssign = [];
      for (const [key, group] of Object.entries(groups)) {
        const withDoctor = group.filter(e => e.doctorId);
        if (withDoctor.length === 0) continue;
        const withoutDoctor = group.filter(e => !e.doctorId);
        if (withoutDoctor.length === 0) continue;

        // Find majority doctor
        const counts = {};
        for (const e of withDoctor) counts[e.doctorId] = (counts[e.doctorId]||0) + 1;
        const majorityDoctor = Object.entries(counts).sort((a,b)=>b[1]-a[1])[0][0];

        for (const e of withoutDoctor) {
          toAssign.push({ id: e.id, date: e.date, location: e.location, patientName: e.patientName, doctorId: majorityDoctor });
        }
      }

      setPreview(toAssign);
    } catch(e) { setError(e.message); }
    setRunning(false);
  };

  const applyAssignments = async () => {
    if (!preview) return;
    setRunning(true);
    try {
      const { writeBatch, doc, collection } = await import('firebase/firestore');
      const BATCH_SIZE = 200;
      for (let i = 0; i < preview.length; i += BATCH_SIZE) {
        const batch = writeBatch(db);
        for (const item of preview.slice(i, i + BATCH_SIZE)) {
          batch.update(doc(collection(db, 'billingEntries'), item.id), {
            doctorId: item.doctorId,
            updatedAt: new Date().toISOString()
          });
        }
        await batch.commit();
        await new Promise(r => setTimeout(r, 200));
      }
      setDone(preview.length);
      setPreview(null);
    } catch(e) { setError(e.message); }
    setRunning(false);
  };

  return (
    <div style={{background:'white',borderRadius:12,padding:20,border:'1px solid #E2E8F0'}}>
      <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:6}}>Auto-Assign Missing Doctor IDs</h3>
      <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>
        Finds billing entries with no doctor assigned and assigns the majority doctor for that date/location.
        Review before applying.
      </p>
      {error&&<div style={{background:'#FEE2E2',borderRadius:8,padding:'10px 14px',color:'#DC2626',fontSize:12,marginBottom:12}}>{error}</div>}
      {done&&<div style={{background:'#D1FAE5',borderRadius:8,padding:'10px 14px',color:'#065F46',fontSize:12,marginBottom:12}}>✓ Assigned doctors to {done} entries.</div>}

      {!preview&&!done&&(
        <button onClick={analyze} disabled={running}
          style={{background:running?'#94A3B8':'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>
          {running?'Analyzing...':'Analyze Missing Assignments'}
        </button>
      )}

      {preview&&(
        <div>
          <p style={{fontSize:13,fontWeight:600,color:'#1B3A5C',marginBottom:10}}>
            Found {preview.length} entries to assign:
          </p>
          <div style={{maxHeight:300,overflow:'auto',border:'1px solid #E2E8F0',borderRadius:8,marginBottom:12}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead style={{background:'#F8FAFC',position:'sticky',top:0}}>
                <tr>
                  {['Date','Location','Patient','Will be assigned to'].map(h=>(
                    <th key={h} style={{padding:'6px 8px',textAlign:'left',fontSize:9,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',borderBottom:'1px solid #E2E8F0'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.slice(0,200).map((item,i)=>(
                  <tr key={item.id} style={{borderBottom:'1px solid #F0F4F8',background:i%2===0?'white':'#FAFBFC'}}>
                    <td style={{padding:'4px 8px'}}>{item.date}</td>
                    <td style={{padding:'4px 8px'}}>{item.location}</td>
                    <td style={{padding:'4px 8px',fontWeight:500}}>{item.patientName}</td>
                    <td style={{padding:'4px 8px',color:'#1D4ED8',fontWeight:600}}>Dr. {item.doctorId}</td>
                  </tr>
                ))}
                {preview.length>200&&<tr><td colSpan={4} style={{padding:'8px',textAlign:'center',color:'#94A3B8',fontSize:11}}>...and {preview.length-200} more</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button onClick={applyAssignments} disabled={running}
              style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>
              {running?'Applying...':'Apply All Assignments'}
            </button>
            <button onClick={()=>setPreview(null)}
              style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'9px 14px',cursor:'pointer',fontSize:13}}>
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Main AdminPanel ────────────────────────────────────────────────────────────
export default function AdminPanel({ currentUser }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ name:'', email:'', password:'', role:'doctor', doctorId:'' });
  const [creating, setCreating] = useState(false);
  const [msg, setMsg] = useState('');
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState('users');

  const [feeSettings, setFeeSettings] = useState(DEFAULT_FEES);
  const [otherItems, setOtherItems] = useState(DEFAULT_OTHER_ITEMS);
  const [discountTypes, setDiscountTypes] = useState(DEFAULT_DISCOUNT_TYPES);
  const [doctorList, setDoctorList] = useState(DEFAULT_DOCTORS);
  const [newOtherItem, setNewOtherItem] = useState('');
  const [newDiscount, setNewDiscount] = useState('');
  const [newDoctor, setNewDoctor] = useState('');

  useEffect(() => {
    loadAllUsers().then(u => { setUsers(u); setLoading(false); });
    getDoc(doc(db,'billingSettings','fees')).then(s => { if(s.exists()) setFeeSettings(s.data()); });
    getDoc(doc(db,'billingSettings','otherItems')).then(s => { if(s.exists()&&s.data().items) setOtherItems(s.data().items); });
    getDoc(doc(db,'billingSettings','discountTypes')).then(s => { if(s.exists()&&s.data().types) setDiscountTypes(s.data().types); });
    getDoc(doc(db,'billingSettings','doctors')).then(s => { if(s.exists()&&s.data().list) setDoctorList(s.data().list); });
  }, []);

  const showMsg = (m) => { setMsg(m); setTimeout(()=>setMsg(''),2500); };
  const saveFeeSettings = async () => { await setDoc(doc(db,'billingSettings','fees'),feeSettings); showMsg('✓ Fee settings saved'); };
  const saveOtherItems = async () => { await setDoc(doc(db,'billingSettings','otherItems'),{items:otherItems}); showMsg('✓ Other items saved'); };
  const saveDiscountTypes = async () => { await setDoc(doc(db,'billingSettings','discountTypes'),{types:discountTypes}); showMsg('✓ Discount types saved'); };
  const saveDoctorList = async () => { await setDoc(doc(db,'billingSettings','doctors'),{list:doctorList}); showMsg('✓ Doctor list saved'); };

  const createUser = async () => {
    if(!form.name||!form.email||!form.password||!form.role) return;
    setCreating(true); setError('');
    try {
      const cred = await createUserWithEmailAndPassword(auth,form.email,form.password);
      const profile = { uid:cred.user.uid, name:form.name, email:form.email, role:form.role, doctorId:form.doctorId||form.name.split(' ').pop().toLowerCase(), createdAt:new Date().toISOString() };
      await saveUserProfile(cred.user.uid,profile);
      setUsers(prev=>[...prev,profile]);
      showMsg(`✓ Created account for ${form.name}`);
      setForm({name:'',email:'',password:'',role:'doctor',doctorId:''});
      setShowForm(false);
    } catch(e) { setError(e.message); }
    setCreating(false);
  };

  const removeUser = async (uid,name) => {
    if(!window.confirm(`Remove ${name}?`)) return;
    await deleteUser(uid);
    setUsers(prev=>prev.filter(u=>u.uid!==uid));
  };

  const resetPassword = async (email,name) => {
    try { await sendPasswordResetEmail(auth,email); showMsg(`✓ Reset email sent to ${name}`); }
    catch(e) { setError('Could not send reset email: '+e.message); }
  };

  const inp = { padding:'7px 10px', border:'1.5px solid #E2E8F0', borderRadius:7, fontSize:13, fontFamily:"'DM Sans',sans-serif", outline:'none', width:'100%', boxSizing:'border-box' };
  const lbl = (label,children) => (
    <label style={{display:'flex',flexDirection:'column',gap:4}}>
      <span style={{fontSize:11,fontWeight:700,color:'#64748B',textTransform:'uppercase',letterSpacing:'0.05em'}}>{label}</span>
      {children}
    </label>
  );

  const tabs = [['users','👥 Users'],['fees','💰 Fees'],['otherItems','🔧 Other Items'],['discounts','🏷 Discounts'],['doctors','👩‍⚕️ Doctors'],['payors','🏦 Payors'],['import','📥 Import'],['autoassign','🔧 Auto-Assign'],['pendingeob','📋 Pending EOB']];

  return (
    <div style={{maxWidth:700,margin:'0 auto',padding:'0 20px'}}>
      {msg&&<div style={{background:'#D1FAE5',border:'1px solid #6EE7B7',borderRadius:8,padding:'10px 16px',fontSize:13,color:'#065F46',marginBottom:14}}>{msg}</div>}
      {error&&<div style={{background:'#FEE2E2',border:'1px solid #FCA5A5',borderRadius:8,padding:'10px 16px',fontSize:13,color:'#DC2626',marginBottom:14}}>{error}</div>}

      <div style={{display:'flex',gap:4,marginBottom:20,flexWrap:'wrap'}}>
        {tabs.map(([tab,label])=>(
          <button key={tab} onClick={()=>setActiveTab(tab)}
            style={{background:activeTab===tab?'#1B3A5C':'#F1F5F9',color:activeTab===tab?'white':'#64748B',border:'none',borderRadius:7,padding:'7px 14px',cursor:'pointer',fontSize:12,fontWeight:activeTab===tab?700:400}}>
            {label}
          </button>
        ))}
      </div>

      {activeTab==='users'&&(
        <div>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
            <div>
              <h2 style={{fontFamily:"'DM Serif Display',serif",fontSize:20,color:'#1B3A5C'}}>User Management</h2>
              <p style={{fontSize:12,color:'#94A3B8',marginTop:2}}>Add, remove, or reset passwords for billing app users.</p>
            </div>
            <button onClick={()=>setShowForm(v=>!v)} style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:8,padding:'9px 18px',cursor:'pointer',fontSize:13,fontWeight:600}}>+ Add User</button>
          </div>
          {showForm&&(
            <div style={{background:'white',borderRadius:14,padding:'20px 22px',marginBottom:20,border:'1px solid #E2E8F0',boxShadow:'0 2px 8px rgba(0,0,0,0.06)'}}>
              <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:16}}>New User</h3>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,marginBottom:14}}>
                {lbl('Full Name',<input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inp} placeholder="Dr. Jane Smith"/>)}
                {lbl('Email',<input type="email" value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} style={inp} placeholder="jane@thesparkoptometry.app"/>)}
                {lbl('Temp Password',<input type="text" value={form.password} onChange={e=>setForm(f=>({...f,password:e.target.value}))} style={inp} placeholder="They will change this"/>)}
                {lbl('Role',<select value={form.role} onChange={e=>setForm(f=>({...f,role:e.target.value}))} style={{...inp,cursor:'pointer'}}>{ROLES.map(r=><option key={r.value} value={r.value}>{r.label}</option>)}</select>)}
                {lbl('Doctor ID',<input value={form.doctorId} onChange={e=>setForm(f=>({...f,doctorId:e.target.value}))} style={inp} placeholder="e.g. fan, kaneta"/>)}
              </div>
              <p style={{fontSize:11,color:'#94A3B8',marginBottom:12}}>Doctor ID should match the last name used in the billing system (e.g. "fan", "kha").</p>
              <div style={{display:'flex',gap:10}}>
                <button onClick={createUser} disabled={creating||!form.name||!form.email||!form.password}
                  style={{background:creating?'#94A3B8':'#10B981',color:'white',border:'none',borderRadius:8,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>
                  {creating?'Creating...':'Create Account'}
                </button>
                <button onClick={()=>setShowForm(false)} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:8,padding:'9px 14px',cursor:'pointer',fontSize:13}}>Cancel</button>
              </div>
            </div>
          )}
          {loading?<p style={{color:'#94A3B8'}}>Loading users...</p>:(
            <div style={{display:'flex',flexDirection:'column',gap:8}}>
              {users.map(u=>(
                <div key={u.uid} style={{background:'white',borderRadius:10,padding:'12px 16px',border:'1px solid #E2E8F0',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                  <div>
                    <div style={{display:'flex',gap:8,alignItems:'center'}}>
                      <span style={{fontWeight:700,fontSize:14,color:'#1B3A5C'}}>{u.name}</span>
                      <span style={{fontSize:10,fontWeight:700,borderRadius:4,padding:'2px 7px',background:u.role==='master'?'#1B3A5C':u.role==='biller'?'#2E7D8C':'#F1F5F9',color:u.role==='master'?'white':u.role==='biller'?'white':'#64748B'}}>{u.role}</span>
                      {u.doctorId&&<span style={{fontSize:11,color:'#94A3B8'}}>ID: {u.doctorId}</span>}
                    </div>
                    <p style={{fontSize:12,color:'#94A3B8',marginTop:2}}>{u.email}</p>
                  </div>
                  <div style={{display:'flex',gap:6}}>
                    <button onClick={()=>resetPassword(u.email,u.name)} style={{background:'#EFF6FF',border:'none',borderRadius:6,padding:'5px 10px',cursor:'pointer',color:'#1D4ED8',fontSize:11,fontWeight:600}}>Reset PW</button>
                    {u.uid!==currentUser?.uid&&<button onClick={()=>removeUser(u.uid,u.name)} style={{background:'#FEE2E2',border:'none',borderRadius:6,padding:'5px 8px',cursor:'pointer',color:'#DC2626',fontSize:11}}>Remove</button>}
                  </div>
                </div>
              ))}
              {users.length===0&&<p style={{color:'#94A3B8',fontSize:13}}>No users yet.</p>}
            </div>
          )}
        </div>
      )}

      {activeTab==='fees'&&(
        <div style={{background:'white',borderRadius:12,padding:'20px',border:'1px solid #E2E8F0'}}>
          <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:6}}>Base Exam Fee by Location</h3>
          <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>Used to auto-calculate CL fee for Self Pay bundle visits.</p>
          <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12,marginBottom:16}}>
            {LOCATIONS.map(loc=>(
              <label key={loc} style={{display:'flex',flexDirection:'column',gap:4}}>
                <span style={{fontSize:11,fontWeight:700,color:'#64748B',textTransform:'uppercase'}}>{loc}</span>
                <input type="number" value={feeSettings[loc]||''} onChange={e=>setFeeSettings(f=>({...f,[loc]:parseFloat(e.target.value)||0}))}
                  style={{padding:'8px 10px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:14,fontWeight:700,textAlign:'right',outline:'none',width:'100%',boxSizing:'border-box'}}/>
              </label>
            ))}
          </div>
          <button onClick={saveFeeSettings} style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>Save Fee Settings</button>
        </div>
      )}

      {activeTab==='otherItems'&&(
        <div style={{background:'white',borderRadius:12,padding:'20px',border:'1px solid #E2E8F0'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
            <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>Other Service Items</h3>
            <button onClick={()=>{setOtherItems([...DEFAULT_OTHER_ITEMS]);}} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'4px 10px',cursor:'pointer',fontSize:11,color:'#64748B'}}>↺ Reset to defaults</button>
          </div>
          <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>Items shown in the "Other" dropdown during billing entry.</p>
          <ListEditor items={otherItems} setItems={setOtherItems} newVal={newOtherItem} setNewVal={setNewOtherItem} onSave={saveOtherItems} placeholder="Add new service item..."/>
        </div>
      )}

      {activeTab==='discounts'&&(
        <div style={{background:'white',borderRadius:12,padding:'20px',border:'1px solid #E2E8F0'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
            <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>Discount Types</h3>
            <button onClick={()=>{setDiscountTypes([...DEFAULT_DISCOUNT_TYPES]);}} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'4px 10px',cursor:'pointer',fontSize:11,color:'#64748B'}}>↺ Reset to defaults</button>
          </div>
          <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>Dropdown options for the discount type field during billing entry.</p>
          <ListEditor items={discountTypes} setItems={setDiscountTypes} newVal={newDiscount} setNewVal={setNewDiscount} onSave={saveDiscountTypes} placeholder="Add new discount type..."/>
        </div>
      )}

      {activeTab==='doctors'&&(
        <div style={{background:'white',borderRadius:12,padding:'20px',border:'1px solid #E2E8F0'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
            <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>Active Doctor List</h3>
            <button onClick={()=>{setDoctorList([...DEFAULT_DOCTORS]);}} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'4px 10px',cursor:'pointer',fontSize:11,color:'#64748B'}}>↺ Reset to defaults</button>
          </div>
          <p style={{fontSize:12,color:'#94A3B8',marginBottom:14}}>
            Names shown in doctor dropdowns, scheduling calendar, and entry forms.
            Former doctors should be removed from here — their historical data is preserved.
          </p>
          <ListEditor items={doctorList} setItems={setDoctorList} newVal={newDoctor} setNewVal={setNewDoctor} onSave={saveDoctorList} placeholder="Add doctor last name..."/>
        </div>
      )}

      {activeTab==='import'&&<ImportTab/>}

      {activeTab==='autoassign'&&<AutoAssignTab/>}
      {activeTab==='payors'&&<PayorsTab/>}
      {activeTab==='pendingeob'&&<PendingEOBTab/>}

      <div style={{marginTop:20,padding:'14px 16px',background:'#F8FAFC',borderRadius:10,border:'1px solid #E2E8F0'}}>
        <p style={{fontSize:11,fontWeight:700,color:'#64748B',marginBottom:6}}>Setup Notes</p>
        <ol style={{fontSize:11,color:'#64748B',lineHeight:1.8,paddingLeft:16,margin:0}}>
          <li>Create accounts for each doctor using their preferred email</li>
          <li>Give them their temporary password — they change it on first login</li>
          <li>Doctor ID must match last name in billing system (lowercase)</li>
          <li>Doctors see only their own log. Billers see everything except Admin.</li>
          <li>Use Import tab to bring in historical data from the production spreadsheet</li>
        </ol>
      </div>
    </div>
  );
}

// ── Payors Tab ────────────────────────────────────────────────────────────────
const DEFAULT_PAYORS = [
  { value: 'VSP',      label: 'VSP',            remit: 'assignment' },
  { value: 'EM',       label: 'EyeMed',          remit: 'eyemed' },
  { value: 'Self',     label: 'Self Pay',        remit: 'none' },
  { value: 'UHC',      label: 'UHC/Spectera',    remit: 'assignment' },
  { value: 'Aetna',    label: 'Aetna',           remit: 'eyemed' },
  { value: 'Avesis',   label: 'Avesis',          remit: 'assignment' },
  { value: 'BCBS',     label: 'BCBS',            remit: 'eyemed' },
  { value: 'Cigna',    label: 'Cigna',           remit: 'eyemed' },
  { value: 'DV',       label: 'Davis Vision',    remit: 'assignment' },
  { value: 'FEP',      label: 'FEP',             remit: 'assignment' },
  { value: 'Heritage', label: 'Heritage',        remit: 'assignment' },
  { value: 'MES',      label: 'MES Vision',      remit: 'assignment' },
  { value: 'NVA',      label: 'NVA',             remit: 'assignment' },
  { value: 'SUP',      label: 'Superior Vision', remit: 'assignment' },
  { value: 'VBA',      label: 'VBA',             remit: 'assignment' },
  { value: 'Other',    label: 'Other',           remit: 'none' },
];

function PayorsTab() {
  const [payors, setPayors] = React.useState([]);
  const [loading, setLoading] = React.useState(true);
  const [saving, setSaving] = React.useState(false);
  const [newLabel, setNewLabel] = React.useState('');
  const [newRemit, setNewRemit] = React.useState('assignment');

  React.useEffect(() => {
    const load = async () => {
      const { getDoc, doc } = await import('firebase/firestore');
      const snap = await getDoc(doc(db, 'billingSettings', 'payors'));
      if (snap.exists() && snap.data().list?.length > 0) {
        setPayors(snap.data().list);
      } else {
        setPayors(DEFAULT_PAYORS);
      }
      setLoading(false);
    };
    load();
  }, []);

  const save = async (updated) => {
    setSaving(true);
    const { setDoc, doc } = await import('firebase/firestore');
    await setDoc(doc(db, 'billingSettings', 'payors'), { list: updated, updatedAt: new Date().toISOString() });
    setSaving(false);
  };

  const addPayor = () => {
    if (!newLabel.trim()) return;
    const updated = [...payors, { value: newLabel.trim().replace(/\s+/g,''), label: newLabel.trim(), remit: newRemit }];
    setPayors(updated);
    save(updated);
    setNewLabel('');
  };

  const removePayor = (idx) => {
    const updated = payors.filter((_, i) => i !== idx);
    setPayors(updated);
    save(updated);
  };

  const updateRemit = (idx, remit) => {
    const updated = payors.map((p, i) => i === idx ? { ...p, remit } : p);
    setPayors(updated);
    save(updated);
  };

  const inp = { padding: '5px 8px', border: '1px solid #E2E8F0', borderRadius: 6, fontSize: 12, fontFamily: "'DM Sans',sans-serif" };

  if (loading) return <div style={{padding:20,color:'#94A3B8'}}>Loading...</div>;

  return (
    <div style={{background:'white',borderRadius:12,padding:20,border:'1px solid #E2E8F0'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
        <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>Insurance Payors</h3>
        <button onClick={()=>{setPayors([...DEFAULT_PAYORS]);save([...DEFAULT_PAYORS]);}} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'4px 10px',cursor:'pointer',fontSize:11,color:'#64748B'}}>↺ Reset to defaults</button>
      </div>
      <p style={{fontSize:12,color:'#94A3B8',marginBottom:16}}>Manage insurance payors and their remittance type (EyeMed or Assignment Plan). This determines how EOB imports are matched.</p>

      <table style={{width:'100%',borderCollapse:'collapse',fontSize:12,marginBottom:16}}>
        <thead>
          <tr style={{borderBottom:'2px solid #F0F4F8'}}>
            {['Code','Label','Remittance Type',''].map(h=>(
              <th key={h} style={{padding:'6px 8px',textAlign:'left',fontSize:10,fontWeight:700,color:'#94A3B8',textTransform:'uppercase'}}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {payors.map((p, i) => (
            <tr key={i} style={{borderBottom:'1px solid #F0F4F8'}}>
              <td style={{padding:'5px 8px',fontWeight:600,color:'#1B3A5C'}}>{p.value}</td>
              <td style={{padding:'5px 8px'}}>{p.label}</td>
              <td style={{padding:'5px 8px'}}>
                <select value={p.remit||'assignment'} onChange={e=>updateRemit(i,e.target.value)}
                  style={{...inp,fontSize:11}}>
                  <option value="eyemed">EyeMed</option>
                  <option value="assignment">Assignment Plan (VSP-style)</option>
                  <option value="none">N/A (Self/Other)</option>
                </select>
              </td>
              <td style={{padding:'5px 8px'}}>
                <button onClick={()=>removePayor(i)}
                  style={{background:'#FEE2E2',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:11,color:'#DC2626'}}>
                  Remove
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={{display:'flex',gap:8,alignItems:'center',borderTop:'1px solid #F0F4F8',paddingTop:12}}>
        <input value={newLabel} onChange={e=>setNewLabel(e.target.value)} placeholder="New payor name (e.g. Humana)"
          style={{...inp,flex:1}} onKeyDown={e=>e.key==='Enter'&&addPayor()}/>
        <select value={newRemit} onChange={e=>setNewRemit(e.target.value)} style={{...inp}}>
          <option value="eyemed">EyeMed</option>
          <option value="assignment">Assignment Plan</option>
          <option value="none">N/A</option>
        </select>
        <button onClick={addPayor} disabled={saving}
          style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'8px 16px',cursor:'pointer',fontSize:12,fontWeight:600}}>
          {saving?'Saving...':'Add Payor'}
        </button>
      </div>
    </div>
  );
}

// ── Pending EOB Tab ───────────────────────────────────────────────────────────
function PendingEOBTab() {
  const [batches, setBatches] = React.useState([]);
  const [loading, setLoading] = React.useState(true);
  const [expanded, setExpanded] = React.useState({});

  React.useEffect(() => {
    const load = async () => {
      const { getDocs, collection, orderBy, query } = await import('firebase/firestore');
      const snap = await getDocs(query(collection(db, 'pendingEOB'), orderBy('importedAt', 'desc')));
      setBatches(snap.docs.map(d => ({ id: d.id, ...d.data() })));
      setLoading(false);
    };
    load();
  }, []);

  const dismissClaim = async (batchId, claimIdx) => {
    const { getDoc, setDoc, doc } = await import('firebase/firestore');
    const batch = batches.find(b => b.id === batchId);
    if (!batch) return;
    const updatedClaims = batch.claims.filter((_, i) => i !== claimIdx);
    const updatedBatch = { ...batch, claims: updatedClaims };
    await setDoc(doc(db, 'pendingEOB', batchId), updatedBatch);
    setBatches(prev => prev.map(b => b.id === batchId
      ? { ...b, claims: updatedClaims }
      : b
    ).filter(b => b.claims.length > 0));
  };

  const dismissBatch = async (batchId) => {
    if (!window.confirm('Dismiss all claims in this batch?')) return;
    const { deleteDoc, doc } = await import('firebase/firestore');
    await deleteDoc(doc(db, 'pendingEOB', batchId));
    setBatches(prev => prev.filter(b => b.id !== batchId));
  };

  if (loading) return <div style={{padding:20,color:'#94A3B8'}}>Loading...</div>;

  if (batches.length === 0) return (
    <div style={{background:'white',borderRadius:12,padding:20,border:'1px solid #E2E8F0',textAlign:'center',color:'#94A3B8',fontSize:13}}>
      ✓ No pending EOB claims. All imports resolved.
    </div>
  );

  return (
    <div style={{background:'white',borderRadius:12,padding:20,border:'1px solid #E2E8F0'}}>
      <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:4}}>Pending EOB Claims</h3>
      <p style={{fontSize:12,color:'#94A3B8',marginBottom:16}}>Unmatched and suggested EOB claims from previous imports. Dismiss when resolved.</p>

      {batches.map(batch => (
        <div key={batch.id} style={{border:'1px solid #E2E8F0',borderRadius:10,marginBottom:12,overflow:'hidden'}}>
          <div style={{background:'#F8FAFC',padding:'8px 12px',display:'flex',justifyContent:'space-between',alignItems:'center',cursor:'pointer'}}
            onClick={()=>setExpanded(prev=>({...prev,[batch.id]:!prev[batch.id]}))}>
            <div>
              <span style={{fontWeight:700,fontSize:13,color:'#1B3A5C'}}>
                {new Date(batch.importedAt).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}
              </span>
              <span style={{marginLeft:8,fontSize:11,color:'#94A3B8',textTransform:'uppercase'}}>{batch.source}</span>
              <span style={{marginLeft:8,background:'#FEE2E2',color:'#DC2626',borderRadius:4,padding:'1px 6px',fontSize:11,fontWeight:700}}>
                {batch.claims.length} pending
              </span>
            </div>
            <div style={{display:'flex',gap:8,alignItems:'center'}}>
              <button onClick={e=>{e.stopPropagation();dismissBatch(batch.id);}}
                style={{background:'#F1F5F9',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:11,color:'#64748B'}}>
                Dismiss All
              </button>
              <span style={{color:'#94A3B8'}}>{expanded[batch.id]?'▼':'▶'}</span>
            </div>
          </div>

          {expanded[batch.id] && (
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead>
                <tr style={{borderBottom:'1px solid #E2E8F0',background:'#FAFBFC'}}>
                  {['Status','EOB Patient','DOS','EOB $','Suggested Match','Score',''].map(h=>(
                    <th key={h} style={{padding:'5px 8px',textAlign:'left',fontSize:9,fontWeight:700,color:'#94A3B8',textTransform:'uppercase'}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {batch.claims.map((claim, idx) => (
                  <tr key={idx} style={{borderBottom:'1px solid #F0F4F8',background:idx%2===0?'white':'#FAFBFC'}}>
                    <td style={{padding:'5px 8px'}}>
                      <span style={{background:claim.status==='unmatched'?'#F1F5F9':'#FDF4FF',color:claim.status==='unmatched'?'#64748B':'#7E22CE',borderRadius:4,padding:'2px 6px',fontSize:10,fontWeight:600}}>
                        {claim.status==='unmatched'?'? No match':'~ Suggested'}
                      </span>
                    </td>
                    <td style={{padding:'5px 8px',fontWeight:600}}>{claim.memberName}</td>
                    <td style={{padding:'5px 8px',color:'#64748B'}}>{claim.dos}</td>
                    <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,color:'#1B3A5C'}}>${claim.amount?.toFixed(2)}</td>
                    <td style={{padding:'5px 8px',color:'#7E22CE',fontStyle: claim.suggestedEntryName?'normal':'italic'}}>
                      {claim.suggestedEntryName||'—'}
                    </td>
                    <td style={{padding:'5px 8px',color:'#94A3B8'}}>
                      {claim.suggestedScore?(claim.suggestedScore*100).toFixed(0)+'%':'—'}
                    </td>
                    <td style={{padding:'5px 8px'}}>
                      <button onClick={()=>dismissClaim(batch.id, idx)}
                        style={{background:'#FEE2E2',border:'none',borderRadius:4,padding:'2px 6px',cursor:'pointer',fontSize:10,color:'#DC2626'}}>
                        Dismiss
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      ))}
    </div>
  );
}
