import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { db } from './firebase';
import { collection, getDocs, setDoc, doc, deleteDoc, orderBy, query, getDoc, addDoc } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import { LOCATIONS, LOC_FULL, LOC_COLORS, PAYORS, DISCOUNT_TYPES, uid, today, fmtDate, fmt$, emptyEntry, computeStatus, getWeek, getFiscalWeekInfo, getFiscalMonthKey } from './utils';
import EOBImportModal from './EOBImport';

const LOC_ORDER = { SC:0, F:1, WC:2, SV:3 };
const DEFAULT_FEES = { SC:89, F:89, SV:89, WC:95 };
const DAY_ABBR = ['Su','Mo','Tu','We','Th','Fr','Sa'];
const DAY_FULL = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DEFAULT_OTHER_ITEMS = ['Topo','Axial Length','LASIK','Materials','Glaucoma Check','MNCL Reimbursement','No-Show Fee','Oasis Materials','Other'];
const DEFAULT_DOCTORS = ['Kha','Pan','Fan','Kaneta','Yang','Ghag','Zhang','So','Burger','Duong','Cheng','Luong'];

async function loadFeeSettings() { try { const s=await getDoc(doc(db,'billingSettings','fees')); return s.exists()?s.data():DEFAULT_FEES; } catch { return DEFAULT_FEES; } }
async function loadOtherItems() { try { const s=await getDoc(doc(db,'billingSettings','otherItems')); return s.exists()?s.data().items:DEFAULT_OTHER_ITEMS; } catch { return DEFAULT_OTHER_ITEMS; } }
async function loadDoctorList() { try { const s=await getDoc(doc(db,'billingSettings','doctors')); return s.exists()?s.data().list:DEFAULT_DOCTORS; } catch { return DEFAULT_DOCTORS; } }
async function loadScheduleTemplate() { try { const s=await getDoc(doc(db,'billingSettings','scheduleTemplate')); return s.exists()?s.data():{assignments:{},updatedAt:''}; } catch { return {assignments:{},updatedAt:''}; } }
async function saveScheduleTemplate(data) { await setDoc(doc(db,'billingSettings','scheduleTemplate'),{...data,updatedAt:new Date().toISOString()}); }
async function createNotification(msg, resolvedAttn, comment) { await addDoc(collection(db,'notifications'), { message:msg, resolvedAttn, comment, createdAt:new Date().toISOString(), read:false }); }

function sortEntries(arr) {
  return [...arr].sort((a,b)=>{
    if(a.date!==b.date) return b.date.localeCompare(a.date);
    const la=LOC_ORDER[a.location]??99, lb=LOC_ORDER[b.location]??99;
    if(la!==lb) return la-lb;
    return (a.createdAt||'').localeCompare(b.createdAt||'');
  });
}
function dayOfWeek(d,abbr=false) { if(!d)return''; const dt=new Date(d+'T12:00:00'); return abbr?DAY_ABBR[dt.getDay()]:DAY_FULL[dt.getDay()]; }
function addDays(d,n) { const dt=new Date(d+'T12:00:00'); dt.setDate(dt.getDate()+n); return dt.toISOString().slice(0,10); }
function weekStart(off=0) { const d=new Date(); d.setDate(d.getDate()-d.getDay()+off*7); return d.toISOString().slice(0,10); }
function getWeekLabel(dateStr) {
  // Uses fiscal week (Sunday start) from getFiscalWeekInfo
  const d=new Date(dateStr+'T12:00:00');
  const sun=new Date(d); sun.setDate(d.getDate()-d.getDay()); // this Sunday
  const sat=new Date(sun); sat.setDate(sun.getDate()+6);
  const fmt=(dt)=>dt.toLocaleDateString('en-US',{month:'short',day:'numeric'});
  return `${fmt(sun)} – ${fmt(sat)}`;
}

// ── Focus ring style ────────────────────────────────────────────────────────
const focusStyle = `
  .spark-input:focus { outline: 2px solid #2E7D8C !important; border-color: #2E7D8C !important; background: #F0F9FF !important; }
  .spark-select:focus { outline: 2px solid #2E7D8C !important; border-color: #2E7D8C !important; background: #EFF6FF !important; }
  .spark-cal-cell:focus { outline: 2px solid #2E7D8C; background: #EFF6FF; }
`;

// ── Status badge ────────────────────────────────────────────────────────────
function StatusCell({entry}) {
  const status=computeStatus(entry);
  const cfg={pending:{bg:'#FEF3C7',color:'#92400E',label:'Pending'},flagged:{bg:'#FEE2E2',color:'#DC2626',label:'Flagged'},done:{bg:'#D1FAE5',color:'#065F46',label:'Done'},partial:{bg:'#E0F2FE',color:'#0369A1',label:'◑ Partial'}}[status]||{bg:'#FEF3C7',color:'#92400E',label:'Pending'};
  return (
    <div style={{display:'flex',alignItems:'center',gap:3}}>
      <span style={{fontSize:9,fontWeight:700,borderRadius:4,padding:'2px 5px',background:cfg.bg,color:cfg.color,whiteSpace:'nowrap'}}>{cfg.label}</span>
      {entry.notes&&<span title={entry.notes} style={{cursor:'help',fontSize:10}}>📝</span>}
      {entry.attn&&!entry.attnResolved&&<span title={entry.attn} style={{cursor:'help',fontSize:10}}>⚠️</span>}
    </div>
  );
}

// ── Ins Paid 4-state ────────────────────────────────────────────────────────
function InsPaidCell({entry,onUpdate}) {
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState('');
  const isSelf=entry.payor1==='Self'&&!entry.payor2;
  if(isSelf) return <span style={{color:'#E2E8F0',fontSize:9}}>—</span>;
  const state=entry.insPaidState||'pending';
  const paid=entry.insurancePaid1; const exp=parseFloat(entry.ins)||0; const paidAmt=parseFloat(paid)||0;
  const matches=paid&&Math.abs(paidAmt-exp)<0.01;
  const cycle=async()=>{
    const next={pending:'confirmed',confirmed:'bad',bad:'bad-done','bad-done':'pending'}[state]||'pending';
    const updates={insPaidState:next};
    if(next==='confirmed'&&!paid) updates.insurancePaid1=String(exp);
    await onUpdate(entry.id,'_multi',updates);
  };
  if(editing) return <input autoFocus type="text" value={val} onChange={e=>setVal(e.target.value)} onFocus={e=>e.target.select()}
    onBlur={()=>{setEditing(false);onUpdate(entry.id,'insurancePaid1',val);}}
    onKeyDown={e=>{if(e.key==='Enter'||e.key==='Escape'){setEditing(false);onUpdate(entry.id,'insurancePaid1',val);}}}
    style={{width:50,padding:'2px 3px',border:'1.5px solid #2E7D8C',borderRadius:4,fontSize:10,outline:'none',textAlign:'right'}}/>;
  const stCfg={
    pending:   {bg:'#FFFBEB',border:'#FCD34D',color:'#92400E',text:'Pend'},
    confirmed: {bg:'#D1FAE5',border:'#6EE7B7',color:'#065F46',text:'✓'+(matches?'':' '+fmt$(paidAmt))},
    bad:       {bg:'#FEE2E2',border:'#FCA5A5',color:'#DC2626',text:'!'+fmt$(paidAmt)},
    'bad-done':{bg:'#FCE7F3',border:'#F9A8D4',color:'#9D174D',text:'~'+fmt$(paidAmt)},
  }[state];
  return (
    <div style={{display:'flex',alignItems:'center',gap:2}}>
      <button onClick={cycle} style={{background:stCfg.bg,border:`1px solid ${stCfg.border}`,borderRadius:4,padding:'2px 5px',cursor:'pointer',fontSize:9,color:stCfg.color,fontWeight:700,whiteSpace:'nowrap'}}>{stCfg.text}</button>
      {state!=='pending'&&<button onClick={()=>{setVal(paid||'');setEditing(true);}} style={{background:'none',border:'none',cursor:'pointer',fontSize:9,color:'#94A3B8',padding:'0 1px'}}>✎</button>}
    </div>
  );
}

function CashToggle({value,onUpdate}) {
  const next={'':'expected','expected':'received','received':''};
  const cfg={
    '':        {label:'—',    bg:'#F1F5F9',color:'#94A3B8'},
    'expected':{label:'Exp',  bg:'#FEF3C7',color:'#92400E'},
    'received':{label:"Rec'd",bg:'#D1FAE5',color:'#065F46'},
  };
  const c=cfg[value||'']||cfg[''];
  return <button onClick={()=>onUpdate(next[value||'']||'')} style={{background:c.bg,color:c.color,border:'none',borderRadius:4,padding:'2px 5px',cursor:'pointer',fontSize:9,fontWeight:700,whiteSpace:'nowrap'}}>{c.label}</button>;
}

function InlineCell({value,onUpdate,field,width=42}) {
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState(value||'');
  useEffect(()=>setVal(value||''),[value]);
  if(!editing) return <span onClick={()=>setEditing(true)} style={{cursor:'pointer',borderBottom:'1px dashed #CBD5E1',minWidth:width,display:'inline-block',fontSize:10,textAlign:'right',color:value?'#1B3A5C':'#CBD5E1'}}>{value?fmt$(value):'—'}</span>;
  return <input autoFocus type="text" value={val} onChange={e=>setVal(e.target.value)}
    onBlur={()=>{setEditing(false);if(val!==String(value||''))onUpdate(field,val);}}
    onKeyDown={e=>{if(e.key==='Enter'||e.key==='Escape'){setEditing(false);if(val!==String(value||''))onUpdate(field,val);}}}
    style={{width:width+8,padding:'2px 3px',border:'1.5px solid #2E7D8C',borderRadius:4,fontSize:10,outline:'none',textAlign:'right'}}/>;
}

// ── Inline payor select (editable in review) ────────────────────────────────
function InlinePayorSelect({value,onUpdate,field}) {
  const [editing,setEditing]=useState(false);
  const [val,setVal]=useState(value||'');
  useEffect(()=>setVal(value||''),[value]);
  if(!editing) return (
    <span onClick={()=>setEditing(true)} style={{cursor:'pointer',borderBottom:'1px dashed #CBD5E1',fontSize:10,color:value?'#64748B':'#CBD5E1',display:'inline-block',minWidth:28}}>
      {value||'—'}
    </span>
  );
  return (
    <select autoFocus value={val} onChange={e=>setVal(e.target.value)}
      onBlur={()=>{setEditing(false);if(val!==String(value||''))onUpdate(field,val);}}
      style={{width:52,padding:'1px 2px',border:'1.5px solid #2E7D8C',borderRadius:4,fontSize:10,outline:'none',background:'white'}}>
      <option value="">—</option>
      {PAYORS.map(p=><option key={p.value} value={p.value}>{p.value}</option>)}
    </select>
  );
}

// ── Entry row ───────────────────────────────────────────────────────────────
// FIELDS in tab/arrow order
const FIELDS = ['patientName','exam','cl','optos','oct','dfe','ov','myopia','otherType','otherAmt','ins1Amt','payor1','ins2Amt','payor2','cashStatus','paymentErrorLoss','attn','notes'];

function EntryRow({entry, onSave, onDelete, feeSettings, lockedLoc, lockedDoctor, otherItems, rowIndex, onFocusRow, registerRefs, onInsertAfter}) {
  const [form, setForm] = useState({...entry});
  const [saved, setSaved] = useState(!!entry.patientName);
  const [dirty, setDirty] = useState(false);
  const [showMore, setShowMore] = useState(false);
  const [focused, setFocused] = useState(null);
  const refs = useRef({});

  const set = useCallback((k, v) => {
    setDirty(true); setSaved(false);
    setForm(prev => {
      const next = {...prev, [k]: v};
      if (k === 'exam' || k === 'payor1') {
        const amt = parseFloat(k === 'exam' ? v : next.exam) || 0;
        const payor = k === 'payor1' ? v : next.payor1;
        const base = feeSettings[next.location || lockedLoc] || 89;
        if (payor === 'Self' && amt > base) { next.exam = String(base); next.cl = String((amt - base).toFixed(2)); }
      }
      return next;
    });
  }, [feeSettings, lockedLoc]);

  const ptPaid = useMemo(() => {
    const sum = ['exam','cl','optos','oct','dfe','ov'].reduce((s,k) => s + (parseFloat(form[k])||0), 0)
      + (parseFloat(form.otherAmt)||0) + (parseFloat(form.otherAmt2)||0);
    return Math.max(0, sum - (parseFloat(form.discountAmount)||0));
  }, [form]);

  // Ins total = sum of P1 + P2 + P3 amounts
  const insTotal = useMemo(() => {
    const p1 = parseFloat(form.ins1Amt)||0;
    const p2 = parseFloat(form.ins2Amt)||0;
    const p3 = parseFloat(form.ins3Amt)||0;
    // If no individual amounts set but ins field has value, use ins as P1
    if (p1 === 0 && p2 === 0 && p3 === 0) return parseFloat(form.ins)||0;
    return p1 + p2 + p3;
  }, [form]);

  const grandTotal = ptPaid + insTotal;

  const save = useCallback(() => {
    if (!form.patientName?.trim()) return;
    const p1 = parseFloat(form.ins1Amt)||0;
    const p2 = parseFloat(form.ins2Amt)||0;
    const p3 = parseFloat(form.ins3Amt)||0;
    const computedIns = (p1+p2+p3) > 0 ? String(p1+p2+p3) : form.ins||'';
    const updated = {
      ...form,
      ins: computedIns,
      ptPaid: String(ptPaid),
      location: form.location||lockedLoc||'',
      doctorId: form.doctorId||lockedDoctor||'',
      updatedAt: new Date().toISOString()
    };
    // Update UI immediately — don't wait for Firestore
    onSave(updated);
    setSaved(true);
    setDirty(false);
    // Write to Firestore in background
    setDoc(doc(db, 'billingEntries', form.id), updated)
      .catch(e => { console.error('Save failed:', e); setSaved(false); setDirty(true); });
  }, [form, ptPaid, lockedLoc, lockedDoctor, onSave]);

  // Track if a select dropdown is open so we don't intercept its arrow keys
  const selectOpen = useRef(false);

  const move = (field, dir) => {
    const idx = FIELDS.indexOf(field);
    if (dir === 'right' || dir === 'next') {
      if (idx < FIELDS.length - 1) refs.current[FIELDS[idx+1]]?.focus();
      else { save().then(() => onSave(null, 'newrow')); }
    } else if (dir === 'left' || dir === 'prev') {
      if (idx > 0) refs.current[FIELDS[idx-1]]?.focus();
    } else if (dir === 'down') {
      onFocusRow && onFocusRow(rowIndex + 1, field);
    } else if (dir === 'up') {
      onFocusRow && onFocusRow(rowIndex - 1, field);
    }
  };

  const onKey = (e, field) => {
    // Tab always moves to next/prev field regardless of element type
    if (e.key === 'Tab') { e.preventDefault(); move(field, e.shiftKey ? 'prev' : 'next'); return; }

    // For selects: Space/Enter opens dropdown (mark open), Escape closes it.
    // Arrow keys only navigate rows/fields when dropdown is closed.
    if (e.target.tagName === 'SELECT') {
      if (e.key === ' ' || e.key === 'Enter') { selectOpen.current = true; return; }
      if (e.key === 'Escape') { selectOpen.current = false; return; }
      // If dropdown is open, let native select handle up/down
      if (selectOpen.current && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) return;
      // Left/Right always move between fields even in a closed select
      if (e.key === 'ArrowRight') { e.preventDefault(); selectOpen.current = false; move(field, 'right'); return; }
      if (e.key === 'ArrowLeft')  { e.preventDefault(); selectOpen.current = false; move(field, 'left');  return; }
      // Up/Down move rows when dropdown is closed
      if (e.key === 'ArrowDown') { e.preventDefault(); selectOpen.current = false; move(field, 'down'); return; }
      if (e.key === 'ArrowUp')   { e.preventDefault(); selectOpen.current = false; move(field, 'up');   return; }
      return;
    }

    // Regular inputs
    if (e.key === 'ArrowRight') { e.preventDefault(); move(field, 'right'); return; }
    if (e.key === 'ArrowLeft')  { e.preventDefault(); move(field, 'left');  return; }
    if (e.key === 'ArrowDown')  { e.preventDefault(); move(field, 'down');  return; }
    if (e.key === 'ArrowUp')    { e.preventDefault(); move(field, 'up');    return; }
  };

  // When a select value changes (user picked something), mark dropdown as closed
  const onSelChange = (field, val) => {
    selectOpen.current = false;
    set(field, val);
    setTimeout(save, 50);
  };

  const focusStyle = (field) => ({
    outline: focused === field ? '2px solid #2E7D8C' : 'none',
    background: focused === field ? '#EFF6FF' : 'white',
  });

  const ns = (w=34) => ({width:w, padding:'2px 3px', border:'1px solid #E2E8F0', borderRadius:3, fontSize:10, textAlign:'right', fontFamily:"'DM Sans',sans-serif"});
  const ts = (w=110) => ({width:w, padding:'2px 3px', border:'1px solid #E2E8F0', borderRadius:3, fontSize:10, fontFamily:"'DM Sans',sans-serif"});

  const numInp = (field, w=34) => (
    <input type="text" inputMode="decimal" value={form[field]||''} placeholder=""
      style={{...ns(w), ...focusStyle(field)}}
      ref={el => refs.current[field] = el}
      onChange={e => set(field, e.target.value)}
      onFocus={() => setFocused(field)}
      onBlur={() => { setFocused(null); save(); }}
      onKeyDown={e => onKey(e, field)}
    />
  );

  const selInp = (field, items, w=68, isPayor=false) => (
    <select value={form[field]||''} style={{...ts(w), fontSize:9, cursor:'pointer', ...focusStyle(field)}}
      ref={el => refs.current[field] = el}
      onChange={e => onSelChange(field, e.target.value)}
      onFocus={() => setFocused(field)}
      onBlur={() => { setFocused(null); selectOpen.current = false; }}
      onKeyDown={e => onKey(e, field)}
      onMouseDown={() => { selectOpen.current = !selectOpen.current; }}
    >
      <option value="">—</option>
      {isPayor ? PAYORS.map(p=><option key={p.value} value={p.value}>{p.value}</option>)
               : items.map(i=><option key={i} value={i}>{i}</option>)}
    </select>
  );

  const savedDot = form.patientName?.trim()
    ? <span style={{fontSize:8, color:dirty?'#F59E0B':saved?'#10B981':'transparent'}}>{dirty?'●':'✓'}</span>
    : null;

  return (
    <>
      <tr style={{borderBottom: showMore?'none':'1px solid #F0F4F8', background:'white'}}>
        <td style={{padding:'2px 2px', width:12, textAlign:'center'}}>{savedDot}</td>
        <td style={{padding:'2px 2px', minWidth:112}}>
          <input type="text" value={form.patientName||''} placeholder="Last, First"
            style={{...ts(110), ...focusStyle('patientName')}}
            ref={el => { refs.current.patientName = el; if(el) registerRefs && registerRefs(rowIndex, refs.current); }}
            onChange={e => set('patientName', e.target.value)}
            onFocus={() => setFocused('patientName')}
            onBlur={() => { setFocused(null); save(); }}
            onKeyDown={e => onKey(e, 'patientName')}
          />
        </td>
        <td style={{padding:'2px 1px'}}>{numInp('exam', 32)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('cl', 30)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('optos', 30)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('oct', 30)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('dfe', 30)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('ov', 30)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('myopia', 30)}</td>
        <td style={{padding:'2px 1px'}}>{selInp('otherType', otherItems, 62)}</td>
        <td style={{padding:'2px 1px'}}>{numInp('otherAmt', 30)}</td>
        {/* Pt Paid — lighter green */}
        <td style={{padding:'2px 2px', fontSize:10, color:'#4ADE80', textAlign:'right', whiteSpace:'nowrap'}}>{ptPaid>0?fmt$(ptPaid):'—'}</td>
        {/* Ins = sum of P1+P2+P3 amounts — read only */}
        <td style={{padding:'2px 2px', fontSize:10, color:'#38BDF8', textAlign:'right', whiteSpace:'nowrap'}}>{insTotal>0?fmt$(insTotal):'—'}</td>
        {/* Grand Total — dark green bold */}
        <td style={{padding:'2px 2px', fontWeight:700, fontSize:10, color:'#16A34A', textAlign:'right', whiteSpace:'nowrap'}}>{grandTotal>0?fmt$(grandTotal):'—'}</td>
        {/* P1$ | P1 */}
        <td style={{padding:'2px 1px'}}>{numInp('ins1Amt', 30)}</td>
        <td style={{padding:'2px 1px'}}>{selInp('payor1', null, 42, true)}</td>
        {/* P2$ | P2 */}
        <td style={{padding:'2px 1px'}}>{numInp('ins2Amt', 30)}</td>
        <td style={{padding:'2px 1px'}}>{selInp('payor2', null, 42, true)}</td>
        {/* Cash */}
        <td style={{padding:'2px 1px'}} onKeyDown={e=>{
          if(e.key==='Tab'){e.preventDefault();move('cashStatus',e.shiftKey?'prev':'next');}
          if(e.key==='ArrowRight'){e.preventDefault();move('cashStatus','right');}
          if(e.key==='ArrowLeft'){e.preventDefault();move('cashStatus','left');}
          if(e.key==='ArrowDown'){e.preventDefault();move('cashStatus','down');}
          if(e.key==='ArrowUp'){e.preventDefault();move('cashStatus','up');}
        }}>
          <div style={{display:'flex',alignItems:'center',gap:2}}>
            <CashToggle value={form.cashStatus||''} onUpdate={v=>{set('cashStatus',v);setTimeout(save,50);}}/>
            {form.squareTransactionId&&<span title="Matched via Square" style={{fontSize:8,color:'#065F46',lineHeight:1}}>💳</span>}
          </div>
        </td>
        {/* Err $ */}
        <td style={{padding:'2px 1px'}}>
          <input type="text" inputMode="decimal" value={form.paymentErrorLoss||''} placeholder="0"
            style={{...ns(30), borderColor: form.paymentErrorLoss?'#EF4444':'#E2E8F0', ...focusStyle('paymentErrorLoss')}}
            ref={el => refs.current.paymentErrorLoss = el}
            onChange={e => set('paymentErrorLoss', e.target.value)}
            onFocus={() => setFocused('paymentErrorLoss')}
            onBlur={() => { setFocused(null); save(); }}
            onKeyDown={e => onKey(e, 'paymentErrorLoss')}
          />
        </td>
        <td style={{padding:'2px 1px', minWidth:70}}>
          <input type="text" value={form.attn||''} placeholder="attn..."
            style={{...ts(68), fontSize:9, borderColor:form.attn?'#8B5CF6':'#E2E8F0', ...focusStyle('attn')}}
            ref={el => refs.current.attn = el}
            onChange={e => set('attn', e.target.value)}
            onFocus={() => setFocused('attn')}
            onBlur={() => { setFocused(null); save(); }}
            onKeyDown={e => onKey(e, 'attn')}
          />
        </td>
        <td style={{padding:'2px 1px', minWidth:70}}>
          <input type="text" value={form.notes||''} placeholder="notes..."
            style={{...ts(68), fontSize:9, ...focusStyle('notes')}}
            ref={el => refs.current.notes = el}
            onChange={e => set('notes', e.target.value)}
            onFocus={() => setFocused('notes')}
            onBlur={() => { setFocused(null); save(); }}
            onKeyDown={e => onKey(e, 'notes')}
          />
        </td>
        <td style={{padding:'2px 2px', whiteSpace:'nowrap'}}>
          <div style={{display:'flex', gap:2}}>
            <button onClick={()=>setShowMore(v=>!v)} style={{background:showMore?'#EFF6FF':'#F8FAFC',border:'1px solid #E2E8F0',borderRadius:3,padding:'2px 4px',cursor:'pointer',fontSize:9,color:'#64748B'}}>{showMore?'▲':'▾'}</button>
            <button onClick={onInsertAfter} title="Insert row below" style={{background:'#F0FDF4',border:'1px solid #BBF7D0',borderRadius:3,padding:'2px 4px',cursor:'pointer',color:'#16A34A',fontSize:9,fontWeight:700}}>+</button>
            <button onClick={()=>onDelete(form.id)} style={{background:'#FEE2E2',border:'none',borderRadius:3,padding:'2px 4px',cursor:'pointer',color:'#DC2626',fontSize:9}}>✕</button>
          </div>
        </td>
      </tr>
      {showMore&&(
        <tr style={{background:'#F8FAFC',borderBottom:'1px solid #F0F4F8'}}>
          <td style={{padding:'2px 3px',fontSize:8,color:'#94A3B8',fontStyle:'italic'}}>↳</td>
          <td colSpan={2} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>Other 2</span>
              {selInp('otherType2', otherItems, 68)}
            </label>
          </td>
          <td style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>$</span>
              {numInp('otherAmt2', 32)}
            </label>
          </td>
          <td colSpan={2} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>Discount $</span>
              {numInp('discountAmount', 46)}
            </label>
          </td>
          <td colSpan={3} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>Discount Type</span>
              <select value={form.discountType||''} onChange={e=>set('discountType',e.target.value)}
                style={{...ts(110),fontSize:9}}>
                <option value="">— type —</option>
                {DISCOUNT_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
              </select>
            </label>
          </td>
          <td colSpan={2} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>Nickname</span>
              <input type="text" value={form.nickname||''} onChange={e=>set('nickname',e.target.value)} onBlur={save}
                style={ts(70)} placeholder="goes by..."/>
            </label>
          </td>
          <td colSpan={3} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>Claim #</span>
              <input type="text" value={form.claimNumber||''} onChange={e=>set('claimNumber',e.target.value)} onBlur={save}
                style={ts(80)} placeholder="optional"/>
            </label>
          </td>
          <td colSpan={2} style={{padding:'2px 2px'}}>
            <label style={{display:'flex',flexDirection:'column',gap:1}}>
              <span style={{fontSize:7,color:'#94A3B8',textTransform:'uppercase'}}>P3</span>
              <div style={{display:'flex',gap:2}}>
                {selInp('payor3', null, 54, true)}
                {form.payor3&&<input type="text" inputMode="decimal" value={form.ins3Amt||''} placeholder="P3$"
                  style={{width:32,padding:'2px 3px',border:'1px solid #E2E8F0',borderRadius:3,fontSize:9,textAlign:'right',fontFamily:"'DM Sans',sans-serif"}}
                  onChange={e=>set('ins3Amt',e.target.value)} onBlur={save}
                />}
              </div>
            </label>
          </td>
        </tr>
      )}
    </>
  );
}

// ── Entry table ─────────────────────────────────────────────────────────────
function EntryTable({entries,lockedDate,lockedLoc,lockedDoctor,feeSettings,otherItems,onSave,onDelete,queue,onSwap,onRemove,onQueueCurrent,initCount,pasteNames,onPasteConsumed}) {
  const [rows, setRows] = useState([]);
  const rowsRef = useRef([]);
  // Keep rowsRef in sync
  useEffect(()=>{ rowsRef.current = rows; }, [rows]);
  // rowRefs[i] = { fieldName: domElement } — updated each render by each EntryRow
  const rowRefs = useRef([]);

  // Build a blank row using latest locked values (via ref to avoid stale closure)
  const lockedSnap = useRef({lockedDate, lockedLoc, lockedDoctor});
  lockedSnap.current = {lockedDate, lockedLoc, lockedDoctor};
  const blank = () => {
    const {lockedDate:d, lockedLoc:l, lockedDoctor:dr} = lockedSnap.current;
    return {...emptyEntry(dr,''), date:d, location:l||'', doctorId:dr||'', createdAt:new Date().toISOString()};
  };

  // Load existing entries when locked params change, or when entries for this date+loc update
  const relevantEntries = useMemo(() =>
    sortEntries(entries.filter(e =>
      e.date === lockedDate &&
      (!lockedLoc    || e.location  === lockedLoc) &&
      (!lockedDoctor || e.doctorId  === lockedDoctor)
    )),
  [entries, lockedDate, lockedLoc, lockedDoctor]); // eslint-disable-line

  const prevLockedKey = useRef('');
  useEffect(() => {
    const key = `${lockedDate}|${lockedLoc}|${lockedDoctor}`;
    const keyChanged = key !== prevLockedKey.current;
    prevLockedKey.current = key;
    // Always reload when navigation target changes
    // On same target, only reload if rows is empty (first load)
    if (keyChanged || rows.length === 0 || (rows.length === 1 && !rows[0].patientName)) {
      setRows(relevantEntries.length > 0 ? relevantEntries : [blank()]);
    }
  }, [lockedDate, lockedLoc, lockedDoctor, relevantEntries]); // eslint-disable-line

  // Open exact count of rows when initCount changes, optionally pre-filling names
  useEffect(() => {
    if (!initCount || initCount <= 0) return;
    const names = pasteNames || [];
    
    // Build all new rows synchronously
    const existingFilled = rowsRef.current ? rowsRef.current.filter(r => r.patientName?.trim()) : [];
    const newRows = [];
    let nameIdx = 0;
    
    while (existingFilled.length + newRows.length < initCount) {
      const r = blank();
      if (names[nameIdx] !== undefined) r.patientName = names[nameIdx];
      nameIdx++;
      newRows.push(r);
    }
    
    const allRows = [...existingFilled, ...newRows];
    setRows(allRows);
    
    // Save to Firestore immediately (not in a timeout)
    if (names.length > 0 && newRows.length > 0) {
      (async () => {
        for (const r of newRows) {
          if (r.patientName?.trim()) {
            try {
              await setDoc(doc(db, 'billingEntries', r.id), {
                ...r,
                updatedAt: new Date().toISOString()
              });
            } catch(e) { console.error('Save error:', e); }
          }
        }
        if (onPasteConsumed) onPasteConsumed();
      })();
    } else {
      if (onPasteConsumed) onPasteConsumed();
    }
  }, [initCount]); // eslint-disable-line

  const handleSave = (updated, signal) => {
    if (updated) {
      onSave(updated);
      setRows(prev => {
        const i = prev.findIndex(r => r.id === updated.id);
        if (i >= 0) { const n=[...prev]; n[i]=updated; return n; }
        return [...prev, updated];
      });
    }
    if (signal === 'newrow') {
      setRows(prev => [...prev, blank()]);
    }
  };

  const handleDelete = async (id) => {
    await deleteDoc(doc(db,'billingEntries',id));
    onDelete(id);
    setRows(prev => { const f=prev.filter(r=>r.id!==id); return f.length>0?f:[blank()]; });
  };

  const insertAfter = (idx) => {
    setRows(prev => { const r=[...prev]; r.splice(idx+1,0,blank()); return r; });
  };

  const appendRow = () => setRows(prev => [...prev, blank()]);

  const rowsLen = useRef(rows.length);
  rowsLen.current = rows.length;
  // Stable callbacks — never cause re-renders
  const handleRegisterRefs = useCallback((rowIdx, refsObj) => {
    rowRefs.current[rowIdx] = refsObj;
  }, []);
  const handleFocusRow = useCallback((rowIdx, field) => {
    if (rowIdx < 0 || rowIdx >= rowsLen.current) return;
    rowRefs.current[rowIdx]?.[field]?.focus();
  }, []);

  const hdrs = ['','Patient','Exam','CL','Optos','OCT','DFE','OV','Myopia','Other','Amt','Pt','Ins','Total','P1$','P1','P2$','P2','Cash','Err$','ATTN','Notes',''];

  return (
    <div>
      <div style={{overflowX:'auto'}}>
        <table style={{borderCollapse:'collapse', fontSize:10, width:'100%'}}>
          <thead>
            <tr style={{borderBottom:'2px solid #F0F4F8', background:'#F8FAFC'}}>
              {hdrs.map((h,i)=><th key={i} style={{padding:'3px 1px',textAlign:'left',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',whiteSpace:'nowrap'}}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
             {(()=>{
               const grouped = [];
               let lastLoc = null;
               rows.forEach((row, i) => {
                 const loc = row.location || lockedLoc || '';
                 if (loc && loc !== lastLoc) {
                   grouped.push({type:'header', loc, key:'hdr_'+loc+i});
                   lastLoc = loc;
                 }
                 grouped.push({type:'row', row, i});
               });
               return grouped.map(item => item.type === 'header' ? (
                 <tr key={item.key} style={{background:LOC_COLORS[item.loc]+'15'}}>
                   <td colSpan={25} style={{padding:'2px 8px',fontSize:10,fontWeight:700,color:LOC_COLORS[item.loc]||'#1B3A5C'}}>{item.loc}</td>
                 </tr>
               ) : (
                 <EntryRow key={item.row.id}
                   entry={item.row} rowIndex={item.i}
                   onSave={handleSave} onDelete={handleDelete}
                   feeSettings={feeSettings} otherItems={otherItems}
                   lockedLoc={lockedLoc} lockedDoctor={lockedDoctor}
                   registerRefs={handleRegisterRefs}
                   onFocusRow={handleFocusRow}
                   onInsertAfter={()=>insertAfter(item.i)}
                 />
               ));
             })()}
          </tbody>
        </table>
      </div>
      <div style={{display:'flex',gap:8,alignItems:'center',marginTop:6}}>
        <button onClick={appendRow}
          style={{background:'#F1F5F9',border:'1px dashed #CBD5E1',borderRadius:5,padding:'4px 14px',cursor:'pointer',fontSize:10,color:'#64748B',fontWeight:600}}>
          + Row
        </button>
        {onQueueCurrent&&<button onClick={onQueueCurrent}
          style={{background:'#FEF3C7',border:'1px solid #FCD34D',borderRadius:5,padding:'4px 10px',cursor:'pointer',fontSize:10,color:'#92400E',fontWeight:600,marginLeft:'auto'}}>
          Send to queue
        </button>}
      </div>
      {queue&&queue.length>0&&(
        <div style={{marginTop:10,background:'white',borderRadius:8,border:'1px solid #E2E8F0',overflow:'hidden'}}>
          <div style={{background:'#F8FAFC',padding:'5px 12px',borderBottom:'1px solid #F0F4F8',display:'flex',alignItems:'center',gap:8}}>
            <span style={{fontSize:11,fontWeight:700,color:'#64748B'}}>Queue</span>
            <span style={{fontSize:10,color:'#94A3B8'}}>{queue.length} waiting</span>
          </div>
          {queue.map((item,i)=>(
            <div key={item.key} style={{display:'flex',alignItems:'center',gap:8,padding:'5px 12px',borderBottom:'1px solid #F0F4F8'}}>
              <div style={{width:6,height:6,borderRadius:'50%',background:'#F59E0B'}}/>
              <span style={{fontSize:10,fontWeight:700,background:LOC_COLORS[item.loc]+'20',color:LOC_COLORS[item.loc],borderRadius:3,padding:'1px 5px'}}>{item.loc}</span>
              <span style={{fontSize:10,color:'#1B3A5C',fontWeight:600}}>{dayOfWeek(item.date,true)} {fmtDate(item.date)}</span>
              {item.doctor&&<span style={{fontSize:9,color:'#94A3B8'}}>Dr. {item.doctor}</span>}
              <div style={{marginLeft:'auto',display:'flex',gap:4}}>
                <button onClick={()=>onSwap(i)} style={{background:'#EFF6FF',border:'none',borderRadius:4,padding:'2px 8px',cursor:'pointer',fontSize:9,color:'#1D4ED8',fontWeight:600}}>Work on this</button>
                <button onClick={()=>onRemove(i)} style={{background:'#D1FAE5',border:'none',borderRadius:4,padding:'2px 8px',cursor:'pointer',fontSize:9,color:'#065F46',fontWeight:600}}>Done</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Resolve modal ───────────────────────────────────────────────────────────
function ResolveModal({entry,onSave,onClose}) {
  const [comment,setComment]=useState('');
  const [saving,setSaving]=useState(false);
  const save=async()=>{
    setSaving(true);
    const updated={...entry,attnResolved:true,attnComment:comment,updatedAt:new Date().toISOString()};
    await setDoc(doc(db,'billingEntries',entry.id),updated);
    if(comment.trim()) await createNotification(`ATTN resolved for ${entry.patientName} (${entry.date}, ${entry.location})`,entry.attn,comment);
    onSave(updated); setSaving(false);
  };
  return (
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2000,display:'flex',alignItems:'center',justifyContent:'center',padding:16}}>
      <div style={{background:'white',borderRadius:12,maxWidth:440,width:'100%',padding:22,boxShadow:'0 8px 32px rgba(0,0,0,0.2)'}}>
        <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',marginBottom:8}}>Resolve ATTN</h3>
        <p style={{fontSize:12,color:'#64748B',marginBottom:6}}>Original flag: <strong>{entry.attn}</strong></p>
        <p style={{fontSize:11,color:'#94A3B8',marginBottom:10}}>Add a comment to notify billing team (optional):</p>
        <textarea value={comment} onChange={e=>setComment(e.target.value)} rows={3}
          style={{width:'100%',padding:'8px 10px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:12,outline:'none',fontFamily:"'DM Sans',sans-serif",resize:'vertical',boxSizing:'border-box'}}
          placeholder="e.g. Insurance confirmed, patient called back..."/>
        <div style={{display:'flex',gap:8,marginTop:12}}>
          <button onClick={save} disabled={saving} style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'8px 18px',cursor:'pointer',fontSize:13,fontWeight:600}}>{saving?'Saving...':'Resolve'+(comment.trim()?' & Notify':'')}</button>
          <button onClick={onClose} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'8px 12px',cursor:'pointer',fontSize:13}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ── Edit modal ──────────────────────────────────────────────────────────────
function EditModal({entry,onSave,onClose,doctorLogs,feeSettings,otherItems}) {
  const [form,setForm]=useState({...entry});
  const [saving,setSaving]=useState(false);
  const [showRef,setShowRef]=useState(false);
  const [showResolve,setShowResolve]=useState(false);

  const set=(k,v)=>setForm(f=>{
    const next={...f,[k]:v};
    if(k==='exam'||k==='payor1'){const a=parseFloat(k==='exam'?v:next.exam)||0;const p=k==='payor1'?v:next.payor1;const b=feeSettings[next.location]||89;if(p==='Self'&&a>b){next.exam=String(b);next.cl=String((a-b).toFixed(2));}}
    return next;
  });

  const ptPaid=['exam','cl','optos','oct','dfe','ov'].reduce((s,k)=>s+(parseFloat(form[k])||0),0)+(parseFloat(form.otherAmt)||0)+(parseFloat(form.otherAmt2)||0)-(parseFloat(form.discountAmount)||0);
  const ins1=parseFloat(form.insurancePaid1)||0;
  const ins2=parseFloat(form.insurancePaid2)||0;
  const ins3=parseFloat(form.insurancePaid3)||0;
  const totalInsPaid=ins1+ins2+ins3;
  const insExp=parseFloat(form.ins)||0;
  const total=Math.max(0,ptPaid)+(totalInsPaid>0?totalInsPaid:insExp);
  const matchingLogs=useMemo(()=>(doctorLogs||[]).filter(l=>l.date===form.date&&(!form.location||l.location===form.location)&&(!form.doctorId||l.doctorId?.toLowerCase()===form.doctorId?.toLowerCase())),[doctorLogs,form]);

  const save=async()=>{
    setSaving(true);
    const u={...form,ptPaid:String(Math.max(0,ptPaid)),updatedAt:new Date().toISOString()};
    await setDoc(doc(db,'billingEntries',form.id),u);
    onSave(u); setSaving(false);
  };

  const inp={padding:'4px 6px',border:'1.5px solid #E2E8F0',borderRadius:6,fontSize:11,fontFamily:"'DM Sans',sans-serif",outline:'none',width:'100%',boxSizing:'border-box'};
  const sel={...inp,background:'white',cursor:'pointer'};
  const lbl=(text,color)=>(<span style={{fontSize:8,fontWeight:700,color:color||'#64748B',textTransform:'uppercase',whiteSpace:'nowrap'}}>{text}</span>);
  const col=(label,children,color)=>(<div style={{display:'flex',flexDirection:'column',gap:2}}>{lbl(label,color)}{children}</div>);

  return (
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center',padding:12,overflow:'auto'}}>
      <div style={{background:'white',borderRadius:14,maxWidth:680,width:'100%',maxHeight:'92vh',overflow:'auto',padding:18,boxShadow:'0 8px 40px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
          <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>Edit — {entry.patientName}</h3>
          <button onClick={onClose} style={{background:'none',border:'none',fontSize:18,cursor:'pointer',color:'#94A3B8'}}>✕</button>
        </div>

        {/* Row 1: identity */}
        <div style={{display:'grid',gridTemplateColumns:'110px 68px 95px 1fr 1fr',gap:6,marginBottom:8}}>
          {[['Date','date','date'],['Loc','location','loc'],['Dr','doctorId','dr'],['Patient','patientName','text'],['Nickname','nickname','text']].map(([label,key,type])=>(
            <label key={key} style={{display:'flex',flexDirection:'column',gap:2}}>
              {lbl(label)}
              {type==='loc'
                ?<select value={form.location||''} onChange={e=>set('location',e.target.value)} style={sel}><option value="">—</option>{LOCATIONS.map(l=><option key={l} value={l}>{l}</option>)}</select>
                :type==='dr'
                ?<select value={form.doctorId||''} onChange={e=>set('doctorId',e.target.value)} style={sel}><option value="">—</option>{DEFAULT_DOCTORS.map(d=><option key={d} value={d}>{d}</option>)}</select>
                :<input type={type} value={form[key]||''} onChange={e=>set(key,e.target.value)} style={inp} placeholder={key==='nickname'?'Goes by...':''}/> }
            </label>
          ))}
        </div>

        {/* Row 2: services + pt paid + cash */}
        {lbl('Services')}
        <div style={{display:'flex',gap:5,alignItems:'flex-end',marginTop:4,marginBottom:8,flexWrap:'wrap'}}>
          {['exam','cl','optos','oct','dfe','ov','myopia'].map(key=>(
            <div key={key} style={{display:'flex',flexDirection:'column',gap:2,minWidth:52}}>
              <span style={{fontSize:8,fontWeight:700,color:key==='myopia'?'#8B5CF6':'#475569',textTransform:'uppercase'}}>{key==='exam'?'Exam':key}{key==='myopia'?' ⚡':''}</span>
              <input type="text" inputMode="decimal" value={form[key]||''} onChange={e=>set(key,e.target.value)} style={{...inp,textAlign:'right',width:52}} placeholder="0"/>
            </div>
          ))}
          <div style={{display:'flex',flexDirection:'column',gap:2,minWidth:64}}>
            {lbl('Pt Paid','#065F46')}
            <div style={{padding:'4px 6px',background:'#F0FDF4',border:'1px solid #6EE7B7',borderRadius:6,fontSize:11,fontWeight:700,color:'#065F46',textAlign:'right',minWidth:64}}>{fmt$(Math.max(0,ptPaid))}</div>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Cash')}
            <CashToggle value={form.cashStatus||''} onUpdate={v=>set('cashStatus',v)}/>
          </div>
        </div>

        {/* Row 3: other services + discount on one line */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 52px 1fr 52px 1fr 64px',gap:6,marginBottom:8}}>
          {col('Other Service',<select value={form.otherType||''} onChange={e=>set('otherType',e.target.value)} style={sel}><option value="">None</option>{otherItems.map(i=><option key={i} value={i}>{i}</option>)}</select>)}
          {col('$',<input type="text" inputMode="decimal" value={form.otherAmt||''} onChange={e=>set('otherAmt',e.target.value)} style={{...inp,textAlign:'right'}} placeholder="0"/>)}
          {col('Other 2',<select value={form.otherType2||''} onChange={e=>set('otherType2',e.target.value)} style={sel}><option value="">None</option>{otherItems.map(i=><option key={i} value={i}>{i}</option>)}</select>)}
          {col('$',<input type="text" inputMode="decimal" value={form.otherAmt2||''} onChange={e=>set('otherAmt2',e.target.value)} style={{...inp,textAlign:'right'}} placeholder="0"/>)}
          {col('Discount Type',<select value={form.discountType||''} onChange={e=>set('discountType',e.target.value)} style={sel}><option value="">—</option>{DISCOUNT_TYPES.map(t=><option key={t}>{t}</option>)}</select>)}
          {col('Discount $',<input type="text" inputMode="decimal" value={form.discountAmount||''} onChange={e=>set('discountAmount',e.target.value)} style={{...inp,textAlign:'right'}} placeholder="0"/>)}
        </div>

        {/* Row 4: payor + exp$ per payor + ins exp total */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 56px 1fr 56px 1fr 56px 64px',gap:6,marginBottom:6}}>
          {[['Payor 1','payor1','ins1Amt'],['Payor 2','payor2','ins2Amt'],['Payor 3','payor3','ins3Amt']].map(([label,key,amtKey])=>(
            <React.Fragment key={key}>
              {col(label,<select value={form[key]||''} onChange={e=>set(key,e.target.value)} style={sel}><option value="">—</option>{PAYORS.map(p=><option key={p.value} value={p.value}>{p.label}</option>)}</select>)}
              {col('Exp $',<input type="text" inputMode="decimal" value={form[amtKey]||''} onChange={e=>set(amtKey,e.target.value)} style={{...inp,textAlign:'right',borderColor:form[key]?'#2E7D8C':'#E2E8F0'}} placeholder="0"/>)}
            </React.Fragment>
          ))}
          {col('Ins Exp',<input type="text" inputMode="decimal" value={form.ins||''} onChange={e=>set('ins',e.target.value)} style={{...inp,textAlign:'right'}}/>)}
        </div>

        {/* Row 5: ins paid green row */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr',gap:6,marginBottom:8,background:'#F0FDF4',borderRadius:8,padding:'8px 10px'}}>
          {[['Payor 1 Paid','insurancePaid1'],['Payor 2 Paid','insurancePaid2'],['Payor 3 Paid','insurancePaid3']].map(([label,key])=>(
            <label key={key} style={{display:'flex',flexDirection:'column',gap:2}}>
              {lbl(label,'#2E7D8C')}
              <input type="text" inputMode="decimal" value={form[key]||''} onChange={e=>set(key,e.target.value)} style={{...inp,textAlign:'right',borderColor:form[key]?'#2E7D8C':'#E2E8F0'}}/>
            </label>
          ))}
          <div style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Total Ins Paid','#065F46')}
            <div style={{padding:'4px 6px',background:'white',border:'1px solid #6EE7B7',borderRadius:6,fontSize:11,fontWeight:700,color:'#065F46',textAlign:'right'}}>{totalInsPaid>0?fmt$(totalInsPaid):'—'}</div>
          </div>
        </div>

        <div style={{fontSize:11,color:'#64748B',marginBottom:8}}>Total: <strong style={{color:'#1B3A5C'}}>{fmt$(total)}</strong></div>

        {/* Loss fields */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6,marginBottom:8,padding:'8px 10px',background:'#FEF2F2',borderRadius:8,border:'1px solid #FECACA'}}>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Payment Error Loss $','#EF4444')}
            <input type="text" inputMode="decimal" value={form.paymentErrorLoss||''} onChange={e=>set('paymentErrorLoss',e.target.value)} style={{...inp,textAlign:'right',borderColor:form.paymentErrorLoss?'#EF4444':'#E2E8F0'}} placeholder="0"/>
            <span style={{fontSize:8,color:'#94A3B8'}}>Forgotten payment, wrong charge, etc.</span>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Insurance Non-Payment Loss $','#EF4444')}
            <input type="text" inputMode="decimal" value={form.insuranceNonpaymentLoss||''} onChange={e=>set('insuranceNonpaymentLoss',e.target.value)} style={{...inp,textAlign:'right',borderColor:form.insuranceNonpaymentLoss?'#EF4444':'#E2E8F0'}} placeholder="0"/>
            <span style={{fontSize:8,color:'#94A3B8'}}>Insurance underpaid or denied.</span>
          </label>
        </div>

        {/* Claims / ATTN / Notes */}
        <div style={{display:'grid',gridTemplateColumns:'90px 90px 1fr 1fr',gap:6,marginBottom:10}}>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Claim #1')}
            <input value={form.claimNumber||''} onChange={e=>set('claimNumber',e.target.value)} style={inp}/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Claim #2')}
            <input value={form.claimNumber2||''} onChange={e=>set('claimNumber2',e.target.value)} style={inp}/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('ATTN','#8B5CF6')}
            <div style={{display:'flex',gap:4}}>
              <input value={form.attn||''} onChange={e=>set('attn',e.target.value)} style={{...inp,flex:1,borderColor:form.attn?'#8B5CF6':'#E2E8F0'}} placeholder="Flag..."/>
              {form.attn&&!form.attnResolved&&<button onClick={()=>setShowResolve(true)} style={{background:'#F5F3FF',border:'none',borderRadius:5,padding:'0 7px',cursor:'pointer',fontSize:9,fontWeight:700,color:'#6D28D9',whiteSpace:'nowrap'}}>Resolve</button>}
              {form.attnResolved&&<span style={{fontSize:8,color:'#10B981',fontWeight:700,whiteSpace:'nowrap'}}>✓ Resolved</span>}
            </div>
            {form.attnComment&&<p style={{fontSize:8,color:'#64748B',marginTop:2,fontStyle:'italic'}}>"{form.attnComment}"</p>}
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:2}}>
            {lbl('Notes')}
            <input value={form.notes||''} onChange={e=>set('notes',e.target.value)} style={inp} placeholder="e.g. pt used discount plan"/>
          </label>
        </div>

        {matchingLogs.length>0&&(
          <div style={{marginBottom:10}}>
            <button onClick={()=>setShowRef(v=>!v)} style={{background:'#F5F3FF',border:'1px solid #DDD6FE',borderRadius:6,padding:'3px 10px',cursor:'pointer',fontSize:11,color:'#6D28D9',fontWeight:600}}>Doctor log ({matchingLogs.length}) {showRef?'hide':'show'}</button>
            {showRef&&<div style={{marginTop:5,background:'#F5F3FF',borderRadius:8,padding:'7px 10px',border:'1px solid #DDD6FE'}}>
              {matchingLogs.map(l=><div key={l.id} style={{fontSize:10,marginBottom:3,padding:'3px 7px',background:'white',borderRadius:4}}>
                <strong>{l.patientName}</strong> — {l.payor1||'—'} — {['exam','cl','optos','oct','dfe','ov','myopia'].filter(k=>parseFloat(l[k]||l.routine)>0).map(k=>`${k}: ${fmt$(l[k]||l.routine)}`).join(', ')}
                {l.notes&&<span style={{color:'#94A3B8'}}> — "{l.notes}"</span>}
              </div>)}
            </div>}
          </div>
        )}
        <div style={{display:'flex',gap:8}}>
          <button onClick={save} disabled={saving} style={{background:saving?'#94A3B8':'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'7px 18px',cursor:'pointer',fontSize:12,fontWeight:600}}>{saving?'Saving...':'Save Changes'}</button>
          <button onClick={onClose} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'7px 12px',cursor:'pointer',fontSize:12}}>Cancel</button>
        </div>
        {showResolve&&<ResolveModal entry={form} onSave={updated=>{setForm(updated);onSave(updated);setShowResolve(false);}} onClose={()=>setShowResolve(false)}/>}
      </div>
    </div>
  );
}

// ── Calendar modal ──────────────────────────────────────────────────────────
function CalendarModal({onClose,doctorList}) {
  const [weekOffset,setWeekOffset]=useState(0);
  const [template,setTemplate]=useState({assignments:{},updatedAt:''});
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const cellRefs=useRef({});
  const [focused,setFocused]=useState(null); // `${di}-${li}`

  useEffect(()=>{ loadScheduleTemplate().then(t=>{setTemplate(t);setLoading(false);}); },[]);

  const ws=weekStart(weekOffset);
  const days=Array.from({length:7},(_,i)=>addDays(ws,i));

  // Base schedule: day-of-week x location defaults
  const getBaseDoc=(d,loc)=>{
    const dow=new Date(d+'T12:00:00').getDay(); // 0=Sun
    return template.base?.[`${dow}|${loc}`]||'';
  };
  const setBase=(dow,loc,val)=>setTemplate(prev=>({...prev,base:{...(prev.base||{}),[`${dow}|${loc}`]:val}}));
  const getA=(d,loc)=>template.assignments[d+'|'+loc]||getBaseDoc(d,loc);
  const setA=(d,loc,val)=>setTemplate(prev=>({...prev,assignments:{...prev.assignments,[d+'|'+loc]:val}}));
  const save=async()=>{ setSaving(true); await saveScheduleTemplate(template); setSaving(false); onClose(); };

  const handleCellKey=(e,di,li)=>{
    const numDays=days.length; const numLocs=LOCATIONS.length;
    if(e.key==='Tab'){e.preventDefault();const nextDi=(di+1)%numDays;const nextLi=di===numDays-1?Math.min(li+1,numLocs-1):li;cellRefs.current[`${nextDi}-${nextLi}`]?.focus();}
    if(e.key==='ArrowRight'){e.preventDefault();cellRefs.current[`${Math.min(di+1,numDays-1)}-${li}`]?.focus();}
    if(e.key==='ArrowLeft'){e.preventDefault();cellRefs.current[`${Math.max(di-1,0)}-${li}`]?.focus();}
    if(e.key==='ArrowDown'){e.preventDefault();cellRefs.current[`${di}-${Math.min(li+1,numLocs-1)}`]?.focus();}
    if(e.key==='ArrowUp'){e.preventDefault();cellRefs.current[`${di}-${Math.max(li-1,0)}`]?.focus();}
  };

  return (
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center',padding:16,overflow:'auto'}}>
      <style>{`.cal-cell:focus{outline:2px solid #2E7D8C!important;background:#EFF6FF!important;}`}</style>
      <div style={{background:'white',borderRadius:14,maxWidth:960,width:'100%',maxHeight:'92vh',overflow:'auto',padding:22,boxShadow:'0 8px 40px rgba(0,0,0,0.2)'}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:14}}>
          <div>
            <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:'#1B3A5C'}}>Doctor Schedule</h3>
            {template.updatedAt&&<p style={{fontSize:10,color:'#94A3B8'}}>Last updated: {new Date(template.updatedAt).toLocaleString()}</p>}
          </div>
          <button onClick={onClose} style={{background:'none',border:'none',fontSize:20,cursor:'pointer',color:'#94A3B8'}}>✕</button>
        </div>
        <div style={{display:'flex',gap:8,marginBottom:12,alignItems:'center',flexWrap:'wrap'}}>
          <button onClick={()=>setWeekOffset(v=>v-1)} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'5px 12px',cursor:'pointer',fontSize:11}}>← Prev</button>
          <span style={{fontSize:12,fontWeight:700,color:'#1B3A5C'}}>{fmtDate(ws)} — {fmtDate(addDays(ws,6))}</span>
          <button onClick={()=>setWeekOffset(v=>v+1)} style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'5px 12px',cursor:'pointer',fontSize:11}}>Next →</button>
          <button onClick={()=>setWeekOffset(0)} style={{background:'#EFF6FF',color:'#1D4ED8',border:'none',borderRadius:6,padding:'5px 10px',cursor:'pointer',fontSize:11,fontWeight:600}}>This Week</button>
          <span style={{fontSize:10,color:'#94A3B8'}}>Tab / arrows to navigate · type to select doctor</span>
        </div>
        {loading?<p style={{color:'#94A3B8'}}>Loading...</p>:(
          <div style={{overflowX:'auto'}}>
            <table style={{borderCollapse:'collapse',width:'100%',fontSize:11}}>
              <thead>
                <tr style={{background:'#F8FAFC'}}>
                  <th style={{padding:'7px 10px',fontSize:10,color:'#94A3B8',textTransform:'uppercase',fontWeight:700,textAlign:'left',width:80}}>Loc</th>
                  {days.map((d,di)=>(
                    <th key={d} style={{padding:'7px 6px',fontSize:10,color:'#1B3A5C',fontWeight:700,textAlign:'center',minWidth:90,background:d===today()?'#EFF6FF':'#F8FAFC'}}>
                      <div>{DAY_ABBR[new Date(d+'T12:00:00').getDay()]}</div>
                      <div style={{fontSize:9,color:'#94A3B8'}}>{fmtDate(d)}</div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {LOCATIONS.map((loc,li)=>(
                  <tr key={loc} style={{borderBottom:'1px solid #F0F4F8'}}>
                    <td style={{padding:'5px 10px'}}>
                      <span style={{fontWeight:700,fontSize:11,background:LOC_COLORS[loc]+'20',color:LOC_COLORS[loc],borderRadius:4,padding:'2px 8px'}}>{loc}</span>
                    </td>
                    {days.map((d,di)=>{
                      const val=getA(d,loc);
                      const isClosed=val==='CLOSED';
                      const cellKey=`${di}-${li}`;
                      const isFocused=focused===cellKey;
                      return (
                        <td key={d} style={{padding:'3px 3px',textAlign:'center',background:isFocused?'#EFF6FF':d===today()?'#F0F9FF':isClosed?'#F8FAFC':'white'}}>
                          <select value={val} onChange={e=>setA(d,loc,e.target.value)}
                            ref={el=>cellRefs.current[cellKey]=el}
                            onKeyDown={e=>handleCellKey(e,di,li)}
                            onFocus={()=>setFocused(cellKey)}
                            onBlur={()=>setFocused(null)}
                            className="cal-cell"
                            style={{width:'100%',padding:'4px 2px',border:isFocused?'2px solid #2E7D8C':'1px solid #E2E8F0',borderRadius:5,fontSize:10,outline:'none',background:isFocused?'#EFF6FF':isClosed?'#F1F5F9':'white',color:isClosed?'#94A3B8':'#1B3A5C'}}>
                            <option value="">— open —</option>
                            <option value="CLOSED">Closed</option>
                            {[...doctorList].sort().map(d=><option key={d} value={d}>{d}</option>)}
                          </select>
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {/* Base schedule */}
        <details style={{marginTop:14,border:'1px solid #E2E8F0',borderRadius:8,padding:'0'}}>
          <summary style={{padding:'8px 14px',cursor:'pointer',fontSize:12,fontWeight:700,color:'#1B3A5C',background:'#F8FAFC',borderRadius:8}}>
            Base Schedule (auto-fill defaults by day of week)
          </summary>
          <div style={{padding:'12px 14px'}}>
            <p style={{fontSize:11,color:'#94A3B8',marginBottom:10}}>Set which doctor normally works each location on each day of the week. These pre-fill any week you view.</p>
            <div style={{overflowX:'auto'}}>
              <table style={{borderCollapse:'collapse',fontSize:10}}>
                <thead>
                  <tr style={{background:'#F8FAFC'}}>
                    <th style={{padding:'5px 8px',textAlign:'left',fontSize:9,color:'#94A3B8',fontWeight:700}}>Loc</th>
                    {['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map((d,i)=>(
                      <th key={i} style={{padding:'5px 8px',textAlign:'center',fontSize:9,color:'#1B3A5C',fontWeight:700,minWidth:75}}>{d}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {LOCATIONS.map(loc=>(
                    <tr key={loc} style={{borderBottom:'1px solid #F0F4F8'}}>
                      <td style={{padding:'4px 8px'}}>
                        <span style={{fontWeight:700,fontSize:10,background:LOC_COLORS[loc]+'20',color:LOC_COLORS[loc],borderRadius:4,padding:'1px 7px'}}>{loc}</span>
                      </td>
                      {[0,1,2,3,4,5,6].map(dow=>(
                        <td key={dow} style={{padding:'3px 4px'}}>
                          <select value={(template.base||{})[`${dow}|${loc}`]||''}
                            onChange={e=>setBase(dow,loc,e.target.value)}
                            style={{width:'100%',padding:'3px 2px',border:'1px solid #E2E8F0',borderRadius:4,fontSize:9,outline:'none',background:'white'}}>
                            <option value="">—</option>
                            <option value="CLOSED">Closed</option>
                            {[...doctorList].sort().map(d=><option key={d} value={d}>{d}</option>)}
                          </select>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </details>
        <div style={{display:'flex',gap:8,marginTop:12}}>
          <button onClick={save} disabled={saving} style={{background:saving?'#94A3B8':'#10B981',color:'white',border:'none',borderRadius:7,padding:'8px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>{saving?'Saving...':'Save Schedule'}</button>
          <button onClick={onClose} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'8px 12px',cursor:'pointer',fontSize:13}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ── Square payment panel ─────────────────────────────────────────────────────
const SQUARE_WORKER = 'https://spark-square.ilikebroccoli.workers.dev';

const DEVICE_TO_LOC = {
  'santa clara ipad': 'SC',
  'fremont ipad':     'F',
  'walnut creek ipad 2.0': 'WC',
  'sunnyvale ipad':   'SV',
};

function deviceToLoc(deviceName) {
  if (!deviceName) return null;
  return DEVICE_TO_LOC[deviceName.toLowerCase().trim()] || null;
}

function fmtTime(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleTimeString('en-US', {hour:'numeric', minute:'2-digit', hour12:true, timeZone:'America/Los_Angeles'});
}

function squareNameSim(note, patientName) {
  if (!note || !patientName) return 0;
  const n = note.toUpperCase().replace(/[^A-Z\s\n]/g,'').trim();
  const p = patientName.toUpperCase().replace(/[^A-Z\s]/g,'').trim();

  // note might have multiple names separated by newline — check each line
  const noteLines = n.split(/[\n,]+/).map(l=>l.trim()).filter(Boolean);

  let best = 0;
  for (const line of noteLines) {
    const noteToks = line.split(/\s+/).filter(t=>t.length>=2);
    const nameToks = p.split(/\s+/).filter(t=>t.length>=2);
    if (!noteToks.length || !nameToks.length) continue;

    // Count how many name tokens appear in note tokens (substring match)
    const hits = nameToks.filter(nt => noteToks.some(no => no.includes(nt) || nt.includes(no))).length;
    const sim = hits / nameToks.length;
    if (sim > best) best = sim;
  }
  return best;
}

function matchSquareToEntries(txn, entries) {
  const amt = (txn.amount_money?.amount || 0) / 100;
  const note = txn.note || '';
  const results = [];

  // Detect if note looks like a service description rather than a patient name
  // (contains words like "exam", "copay", "myopia", "contact", "optomap", "pay")
  const svcKeywords = /\b(exam|copay|myopia|contact|optomap|general|private|control|lens|pay)\b/i;
  const noteIsServiceDesc = svcKeywords.test(note);

  for (const e of entries) {
    const alreadyLinked = e.squareTransactionId === txn.id;
    if (alreadyLinked) return [{entry: e, score: 2.0, amtMatch: true, nameSim: 1, alreadyLinked: true}];

    const ptPaid = parseFloat(e.ptPaid) || 0;
    const svcTotal = ['exam','cl','optos','oct','dfe','ov','myopia','otherAmt','otherAmt2']
      .reduce((s,k) => s + (parseFloat(e[k])||0), 0);
    const disc = parseFloat(e.discountAmount)||0;
    const expectedPt = Math.max(0, svcTotal - disc);

    const amtMatchPaid   = amt > 0 && ptPaid > 0 && Math.abs(amt - ptPaid) < 1.0;
    const amtMatchExpect = amt > 0 && expectedPt > 0 && Math.abs(amt - expectedPt) < 1.0;
    const amtMatch       = amtMatchPaid || amtMatchExpect;
    const isZero         = amt === 0;

    // Name match — skip if note is a service description
    const nameSim = (!noteIsServiceDesc && note) ? squareNameSim(note, e.patientName) : 0;

    // Scoring: name match is primary signal, amount is secondary
    // $0 transactions: match by name only
    const score = isZero
      ? nameSim
      : (nameSim * 0.55) + (amtMatch ? 0.45 : 0);

    if (nameSim >= 0.4 || (amtMatch && score > 0.2)) {
      results.push({entry: e, score, amtMatch, nameSim, alreadyLinked: false});
    }
  }
  return results.sort((a,b) => b.score - a.score);
}

function SquarePanel({date, loc, entries, onEntryUpdated}) {
  const [txns, setTxns]           = React.useState(null);
  const [loading, setLoading]     = React.useState(false);
  const [error, setError]         = React.useState('');
  const [approving, setApproving] = React.useState({});
  const [overrides, setOverrides] = React.useState({});
  const [open, setOpen]           = React.useState(false);

  const load = async () => {
    if (loading) return;
    setLoading(true); setError('');
    try {
      const beginTime = `${date}T00:00:00-07:00`;
      const endTime   = `${date}T23:59:59-07:00`;
      const res = await fetch(SQUARE_WORKER, {
        method: 'POST',
        headers: {'Content-Type':'application/json'},
        body: JSON.stringify({begin_time: beginTime, end_time: endTime})
      });
      if (!res.ok) throw new Error(`Worker error ${res.status}`);
      const data = await res.json();
      const payments = (data.payments || [])
        .filter(p => {
          if (p.status !== 'COMPLETED') return false;
          const pLoc = deviceToLoc(p.device_details?.device_name);
          return pLoc === loc;
        })
        // Sort chronologically — earliest first, matching entry order
        .sort((a,b) => (a.created_at||'').localeCompare(b.created_at||''));
      setTxns(payments);
    } catch(e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const approve = async (txn, entry) => {
    const key = txn.id;
    setApproving(p => ({...p, [key]: true}));
    try {
      const amt = (txn.amount_money?.amount || 0) / 100;
      const {setDoc, doc: fsDoc} = await import('firebase/firestore');
      const upd = {
        ...entry,
        ptPaid: String(amt.toFixed(2)),
        cashStatus: 'received',
        squareTransactionId: txn.id,
        squareApprovedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      await setDoc(fsDoc(db, 'billingEntries', entry.id), upd);
      onEntryUpdated(upd);
      setTxns(prev => prev ? [...prev] : prev);
    } catch(e) {
      alert('Error saving: ' + e.message);
    } finally {
      setApproving(p => ({...p, [key]: false}));
    }
  };

  // Sort entries same way as table — by createdAt
  const locEntries = [...entries].sort((a,b)=>(a.createdAt||'').localeCompare(b.createdAt||''));

  const getTxnNote = (txn) =>
    txn.note ||
    txn.tender?.[0]?.note ||
    txn.order?.line_items?.[0]?.note ||
    '';

  const getSvcBreakdown = (entry) => {
    const svcs = [
      ['Exam', 'exam'], ['CL', 'cl'], ['Optos', 'optos'], ['OCT', 'oct'],
      ['DFE', 'dfe'], ['OV', 'ov'], ['Myopia', 'myopia'],
      ['Other', 'otherAmt'], ['Other2', 'otherAmt2']
    ].filter(([,k]) => parseFloat(entry[k]) > 0)
     .map(([label,k]) => `${label} ${fmt$(parseFloat(entry[k]))}`);
    const disc = parseFloat(entry.discountAmount);
    if (disc > 0) svcs.push(`Disc -${fmt$(disc)}`);
    return svcs.join(' · ');
  };

  if (!open) return (
    <button onClick={()=>{setOpen(true);load();}}
      style={{marginBottom:8,background:'#F0FDF4',border:'1px solid #6EE7B7',borderRadius:5,padding:'4px 12px',cursor:'pointer',fontSize:10,color:'#065F46',fontWeight:700,display:'inline-flex',alignItems:'center',gap:4}}>
      💳 Square Payments
    </button>
  );

  return (
    <div style={{marginBottom:10,background:'#F8FFFE',border:'1px solid #6EE7B7',borderRadius:8,padding:'10px 12px'}}>
      <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:8}}>
        <span style={{fontSize:11,fontWeight:700,color:'#065F46'}}>💳 Square Payments — {date} {loc}</span>
        <button onClick={load} disabled={loading}
          style={{background:'#F0FDF4',border:'1px solid #6EE7B7',borderRadius:4,padding:'2px 8px',cursor:'pointer',fontSize:9,color:'#065F46'}}>
          {loading ? '...' : '↺ Refresh'}
        </button>
        <button onClick={()=>setOpen(false)}
          style={{marginLeft:'auto',background:'none',border:'none',cursor:'pointer',fontSize:12,color:'#94A3B8'}}>✕</button>
      </div>

      {error && <div style={{fontSize:10,color:'#DC2626',marginBottom:6}}>Error: {error}</div>}
      {loading && <div style={{fontSize:10,color:'#94A3B8'}}>Loading...</div>}

      {txns && txns.length === 0 && (
        <div style={{fontSize:10,color:'#94A3B8'}}>No completed Square transactions for {loc} on {date}.</div>
      )}

      {txns && txns.map(txn => {
        const amt = (txn.amount_money?.amount || 0) / 100;
        const note = getTxnNote(txn);
        const time = fmtTime(txn.created_at);
        const overrideId = overrides[txn.id];
        const matches = matchSquareToEntries({...txn, note}, locEntries);
        const best = overrideId
          ? locEntries.find(e => e.id === overrideId)
          : matches[0]?.entry;
        const bestScore = matches[0]?.score || 0;
        const alreadyLinked = best?.squareTransactionId === txn.id;
        const isApproving = approving[txn.id];

        const confidence = alreadyLinked ? 'linked'
          : bestScore >= 0.9 ? 'high'
          : bestScore >= 0.5 ? 'med'
          : best ? 'low' : 'none';

        const confColor = {linked:'#065F46',high:'#16A34A',med:'#92400E',low:'#64748B',none:'#DC2626'}[confidence];
        const confBg    = {linked:'#F0FDF4',high:'#DCFCE7',med:'#FEF3C7',low:'#F8FAFC',none:'#FEF2F2'}[confidence];
        const confLabel = {linked:'✓ Linked',high:'Strong',med:'Possible',low:'Weak',none:'No match'}[confidence];

        return (
          <div key={txn.id} style={{marginBottom:6,padding:'7px 10px',background:confBg,borderRadius:6,border:`1px solid ${confColor}40`}}>
            {/* Transaction header */}
            <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:4}}>
              <span style={{fontSize:11,fontWeight:700,color:'#1B3A5C'}}>{fmt$(amt)}</span>
              <span style={{fontSize:9,color:'#94A3B8'}}>{time}</span>
              {note && <span style={{fontSize:9,color:'#475569',fontStyle:'italic',flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>"{note}"</span>}
              <span style={{fontSize:7,fontWeight:700,color:confColor,background:'white',border:`1px solid ${confColor}`,borderRadius:3,padding:'1px 4px',whiteSpace:'nowrap'}}>{confLabel}</span>
            </div>

            {/* Match row */}
            <div style={{display:'flex',alignItems:'flex-start',gap:4,flexDirection:'column'}}>
              <div style={{display:'flex',alignItems:'center',gap:4,width:'100%'}}>
                {best || overrideId ? (
                  <>
                    <select
                      value={overrideId || best?.id || ''}
                      onChange={e => setOverrides(p => ({...p, [txn.id]: e.target.value}))}
                      style={{flex:1,fontSize:9,padding:'2px 4px',border:'1px solid #E2E8F0',borderRadius:4,background:'white',fontFamily:"'DM Sans',sans-serif"}}>
                      {[...(best ? [best] : []), ...locEntries.filter(e=>e.id!==best?.id)].map(e => (
                        <option key={e.id} value={e.id}>{e.patientName}</option>
                      ))}
                    </select>
                    {alreadyLinked ? (
                      <span style={{fontSize:9,color:'#065F46',fontWeight:700,whiteSpace:'nowrap'}}>✓ Done</span>
                    ) : (
                      <button
                        onClick={() => approve(txn, locEntries.find(e=>e.id===(overrideId||best?.id))||best)}
                        disabled={isApproving}
                        style={{background:isApproving?'#94A3B8':'#065F46',color:'white',border:'none',borderRadius:4,padding:'3px 10px',cursor:'pointer',fontSize:9,fontWeight:700,whiteSpace:'nowrap'}}>
                        {isApproving ? '...' : '✓ Write'}
                      </button>
                    )}
                  </>
                ) : (
                  <>
                    <select
                      value={overrideId || ''}
                      onChange={e => setOverrides(p => ({...p, [txn.id]: e.target.value}))}
                      style={{flex:1,fontSize:9,padding:'2px 4px',border:'1px solid #E2E8F0',borderRadius:4,background:'white',fontFamily:"'DM Sans',sans-serif"}}>
                      <option value="">— Select patient —</option>
                      {locEntries.map(e => (
                        <option key={e.id} value={e.id}>{e.patientName}</option>
                      ))}
                    </select>
                    {overrideId && (
                      <button
                        onClick={() => approve(txn, locEntries.find(e=>e.id===overrideId))}
                        disabled={isApproving}
                        style={{background:isApproving?'#94A3B8':'#065F46',color:'white',border:'none',borderRadius:4,padding:'3px 10px',cursor:'pointer',fontSize:9,fontWeight:700,whiteSpace:'nowrap'}}>
                        {isApproving ? '...' : '✓ Write'}
                      </button>
                    )}
                  </>
                )}
              </div>
              {/* Service breakdown for matched entry */}
              {(best || (overrideId && locEntries.find(e=>e.id===overrideId))) && (() => {
                const e = locEntries.find(e=>e.id===(overrideId||best?.id)) || best;
                const breakdown = e ? getSvcBreakdown(e) : '';
                return breakdown ? (
                  <div style={{fontSize:8,color:'#64748B',paddingLeft:2}}>{breakdown}</div>
                ) : null;
              })()}
            </div>

            {/* Split tender */}
            {txn.tender && txn.tender.length > 1 && (
              <div style={{fontSize:8,color:'#92400E',marginTop:3}}>
                ⚠ Split: {txn.tender.map(t=>`${t.type} ${fmt$((t.amount_money?.amount||0)/100)}`).join(' + ')}
              </div>
            )}
          </div>
        );
      })}

    </div>
  );
}

// ── Days tab ────────────────────────────────────────────────────────────────
function DaysTab({entries,onStartDay,onViewDay,onAddToQueue,scheduleTemplate,onEntryUpdated}) {
  const [filter,setFilter]=useState('thisweek');
  const [closedDays,setClosedDays]=useState({});
  const [expanded,setExpanded]=useState({});
  const [expandedDay,setExpandedDay]=useState(null);

  useEffect(()=>{ getDoc(doc(db,'billingSettings','closedDays')).then(s=>{if(s.exists())setClosedDays(s.data());}); },[]);

  const toggleClosed=async(key,cur)=>{
    if(!cur){if(!window.confirm('Mark this day as closed?'))return;}
    const u={...closedDays,[key]:!cur}; setClosedDays(u);
    await setDoc(doc(db,'billingSettings','closedDays'),u);
  };

  const entryIndex=useMemo(()=>{
    const idx={};
    for(const e of entries){const key=e.date+'|'+e.location;if(!idx[key])idx[key]={entries:[],doctors:new Set()};idx[key].entries.push(e);if(e.doctorId)idx[key].doctors.add(e.doctorId);}
    return idx;
  },[entries]);

  const getDates=()=>{
    const now=new Date();
    if(filter==='thisweek'||filter==='lastweek'){
      const off=filter==='lastweek'?-7:0;
      const start=new Date(now);start.setDate(start.getDate()-start.getDay()+off);
      return Array.from({length:7},(_,i)=>{const d=new Date(start);d.setDate(d.getDate()+i);return d.toISOString().slice(0,10);}).sort((a,b)=>b.localeCompare(a));
    }
    if(filter==='thismonth'){const y=now.getFullYear(),m=now.getMonth();const days=new Date(y,m+1,0).getDate();return Array.from({length:days},(_,i)=>`${y}-${String(m+1).padStart(2,'0')}-${String(i+1).padStart(2,'0')}`).sort((a,b)=>b.localeCompare(a));}
    const s=new Set(entries.map(e=>e.date));const d=new Date(now);
    for(let i=0;i<60;i++){s.add(d.toISOString().slice(0,10));d.setDate(d.getDate()-1);}
    return [...s].sort((a,b)=>b.localeCompare(a));
  };
  const dates=getDates();

  const grouped=useMemo(()=>{
    if(filter!=='all')return null;
    const months={};
    for(const d of dates){
      const dt=new Date(d+'T12:00:00');
      const mKey=`${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}`;
      const mLabel=`${MONTH_NAMES[dt.getMonth()]} ${dt.getFullYear()}`;
      const wi=getFiscalWeekInfo(d);const wNum=wi.week;const wKey=`${mKey}-W${String(wNum).padStart(2,'0')}`;
      if(!months[mKey])months[mKey]={label:mLabel,weeks:{}};
      if(!months[mKey].weeks[wKey])months[mKey].weeks[wKey]={label:getWeekLabel(d),dates:[]};
      if(!months[mKey].weeks[wKey].dates.includes(d))months[mKey].weeks[wKey].dates.push(d);
    }
    return months;
  },[dates,filter]);

  const stCfg={closed:{bg:'#F8FAFC',color:'#94A3B8',label:'Closed',dot:'#CBD5E1'},empty:{bg:'#F8FAFC',color:'#64748B',label:'Not started',dot:'#E2E8F0'},flagged:{bg:'#FFF7ED',color:'#DC2626',label:'Flagged',dot:'#EF4444'},pending:{bg:'white',color:'#92400E',label:'Pending',dot:'#F59E0B'},done:{bg:'#F0FDF4',color:'#065F46',label:'Complete',dot:'#10B981'}};
  const getStatus=(key,isClosed)=>{if(isClosed)return'closed';const data=entryIndex[key];if(!data||!data.entries.length)return'empty';if(data.entries.some(e=>computeStatus(e)==='flagged'))return'flagged';if(data.entries.some(e=>computeStatus(e)==='pending'))return'pending';return'done';};

  const DayRow=({date,loc})=>{
    const key=date+'|'+loc;const isClosed=!!closedDays[key];const data=entryIndex[key];
    const status=getStatus(key,isClosed);const cfg=stCfg[status];
    const doctors=data?[...data.doctors].join(', '):'';
    // Check specific assignment first, then base schedule
    const specificDoc = scheduleTemplate?.assignments?.[key]||'';
    const dowIdx = new Date(date+'T12:00:00').getDay(); // 0=Sun
    const baseDoc = scheduleTemplate?.base?.[`${dowIdx}|${loc}`]||'';
    const scheduledDoc = specificDoc || baseDoc;
    const flagged=data?.entries.filter(e=>computeStatus(e)==='flagged').length||0;
    const pending=data?.entries.filter(e=>computeStatus(e)==='pending').length||0;
    const isOpen=expandedDay===key;
    return (
      <div style={{marginBottom:1}}>
        <div style={{background:cfg.bg,borderRadius:isOpen?'5px 5px 0 0':5,padding:'3px 10px',border:'1px solid #F0F4F8',display:'flex',gap:6,alignItems:'center'}}>
          <div style={{width:6,height:6,borderRadius:'50%',background:cfg.dot,flexShrink:0}}/>
          <span style={{fontSize:9,fontWeight:700,background:LOC_COLORS[loc]+'20',color:LOC_COLORS[loc],borderRadius:3,padding:'1px 4px',minWidth:20,textAlign:'center',flexShrink:0}}>{loc}</span>
          <span style={{fontSize:9,color:cfg.color,minWidth:64,flexShrink:0}}>{cfg.label}{data?.entries.length>0?` (${data.entries.length})`:''}</span>
          {(doctors||scheduledDoc)&&<span style={{fontSize:8,color:'#94A3B8',flexShrink:0}}>{doctors||scheduledDoc}</span>}
          {flagged>0&&<span style={{fontSize:8,fontWeight:700,color:'#DC2626',background:'#FEE2E2',borderRadius:3,padding:'1px 3px'}}>⚠{flagged}</span>}
          {pending>0&&!flagged&&<span style={{fontSize:8,color:'#92400E',background:'#FEF3C7',borderRadius:3,padding:'1px 3px'}}>{pending}p</span>}
          <div style={{marginLeft:'auto',display:'flex',gap:3}}>
            {!isClosed&&data?.entries.length>0&&<button onClick={()=>setExpandedDay(isOpen?null:key)} style={{background:'#EFF6FF',border:'none',borderRadius:3,padding:'2px 6px',cursor:'pointer',fontSize:8,color:'#1D4ED8',fontWeight:600}}>{isOpen?'Hide':'View'}</button>}
            {!isClosed&&<button onClick={()=>onStartDay(date,loc,scheduledDoc)} style={{background:'#1B3A5C',border:'none',borderRadius:3,padding:'2px 6px',cursor:'pointer',fontSize:8,color:'white',fontWeight:600}}>{data?.entries.length>0?'+':'Start'}</button>}
            {!isClosed&&onAddToQueue&&<button onClick={()=>onAddToQueue(date,loc,scheduledDoc)} style={{background:'#FEF3C7',border:'none',borderRadius:3,padding:'2px 4px',cursor:'pointer',fontSize:8,color:'#92400E'}}>+Q</button>}
            <button onClick={()=>toggleClosed(key,isClosed)} style={{background:isClosed?'#F1F5F9':'#FEF3C7',border:'none',borderRadius:3,padding:'2px 4px',cursor:'pointer',fontSize:8,color:isClosed?'#64748B':'#92400E'}}>{isClosed?'Reopen':'Closed'}</button>
          </div>
        </div>
        {isOpen&&data?.entries&&(
          <div style={{background:'white',border:'1px solid #E2E8F0',borderTop:'none',borderRadius:'0 0 5px 5px',padding:'6px 10px'}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:5}}>
              <span style={{fontSize:9,fontWeight:700,color:'#64748B'}}>{data.entries.length} patients</span>
              <button onClick={()=>onStartDay(date,loc,scheduledDoc)} style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:4,padding:'2px 8px',cursor:'pointer',fontSize:8,fontWeight:600}}>Edit / Add</button>
            </div>
            {/* Compact inline table — minimum width */}
            <table style={{borderCollapse:'collapse',fontSize:9}}>
              <thead>
                <tr style={{borderBottom:'1px solid #F0F4F8'}}>
                  {['Name','Services','Disc','Pt Paid','Ins','Total','P1','P2'].map(h=><th key={h} style={{padding:'2px 5px',textAlign:'left',fontWeight:700,color:'#94A3B8',fontSize:8,whiteSpace:'nowrap'}}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {sortEntries(data.entries).map(e=>{
                  const ptP=parseFloat(e.ptPaid)||0;const ins=parseFloat(e.ins)||0;const insPaid=parseFloat(e.insurancePaid1)||0;const total=ptP+(insPaid>0?insPaid:ins);
                  const svcs=['exam','cl','optos','oct','dfe','ov','myopia'].filter(k=>parseFloat(e[k]||e.routine)>0).map(k=>`${k==='exam'?'Ex':k.toUpperCase()}`).join(' ');
                  return (
                    <tr key={e.id} style={{borderBottom:'1px solid #F8FAFC'}}>
                      <td style={{padding:'2px 5px',fontWeight:600,whiteSpace:'nowrap'}}>{e.patientName}</td>
                      <td style={{padding:'2px 5px',color:'#64748B',whiteSpace:'nowrap'}}>{svcs}</td>
                      <td style={{padding:'2px 5px',color:'#92400E'}}>{e.discountAmount&&parseFloat(e.discountAmount)>0?'%':''}</td>
                      <td style={{padding:'2px 5px',color:'#16A34A',textAlign:'right',whiteSpace:'nowrap'}}>{fmt$(ptP)}</td>
                      <td style={{padding:'2px 5px',color:'#1D4ED8',textAlign:'right',whiteSpace:'nowrap'}}>{fmt$(ins)}</td>
                      <td style={{padding:'2px 5px',fontWeight:600,textAlign:'right',whiteSpace:'nowrap'}}>{total===0?'$0.00':fmt$(total)||'—'}</td>
                      <td style={{padding:'2px 5px',color:'#64748B',whiteSpace:'nowrap'}}>{e.payor1||'—'}</td>
                      <td style={{padding:'2px 5px',color:'#94A3B8',whiteSpace:'nowrap'}}>{e.payor2||''}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                {(()=>{
                  const totPt=data.entries.reduce((s,e)=>s+(parseFloat(e.ptPaid)||0),0);
                  const totIns=data.entries.reduce((s,e)=>s+(parseFloat(e.ins)||0),0);
                  const totInsPaid=data.entries.reduce((s,e)=>s+(parseFloat(e.insurancePaid1)||0),0);
                  const totTotal=data.entries.reduce((s,e)=>{const p=parseFloat(e.ptPaid)||0;const ip=parseFloat(e.insurancePaid1)||0;const ie=parseFloat(e.ins)||0;return s+p+(ip>0?ip:ie);},0);
                  return (
                    <tr style={{borderTop:'2px solid #E2E8F0',background:'#F8FAFC',fontWeight:700}}>
                      <td style={{padding:'3px 5px',fontSize:9,color:'#1B3A5C'}} colSpan={3}>Total ({data.entries.length} pts)</td>
                      <td style={{padding:'3px 5px',color:'#16A34A',textAlign:'right',fontSize:9,whiteSpace:'nowrap'}}>{fmt$(totPt)}</td>
                      <td style={{padding:'3px 5px',color:'#1D4ED8',textAlign:'right',fontSize:9,whiteSpace:'nowrap'}}>{fmt$(totInsPaid>0?totInsPaid:totIns)}</td>
                      <td style={{padding:'3px 5px',textAlign:'right',fontSize:9,whiteSpace:'nowrap',color:'#1B3A5C'}}>{fmt$(totTotal)}</td>
                      <td colSpan={2}></td>
                    </tr>
                  );
                })()}
              </tfoot>
            </table>
          </div>
        )}
      </div>
    );
  };

  const toggleExpand=key=>setExpanded(prev=>({...prev,[key]:!prev[key]}));

  return (
    <div>
      <div style={{display:'flex',gap:5,marginBottom:10,flexWrap:'wrap'}}>
        {[['thisweek','This Week'],['lastweek','Last Week'],['thismonth','This Month'],['all','All']].map(([v,l])=>(
          <button key={v} onClick={()=>setFilter(v)} style={{background:filter===v?'#1B3A5C':'#F1F5F9',color:filter===v?'white':'#64748B',border:'none',borderRadius:6,padding:'4px 10px',cursor:'pointer',fontSize:11,fontWeight:filter===v?700:400}}>{l}</button>
        ))}
      </div>
      {filter==='all'&&grouped?(
        <div>{Object.entries(grouped).map(([mKey,month])=>(
          <div key={mKey} style={{marginBottom:6}}>
            <button onClick={()=>toggleExpand(mKey)} style={{background:'#3B6B8A',color:'white',border:'none',borderRadius:7,padding:'5px 14px',cursor:'pointer',fontSize:11,fontWeight:700,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between'}}>
              <span>{month.label}</span><span>{expanded[mKey]?'▲':'▼'}</span>
            </button>
            {expanded[mKey]&&<div style={{paddingLeft:10,marginTop:3}}>
              {Object.entries(month.weeks).map(([wKey,week])=>(
                <div key={wKey} style={{marginBottom:3}}>
                  <button onClick={()=>toggleExpand(wKey)} style={{background:'#F1F5F9',color:'#1B3A5C',border:'none',borderRadius:6,padding:'3px 12px',cursor:'pointer',fontSize:10,fontWeight:600,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between',marginBottom:2}}>
                    <span>Week {getWeek(week.dates[0])} — {week.label}</span><span>{expanded[wKey]?'▲':'▼'}</span>
                  </button>
                  {expanded[wKey]&&<div style={{paddingLeft:8}}>
                    {week.dates.map(date=>(
                      <div key={date} style={{marginBottom:4}}>
                        <p style={{fontSize:9,fontWeight:700,color:'#64748B',marginBottom:2,paddingLeft:2}}>{dayOfWeek(date)}, {fmtDate(date)}</p>
                        {LOCATIONS.map(loc=><DayRow key={loc} date={date} loc={loc}/>)}
                      </div>
                    ))}
                  </div>}
                </div>
              ))}
            </div>}
          </div>
        ))}</div>
      ):(
        <div>{dates.map(date=>(
          <div key={date} style={{marginBottom:5}}>
            <p style={{fontSize:10,fontWeight:700,color:'#1B3A5C',marginBottom:2,paddingLeft:2}}>{dayOfWeek(date)}, {fmtDate(date)}</p>
            {LOCATIONS.map(loc=><DayRow key={loc} date={date} loc={loc}/>)}
          </div>
        ))}</div>
      )}
    </div>
  );
}


// ── Export modal ─────────────────────────────────────────────────────────────
function ExportModal({entries,onClose,exportFn}) {
  const [from,setFrom]=useState('');
  const [to,setTo]=useState('');
  const doExport=()=>{
    let filtered=entries;
    if(from) filtered=filtered.filter(e=>e.date>=from);
    if(to) filtered=filtered.filter(e=>e.date<=to);
    exportFn(filtered);
    onClose();
  };
  return (
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center',padding:16}}>
      <div style={{background:'white',borderRadius:12,maxWidth:400,width:'100%',padding:24,boxShadow:'0 8px 32px rgba(0,0,0,0.2)'}}>
        <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:17,color:'#1B3A5C',marginBottom:14}}>Export to Excel</h3>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,marginBottom:16}}>
          <label style={{display:'flex',flexDirection:'column',gap:4}}>
            <span style={{fontSize:11,fontWeight:700,color:'#64748B',textTransform:'uppercase'}}>From (optional)</span>
            <input type="date" value={from} onChange={e=>setFrom(e.target.value)}
              style={{padding:'7px 10px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:13,outline:'none'}}/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:4}}>
            <span style={{fontSize:11,fontWeight:700,color:'#64748B',textTransform:'uppercase'}}>To (optional)</span>
            <input type="date" value={to} onChange={e=>setTo(e.target.value)}
              style={{padding:'7px 10px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:13,outline:'none'}}/>
          </label>
        </div>
        <p style={{fontSize:11,color:'#94A3B8',marginBottom:16}}>
          {(!from&&!to)?'Exporting all entries.':
           (from&&to)?`Exporting ${from} to ${to}.`:
           from?`Exporting from ${from} onwards.`:`Exporting up to ${to}.`}
        </p>
        <div style={{display:'flex',gap:8}}>
          <button onClick={doExport} style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:600}}>Export</button>
          <button onClick={onClose} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'9px 12px',cursor:'pointer',fontSize:13}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ── Loss Report ──────────────────────────────────────────────────────────────
function LossReport({entries,doctorList}) {
  const [filterLoc,setFilterLoc]=useState('');
  const [filterDoctor,setFilterDoctor]=useState('');
  const [filterFrom,setFilterFrom]=useState('');
  const [filterTo,setFilterTo]=useState('');
  const [filterType,setFilterType]=useState(''); // '' | 'payment' | 'insurance'
  const [showPrint,setShowPrint]=useState(false);

  const lossEntries=useMemo(()=>{
    return sortEntries(entries.filter(e=>{
      const hasLoss=(parseFloat(e.paymentErrorLoss)||0)>0||(parseFloat(e.insuranceNonpaymentLoss)||0)>0;
      if(!hasLoss)return false;
      if(filterLoc&&e.location!==filterLoc)return false;
      if(filterDoctor&&e.doctorId!==filterDoctor)return false;
      if(filterFrom&&e.date<filterFrom)return false;
      if(filterTo&&e.date>filterTo)return false;
      if(filterType==='payment'&&!(parseFloat(e.paymentErrorLoss)||0))return false;
      if(filterType==='insurance'&&!(parseFloat(e.insuranceNonpaymentLoss)||0))return false;
      return true;
    }));
  },[entries,filterLoc,filterDoctor,filterFrom,filterTo,filterType]);

  const totalPaymentLoss=lossEntries.reduce((s,e)=>s+(parseFloat(e.paymentErrorLoss)||0),0);
  const totalInsLoss=lossEntries.reduce((s,e)=>s+(parseFloat(e.insuranceNonpaymentLoss)||0),0);
  const totalLoss=totalPaymentLoss+totalInsLoss;

  // Group by doctor for summary
  const byDoctor={};
  for(const e of lossEntries){
    const d=e.doctorId||'Unknown';
    if(!byDoctor[d])byDoctor[d]={payment:0,insurance:0};
    byDoctor[d].payment+=parseFloat(e.paymentErrorLoss)||0;
    byDoctor[d].insurance+=parseFloat(e.insuranceNonpaymentLoss)||0;
  }

  // Group by location
  const byLoc={};
  for(const e of lossEntries){
    const l=e.location||'?';
    if(!byLoc[l])byLoc[l]={payment:0,insurance:0};
    byLoc[l].payment+=parseFloat(e.paymentErrorLoss)||0;
    byLoc[l].insurance+=parseFloat(e.insuranceNonpaymentLoss)||0;
  }

  const inp={padding:'5px 8px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:11,fontFamily:"'DM Sans',sans-serif",outline:'none',background:'white'};

  const PrintReport=()=>(
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center',padding:16}}>
      <style>{`@media print{@page{size:letter portrait;margin:0.7in;}body *{visibility:hidden;}#loss-print,#loss-print *{visibility:visible;}#loss-print{position:absolute;left:0;top:0;width:100%;}.no-print{display:none!important;}}`}</style>
      <div id="loss-print" style={{background:'white',borderRadius:14,maxWidth:720,width:'100%',maxHeight:'90vh',overflow:'auto',padding:28,boxShadow:'0 8px 40px rgba(0,0,0,0.2)'}}>
        <div className="no-print" style={{display:'flex',gap:8,marginBottom:14,justifyContent:'flex-end'}}>
          <button onClick={()=>window.print()} style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'7px 16px',cursor:'pointer',fontSize:12,fontWeight:600}}>Print / PDF</button>
          <button onClick={()=>setShowPrint(false)} style={{background:'#F1F5F9',color:'#64748B',border:'none',borderRadius:7,padding:'7px 12px',cursor:'pointer',fontSize:12}}>Close</button>
        </div>
        <div style={{borderBottom:'2px solid #1B3A5C',paddingBottom:12,marginBottom:18}}>
          <h1 style={{fontFamily:"'DM Serif Display',serif",fontSize:20,color:'#1B3A5C'}}>The Spark Optometry — Loss Report</h1>
          <p style={{fontSize:12,color:'#64748B',marginTop:4}}>
            {filterFrom||filterTo?`${filterFrom||'start'} to ${filterTo||'today'}`:'All dates'}
            {filterLoc?` · ${LOC_FULL[filterLoc]}`:''}
            {filterDoctor?` · Dr. ${filterDoctor}`:''}
            {' · Generated '}{new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'})}
          </p>
        </div>

        {/* Summary */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:12,marginBottom:20}}>
          {[['Payment Error Losses',totalPaymentLoss,'#EF4444'],['Insurance Non-Payment',totalInsLoss,'#F59E0B'],['Total Losses',totalLoss,'#1B3A5C']].map(([l,v,c])=>(
            <div key={l} style={{background:'#F8FAFC',borderRadius:8,padding:'10px 14px',border:'1px solid #E2E8F0'}}>
              <p style={{fontSize:10,color:'#64748B',marginBottom:4}}>{l}</p>
              <p style={{fontSize:18,fontWeight:800,color:c}}>${v.toFixed(2)}</p>
            </div>
          ))}
        </div>

        {/* By doctor */}
        {Object.keys(byDoctor).length>0&&(
          <div style={{marginBottom:18}}>
            <h3 style={{fontSize:13,fontWeight:700,color:'#1B3A5C',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>By Doctor</h3>
            {Object.entries(byDoctor).sort(([,a],[,b])=>(b.payment+b.insurance)-(a.payment+a.insurance)).map(([dr,v])=>(
              <div key={dr} style={{display:'flex',justifyContent:'space-between',padding:'5px 0',borderBottom:'1px solid #F0F4F8',fontSize:12}}>
                <span>Dr. {dr}</span>
                <span style={{color:'#EF4444'}}>Payment: ${v.payment.toFixed(2)} &nbsp; Insurance: ${v.insurance.toFixed(2)} &nbsp; <strong>Total: ${(v.payment+v.insurance).toFixed(2)}</strong></span>
              </div>
            ))}
          </div>
        )}

        {/* Detail */}
        <h3 style={{fontSize:13,fontWeight:700,color:'#1B3A5C',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>Detail</h3>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
          <thead>
            <tr style={{borderBottom:'2px solid #1B3A5C'}}>
              {['Date','Location','Doctor','Patient','Payment Err $','Ins Non-Pay $','Notes/ATTN'].map(h=>(
                <th key={h} style={{padding:'5px 6px',textAlign:'left',fontSize:9,fontWeight:700,color:'#64748B',textTransform:'uppercase'}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {lossEntries.map((e,i)=>(
              <tr key={e.id} style={{borderBottom:'1px solid #F0F4F8',background:i%2===0?'white':'#FAFBFC'}}>
                <td style={{padding:'4px 6px',fontSize:11}}>{fmtDate(e.date)}</td>
                <td style={{padding:'4px 6px',fontSize:11}}>{LOC_FULL[e.location]||e.location}</td>
                <td style={{padding:'4px 6px',fontSize:11}}>{e.doctorId}</td>
                <td style={{padding:'4px 6px',fontSize:11,fontWeight:600}}>{e.patientName}</td>
                <td style={{padding:'4px 6px',fontSize:11,textAlign:'right',color:'#EF4444',fontWeight:600}}>{parseFloat(e.paymentErrorLoss)>0?`$${parseFloat(e.paymentErrorLoss).toFixed(2)}`:''}</td>
                <td style={{padding:'4px 6px',fontSize:11,textAlign:'right',color:'#F59E0B',fontWeight:600}}>{parseFloat(e.insuranceNonpaymentLoss)>0?`$${parseFloat(e.insuranceNonpaymentLoss).toFixed(2)}`:''}</td>
                <td style={{padding:'4px 6px',fontSize:10,color:'#64748B'}}>{e.attn||e.notes||''}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr style={{borderTop:'2px solid #1B3A5C',fontWeight:700}}>
              <td colSpan={4} style={{padding:'6px 6px',fontSize:11}}>Total</td>
              <td style={{padding:'6px 6px',textAlign:'right',color:'#EF4444',fontSize:12}}>${totalPaymentLoss.toFixed(2)}</td>
              <td style={{padding:'6px 6px',textAlign:'right',color:'#F59E0B',fontSize:12}}>${totalInsLoss.toFixed(2)}</td>
              <td style={{padding:'6px 6px',textAlign:'right',color:'#1B3A5C',fontSize:12}}><strong>${totalLoss.toFixed(2)}</strong></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );

  return (
    <div style={{maxWidth:1100,margin:'0 auto',padding:'0 14px'}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:14,flexWrap:'wrap',gap:10}}>
        <div>
          <h2 style={{fontFamily:"'DM Serif Display',serif",fontSize:20,color:'#1B3A5C'}}>Loss Report</h2>
          <p style={{fontSize:11,color:'#94A3B8',marginTop:2}}>{lossEntries.length} entries with losses</p>
        </div>
        <button onClick={()=>setShowPrint(true)} style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'8px 16px',cursor:'pointer',fontSize:12,fontWeight:600}}>Print / PDF Report</button>
      </div>

      {/* Filters */}
      <div style={{display:'flex',gap:8,marginBottom:14,flexWrap:'wrap',alignItems:'center',background:'white',padding:'8px 12px',borderRadius:8,border:'1px solid #E2E8F0'}}>
        <input type="date" value={filterFrom} onChange={e=>setFilterFrom(e.target.value)} style={inp} placeholder="From"/>
        <input type="date" value={filterTo} onChange={e=>setFilterTo(e.target.value)} style={inp} placeholder="To"/>
        <select value={filterLoc} onChange={e=>setFilterLoc(e.target.value)} style={inp}><option value="">All Locs</option>{LOCATIONS.map(l=><option key={l} value={l}>{l}</option>)}</select>
        <select value={filterDoctor} onChange={e=>setFilterDoctor(e.target.value)} style={inp}><option value="">All Drs</option>{[...doctorList].sort().map(d=><option key={d} value={d}>{d}</option>)}</select>
        <select value={filterType} onChange={e=>setFilterType(e.target.value)} style={inp}>
          <option value="">All loss types</option>
          <option value="payment">Payment errors only</option>
          <option value="insurance">Insurance only</option>
        </select>
        {(filterLoc||filterDoctor||filterFrom||filterTo||filterType)&&<button onClick={()=>{setFilterLoc('');setFilterDoctor('');setFilterFrom('');setFilterTo('');setFilterType('');}} style={{background:'none',border:'none',color:'#94A3B8',cursor:'pointer',fontSize:11}}>Clear</button>}
      </div>

      {/* Summary cards */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(180px,1fr))',gap:10,marginBottom:16}}>
        {[['Payment Error Losses',totalPaymentLoss,'#EF4444'],['Insurance Non-Payment',totalInsLoss,'#F59E0B'],['Total Losses',totalLoss,'#1B3A5C']].map(([l,v,c])=>(
          <div key={l} style={{background:'white',borderRadius:10,padding:'12px 16px',border:'1px solid #E2E8F0'}}>
            <p style={{fontSize:10,color:'#64748B',marginBottom:4}}>{l}</p>
            <p style={{fontSize:20,fontWeight:800,color:c}}>${v.toFixed(2)}</p>
          </div>
        ))}
      </div>

      {/* By location summary */}
      {Object.keys(byLoc).length>0&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(200px,1fr))',gap:8,marginBottom:16}}>
          {Object.entries(byLoc).map(([loc,v])=>(
            <div key={loc} style={{background:'white',borderRadius:8,padding:'10px 14px',border:'1px solid #E2E8F0'}}>
              <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:6}}>
                <span style={{fontWeight:700,fontSize:12,background:LOC_COLORS[loc]+'20',color:LOC_COLORS[loc],borderRadius:4,padding:'1px 7px'}}>{loc}</span>
              </div>
              <p style={{fontSize:11,color:'#EF4444'}}>Err: ${v.payment.toFixed(2)}</p>

              <p style={{fontSize:12,fontWeight:700,color:'#1B3A5C'}}>Total: ${(v.payment+v.insurance).toFixed(2)}</p>
            </div>
          ))}
        </div>
      )}

      {/* Detail table */}
      <div style={{overflowX:'auto'}}>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
          <thead>
            <tr style={{borderBottom:'2px solid #F0F4F8',background:'#F8FAFC'}}>
              {['Date','Loc','Dr','Patient','Err $','Ins Non-Pay $','ATTN / Notes'].map(h=>(
                <th key={h} style={{padding:'3px 4px',textAlign:'left',fontSize:9,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',borderBottom:'2px solid #E2E8F0',whiteSpace:'nowrap'}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {lossEntries.map((e,i)=>(
              <tr key={e.id} style={{borderBottom:'1px solid #F0F4F8',background:i%2===0?'white':'#FAFBFC'}}>
                <td style={{padding:'3px 4px',whiteSpace:'nowrap',fontSize:10}}>{fmtDate(e.date)}</td>
                <td style={{padding:'3px 3px'}}>{e.location&&<span style={{background:LOC_COLORS[e.location]+'25',color:LOC_COLORS[e.location],borderRadius:3,padding:'1px 4px',fontSize:10,fontWeight:700}}>{e.location}</span>}</td>
                <td style={{padding:'3px 3px',color:'#64748B',fontSize:10}}>{e.doctorId}</td>
                <td style={{padding:'3px 4px',fontWeight:600,fontSize:11,maxWidth:110,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{e.patientName}</td>
                <td style={{padding:'3px 4px',textAlign:'right',color:'#EF4444',fontWeight:700}}>{parseFloat(e.paymentErrorLoss)>0?fmt$(e.paymentErrorLoss):''}</td>
                <td style={{padding:'3px 4px',textAlign:'right',color:'#F59E0B',fontWeight:700}}>{parseFloat(e.insuranceNonpaymentLoss)>0?fmt$(e.insuranceNonpaymentLoss):''}</td>
                <td style={{padding:'3px 4px',fontSize:10,color:'#6D28D9',maxWidth:160,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{[e.attn,e.notes].filter(Boolean).join(' · ')||''}</td>
              </tr>
            ))}
          </tbody>
          {lossEntries.length>0&&(
            <tfoot>
              <tr style={{borderTop:'2px solid #1B3A5C',background:'#F8FAFC',fontWeight:700}}>
                <td style={{padding:'6px 8px',fontSize:12,fontWeight:700}}>Total</td>
                <td style={{padding:'6px 8px',textAlign:'right',color:'#EF4444',fontSize:13}}>{fmt$(totalPaymentLoss)}</td>
                <td colSpan={3}></td>
                <td colSpan={2} style={{padding:'6px 8px',textAlign:'right',color:'#F59E0B',fontSize:13}}>{fmt$(totalInsLoss)}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
      {showPrint&&<PrintReport/>}
    </div>
  );
}


// ── Main ────────────────────────────────────────────────────────────────────
export default function BillingSheet({user,profile,isMaster}) {
  const [entries,setEntries]=useState([]);
  const [doctorLogs,setDoctorLogs]=useState([]);
  const [loading,setLoading]=useState(true);
  const [feeSettings,setFeeSettings]=useState(DEFAULT_FEES);
  const [otherItems,setOtherItems]=useState(DEFAULT_OTHER_ITEMS);
  const [doctorList,setDoctorList]=useState(DEFAULT_DOCTORS);
  const [scheduleTemplate,setScheduleTemplate]=useState({assignments:{},updatedAt:''});
  const [mode,setMode]=useState('review');
  // Column width customization — persisted to localStorage
  const [colWidths, setColWidths] = React.useState(()=>{
    try { return JSON.parse(localStorage.getItem('sparkColWidths')||'{}'); } catch { return {}; }
  });
  const [showColSettings, setShowColSettings] = React.useState(false);
  const setColWidth = React.useCallback((tableKey, colKey, val) => {
    setColWidths(prev => {
      const updated = {...prev, [tableKey]: {...(prev[tableKey]||{}), [colKey]: parseInt(val)||0}};
      localStorage.setItem('sparkColWidths', JSON.stringify(updated));
      return updated;
    });
  }, []);
  const cw = (tableKey, colKey, def) => colWidths?.[tableKey]?.[colKey] || def;


  // EOB navigate handler — allows EOBScreen to navigate to Review with patient search
  React.useEffect(()=>{ window._eobNavigate=(name)=>{setSearchName(name);setSearchInput(name);setMode('review');}; return()=>{delete window._eobNavigate;}; },[]);
  const [entryDate,setEntryDate]=useState(today());
  const [entryLoc,setEntryLoc]=useState('');
  const [entryDoctor,setEntryDoctor]=useState('');
  const [entryInitCount,setEntryInitCount]=useState(0);
  const [entryKey,setEntryKey]=useState(0);
  const [todayLoc,setTodayLoc]=useState('');
  const [todayDoctor,setTodayDoctor]=useState('');
  const [todayInitCount,setTodayInitCount]=useState(0);
  const [showPaste,setShowPaste]=useState(false);
  const [pasteText,setPasteText]=useState('');
  const [pasteNames,setPasteNames]=useState([]);
  const [pasteSaving,setPasteSaving]=useState(false);
  const [pasteCount,setPasteCount]=useState(0);
  const [todayKey,setTodayKey]=useState(0);
  const [reviewFilter,setReviewFilter]=useState('twomonths');
  const [customMonth,setCustomMonth]=useState('');
  const [filterDate,setFilterDate]=useState('');
  const [filterLoc,setFilterLoc]=useState('');
  const [filterDoctor,setFilterDoctor]=useState('');
  const [filterStatus,setFilterStatus]=useState('');
  const [filterCash,setFilterCash]=useState('');
  const [filterIns,setFilterIns]=useState('');
  const [searchName,setSearchName]=useState('');
  const [searchInput,setSearchInput]=useState('');
  const [expandedRows,setExpandedRows]=useState({});
  const [editEntry,setEditEntry]=useState(null);
  const [resolveEntry,setResolveEntry]=useState(null);
  const [showCalendar,setShowCalendar]=useState(false);
  const [showEOB,setShowEOB]=useState(false);
  const [showExport,setShowExport]=useState(false);
  const [showSvcs,setShowSvcs]=useState(false);
  const [showNotes,setShowNotes]=useState(false);
  const [queue,setQueue]=useState([]);

  // Patient count inputs — separate state for today vs entry, isolated from table



  // Parse TAB schedule text to extract patient names in order
  const parseScheduleText = (text) => {
    const examPattern = /\b(EP|NP|OV|RGP|Myopia|CL|EEX)\b/i;
    const ignorePattern = /^[\s(]|No Doctor|Lunch|\d{1,2}:\d{2}\s*(AM|PM)/i;
    const names = [];
    const seen = new Set();
    for (const line of text.split('\n')) {
      const cells = line.split('\t');
      for (const cell of cells) {
        const trimmed = cell.trim();
        if (!trimmed || ignorePattern.test(trimmed)) continue;
        const match = trimmed.match(examPattern);
        if (match) {
          const name = trimmed.slice(0, match.index).trim().replace(/[^a-zA-Z\s\-']/g,'').trim();
          if (name.length > 2 && name.includes(' ') && !seen.has(name.toLowerCase())) {
            seen.add(name.toLowerCase());
            // Convert "First Last" to "Last, First" format, preserve original casing
            const parts = name.trim().split(/\s+/);
            const formatted = parts.length >= 2
              ? parts[parts.length-1] + ', ' + parts.slice(0,-1).join(' ')
              : name;
            names.push(formatted);
          }
        }
      }
    }
    return names;
  };

  useEffect(()=>{
    async function load(){
      try{
        const [snap,logSnap,fees,items,docs,sched]=await Promise.all([
          getDocs(query(collection(db,'billingEntries'),orderBy('date','desc'))), // TODO: paginate for performance
          getDocs(collection(db,'doctorLogs')),
          loadFeeSettings(),loadOtherItems(),loadDoctorList(),loadScheduleTemplate(),
        ]);
        setEntries(snap.docs.map(d=>({id:d.id,...d.data()})));
        setDoctorLogs(logSnap.docs.map(d=>({id:d.id,...d.data()})));
        setFeeSettings(fees);setOtherItems(items);setDoctorList(docs);setScheduleTemplate(sched);
      }catch(e){console.error(e);}
      setLoading(false);
    }
    load();
  },[]);

  const handleSave=useCallback(u=>setEntries(prev=>{const i=prev.findIndex(e=>e.id===u.id);if(i>=0){const n=[...prev];n[i]=u;return n;}return[u,...prev];}),[]);
  const handleDelete=useCallback(id=>setEntries(prev=>prev.filter(e=>e.id!==id)),[]);
  const handleUpdate=useCallback(async(id,field,value)=>{
    const entry=entries.find(e=>e.id===id);if(!entry)return;
    const updated=field==='_multi'?{...entry,...value,updatedAt:new Date().toISOString()}:{...entry,[field]:value,updatedAt:new Date().toISOString()};
    await setDoc(doc(db,'billingEntries',id),updated);
    setEntries(prev=>prev.map(e=>e.id===id?updated:e));
  },[entries]);

  const addToQueue=(date,loc,doctor)=>{const key=date+'|'+loc;if(queue.find(q=>q.key===key))return;setQueue(prev=>[...prev,{key,date,loc,doctor}]);};
  const swapFromQueue=(idx)=>{
    const item=queue[idx];const newQ=[...queue];newQ.splice(idx,1);
    if(entryDate&&entryLoc){const ck=entryDate+'|'+entryLoc;if(!newQ.find(q=>q.key===ck))newQ.push({key:ck,date:entryDate,loc:entryLoc,doctor:entryDoctor});}
    setQueue(newQ);setEntryDate(item.date);setEntryLoc(item.loc);setEntryDoctor(item.doctor||'');
    setEntryInitCount(0);setEntryKey(k=>k+1);setMode('entry');
  };
  const removeFromQueue=idx=>setQueue(prev=>prev.filter((_,i)=>i!==idx));
  const queueCurrent=()=>{if(entryDate&&entryLoc)addToQueue(entryDate,entryLoc,entryDoctor);};

  const flaggedCount=entries.filter(e=>computeStatus(e)==='flagged').length;

  const filtered=useMemo(()=>{
    const now=new Date();
    const getRange=(off=0)=>{const d=new Date(now);d.setDate(d.getDate()-d.getDay()+off*7);const s=d.toISOString().slice(0,10);const e=new Date(d.setDate(d.getDate()+6)).toISOString().slice(0,10);return[s,e];};
    return sortEntries(entries.filter(e=>{
      if(!e.date)return false;
      // Search bypasses all date filters — searches all entries
      if(searchName){
        const n=searchName.toLowerCase();
        if(!e.patientName?.toLowerCase().includes(n)&&!e.nickname?.toLowerCase().includes(n))return false;
        if(filterLoc&&e.location!==filterLoc)return false;
        if(filterDoctor&&e.doctorId!==filterDoctor)return false;
        return true;
      }
      if(filterDate&&e.date!==filterDate)return false;
      else if(!filterDate){
        if(reviewFilter==='thisweek'){const[s,en]=getRange(0);if(e.date<s||e.date>en)return false;}
        else if(reviewFilter==='lastweek'){const[s,en]=getRange(-1);if(e.date<s||e.date>en)return false;}
        else if(reviewFilter==='thismonth'){const m=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');if(!e.date?.startsWith(m))return false;}
        else if(reviewFilter==='twomonths'){
          const m1=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
          const lastMo=new Date(now.getFullYear(),now.getMonth()-1,1);
          const m2=lastMo.getFullYear()+'-'+String(lastMo.getMonth()+1).padStart(2,'0');
          if(!e.date?.startsWith(m1)&&!e.date?.startsWith(m2))return false;
        }
        else if(reviewFilter==='custom'&&customMonth){if(!e.date?.startsWith(customMonth))return false;}
      }
      if(filterLoc&&e.location!==filterLoc)return false;
      if(filterDoctor&&e.doctorId!==filterDoctor)return false;
      if(filterStatus&&computeStatus(e)!==filterStatus)return false;
      if(filterCash==='expected'&&e.cashStatus!=='expected')return false;
      if(filterIns&&(e.insPaidState||'pending')!==filterIns)return false;
      return true;
    }));
  },[entries,reviewFilter,customMonth,filterDate,filterLoc,filterDoctor,filterStatus,filterCash,filterIns,searchName]);

  const groupedByWeek=useMemo(()=>{
    if(reviewFilter!=='custom'||!customMonth||filterDate)return null;
    const weeks={};
    for(const e of filtered){const wi=getFiscalWeekInfo(e.date);const wNum=wi.week;const wKey=`W${String(wNum).padStart(2,'0')}`;if(!weeks[wKey])weeks[wKey]={label:`W${wNum}: ${getWeekLabel(e.date)}`,entries:[]};weeks[wKey].entries.push(e);}
    return weeks;
  },[filtered,reviewFilter,customMonth,filterDate]);

  const exportExcel=(filteredEntries)=>{
    const headers=['Location','Day','Month','Week','Date','Status','Doctor','Gross','Patient Name','Routine','CL','Optos','DFE','OV','OCT','Other','Myopia','Paid','Ins','Total','Payor','Payor 2','Cash','Insurance Paid','Payment Error Loss','Insurance Nonpayment Loss','ATTN'];
    const rows=(filteredEntries||entries).map(e=>{
      const d=new Date((e.date||today())+'T00:00:00');
      const gross=['exam','cl','optos','oct','dfe','ov'].reduce((s,k)=>s+(parseFloat(e[k])||0),0)+(parseFloat(e.otherAmt)||0)+(parseFloat(e.otherAmt2)||0);
      const insPaid=(parseFloat(e.insurancePaid1)||0)+(parseFloat(e.insurancePaid2)||0);
      const total=gross+(parseFloat(e.myopia)||0);
      return[e.location,d.toLocaleDateString('en-US',{weekday:'long'}),d.getMonth()+1,getWeek(e.date),e.date,computeStatus(e),e.doctorId,gross,e.patientName,
        e.exam||e.routine||'',e.cl||'',e.optos||'',e.dfe||'',e.ov||'',e.oct||'',e.otherType?`${e.otherType}:${e.otherAmt||0}${e.otherType2?` / ${e.otherType2}:${e.otherAmt2||0}`:''}` :'',e.myopia||'',
        e.ptPaid||'',e.ins||'',total,e.payor1||'',e.payor2||'',e.cashStatus||'',insPaid>0?insPaid:'','','',e.attn||''];
    });
    const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'2026');
    XLSX.writeFile(wb,`spark_billing_${today()}.xlsx`);
  };

  const inp={padding:'5px 8px',border:'1.5px solid #E2E8F0',borderRadius:7,fontSize:11,fontFamily:"'DM Sans',sans-serif",outline:'none',background:'white'};

  const EntryTopBar=({isToday})=>{
    const loc=isToday?todayLoc:entryLoc;
    const setLoc=isToday?(v=>setTodayLoc(v)):(v=>setEntryLoc(v));
    const doctor=isToday?todayDoctor:entryDoctor;
    const setDoctor=isToday?(v=>setTodayDoctor(v)):(v=>setEntryDoctor(v));
    const openRows=(n)=>{
      if(n<=0)return;
      if(isToday) setTodayInitCount(n);
      else setEntryInitCount(n);
    };
    return (
      <div style={{display:'flex',gap:7,marginBottom:10,padding:'7px 12px',background:isToday?'#EFF6FF':'#F0FDF4',borderRadius:8,border:`1px solid ${isToday?'#BFDBFE':'#BBF7D0'}`,flexWrap:'wrap',alignItems:'center'}}>
        <span style={{fontSize:11,fontWeight:700,color:isToday?'#1D4ED8':'#065F46',flexShrink:0}}>
          {isToday?`${DAY_FULL[new Date().getDay()]}, ${fmtDate(today())}`:'Entry Mode'}
        </span>
        {!isToday&&<input type="date" value={entryDate} onChange={e=>setEntryDate(e.target.value)} style={inp}/>}
        <select value={loc} onChange={e=>setLoc(e.target.value)} style={inp}>
          <option value="">Select location...</option>
          {LOCATIONS.map(l=><option key={l} value={l}>{l} — {LOC_FULL[l]}</option>)}
        </select>
        <select value={doctor} onChange={e=>setDoctor(e.target.value)} style={inp}>
          <option value="">Select doctor...</option>
          {[...doctorList].sort().map(d=><option key={d} value={d}>{d}</option>)}
        </select>
        {/* Patient count — uncontrolled input to avoid stale closure / bleed issues */}
        <div style={{display:'flex',alignItems:'center',gap:4,background:'white',borderRadius:6,padding:'3px 8px',border:'1px solid #E2E8F0',flexShrink:0}}>
          <span style={{fontSize:10,color:'#64748B',userSelect:'none'}}>Pts:</span>
          <input
            type="text"
            defaultValue=""
            id={isToday?'ptcount-today':'ptcount-entry'}
            onKeyDown={e=>{
              e.stopPropagation();
              if(e.key==='Enter'){
                e.preventDefault();
                const n=parseInt(e.target.value)||0;
                if(n>0){
                  openRows(n);
                  e.target.value='';
                }
              }
            }}
            style={{width:36,padding:'2px 4px',border:'none',fontSize:11,textAlign:'center',outline:'none',fontFamily:"'DM Sans',sans-serif",appearance:'none',MozAppearance:'textfield',WebkitAppearance:'none'}}
            placeholder="#"
          />
          <span style={{fontSize:9,color:'#94A3B8',userSelect:'none'}}>↵</span>
        </div>
        <button onClick={()=>{setShowPaste(v=>!v);setPasteText('');}}
          style={{background:'#EFF6FF',border:'1px solid #BFDBFE',borderRadius:6,padding:'3px 8px',cursor:'pointer',fontSize:10,color:'#1D4ED8',fontWeight:600,flexShrink:0}}>
          📋 Paste Schedule
        </button>
      </div>
    );
  };

  const ReviewTable=({items,pageSize=75})=>{
    const [page,setPage]=React.useState(1);
    // Reset page when items change
    React.useEffect(()=>setPage(1),[items]);
    const visible=items.slice(0,page*pageSize);
    const hasMore=visible.length<items.length;
    return (<div style={{overflowX:'auto'}}>
      <style>{focusStyle}</style>
      <table style={{width:'100%',borderCollapse:'collapse',fontSize:10}}>
        <thead>
          <tr style={{borderBottom:'2px solid #F0F4F8',background:'#F8FAFC'}}>
            {/* align: expand | Day | Date | Loc | Dr | Patient | (svcs) | PtPaid | InsExp | Total | % | InsPaid | P1 | P2 | Cash | (attn) (notes) | Status | actions */}
            <th style={{width:12}}></th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',whiteSpace:'nowrap',textAlign:'left'}}>Day</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',whiteSpace:'nowrap',textAlign:'left'}}>Date</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Loc</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Dr</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Patient</th>
            {showSvcs&&<th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Services</th>}
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'right'}}>Pt Paid</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'right'}}>Ins Exp</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'right'}}>Total</th>
            <th style={{padding:'4px 3px',width:16}}></th>{/* discount badge */}
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Cash</th>
            <th style={{padding:'4px 1px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'right'}}>P1$</th>
            <th style={{padding:'4px 2px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>P1</th>
            <th style={{padding:'4px 1px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'right'}}>P2$</th>
            <th style={{padding:'4px 1px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>P2</th>
            <th style={{padding:'4px 1px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Ins Paid</th>
            <th style={{padding:'4px 1px',fontSize:7,fontWeight:700,color:'#EF4444',textTransform:'uppercase',textAlign:'right'}}>Err $</th>
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#EF4444',textTransform:'uppercase',textAlign:'right'}}>Ins $</th>
            {showNotes&&<th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#8B5CF6',textTransform:'uppercase',textAlign:'left'}}>ATTN</th>}
            {showNotes&&<th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Notes</th>}
            <th style={{padding:'4px 3px',fontSize:7,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',textAlign:'left'}}>Status</th>
            <th style={{width:60}}></th>
          </tr>
        </thead>
        <tbody>
          {visible.map((e,i)=>{
            const ptPaid=parseFloat(e.ptPaid)||0;const insPaid=parseFloat(e.insurancePaid1)||0;const insExp=parseFloat(e.ins)||0;
            const total=ptPaid+(insPaid>0?insPaid:insExp);const isExp=expandedRows[e.id];const hasDiscount=parseFloat(e.discountAmount)>0;
            return (
              <React.Fragment key={e.id}>
                <tr style={{background:computeStatus(e)==='flagged'?'#FFF7ED':i%2===0?'white':'#FAFBFC',borderBottom:isExp?'none':'1px solid #F0F4F8',cursor:'pointer'}}
                  onClick={()=>setExpandedRows(prev=>({...prev,[e.id]:!prev[e.id]}))}>
                  <td style={{padding:'3px 3px',textAlign:'center',fontSize:8,color:'#94A3B8'}}>{isExp?'▼':'▶'}</td>
                  <td style={{padding:'3px 3px',color:'#94A3B8',fontSize:8,whiteSpace:'nowrap'}}>{dayOfWeek(e.date,true)}</td>
                  <td style={{padding:'3px 3px 3px 1px',whiteSpace:'nowrap',color:'#64748B',fontSize:9}}>{fmtDate(e.date)}</td>
                  <td style={{padding:'3px 1px'}}>{e.location&&<span style={{background:LOC_COLORS[e.location]+'25',color:LOC_COLORS[e.location],borderRadius:3,padding:'1px 3px',fontSize:8,fontWeight:700}}>{e.location}</span>}</td>
                  <td style={{padding:'3px 3px',color:'#64748B',fontSize:9}}>{e.doctorId}</td>
                  <td style={{padding:'2px 3px',fontWeight:600,width:cw('review','patient',83),maxWidth:cw('review','patient',83),overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:10}} onClick={ev=>{ev.stopPropagation();setEditEntry(e);}}>
                    {e.patientName}{e.nickname&&<span style={{fontSize:8,color:'#94A3B8',marginLeft:2}}>({e.nickname})</span>}
                  </td>
                  {showSvcs&&<td style={{padding:'3px 3px',maxWidth:100}}>
                    <div style={{display:'flex',gap:2,flexWrap:'wrap'}}>
                      {['exam','cl','optos','oct','dfe','ov','myopia'].filter(k=>parseFloat(e[k]||e.routine)>0).map(k=>(
                        <span key={k} style={{background:k==='myopia'?'#F5F3FF':'#F1F5F9',color:k==='myopia'?'#8B5CF6':'#475569',borderRadius:3,padding:'1px 3px',fontSize:8}}>{k==='exam'?'Ex':k.toUpperCase()}</span>
                      ))}
                    </div>
                  </td>}
                  <td style={{padding:'3px 3px',textAlign:'right'}} onClick={ev=>ev.stopPropagation()}><InlineCell value={e.ptPaid||'0'} onUpdate={handleUpdate} field="ptPaid" width={40}/></td>
                  <td style={{padding:'3px 3px',textAlign:'right'}} onClick={ev=>ev.stopPropagation()}><InlineCell value={e.ins} onUpdate={handleUpdate} field="ins" width={40}/></td>
                  <td style={{padding:'3px 3px',textAlign:'right',fontWeight:700,color:'#1B3A5C',fontSize:10}}>{total===0?'$0.00':fmt$(total)||'—'}</td>
                  <td style={{padding:'3px 1px',textAlign:'center'}}>
                    {hasDiscount&&<span title={`${e.discountType||'Discount'}: $${e.discountAmount}`} style={{cursor:'help',fontSize:9,background:'#FEF3C7',color:'#92400E',borderRadius:3,padding:'0 3px'}}>%</span>}
                  </td>
                  <td style={{padding:'3px 3px'}} onClick={ev=>ev.stopPropagation()}><CashToggle value={e.cashStatus||''} onUpdate={v=>handleUpdate(e.id,'cashStatus',v)}/></td>
                  <td style={{padding:'3px 1px'}} onClick={ev=>ev.stopPropagation()}><InlineCell value={e.ins1Amt} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="ins1Amt" width={28}/></td>
                  <td style={{padding:'3px 2px 3px 2px'}} onClick={ev=>ev.stopPropagation()}><InlinePayorSelect value={e.payor1} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="payor1"/></td>
                  <td style={{padding:'3px 1px'}} onClick={ev=>ev.stopPropagation()}><InlineCell value={e.ins2Amt} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="ins2Amt" width={28}/></td>
                  <td style={{padding:'3px 0px 3px 0px'}} onClick={ev=>ev.stopPropagation()}><InlinePayorSelect value={e.payor2} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="payor2"/></td>
                  <td style={{padding:'3px 1px 3px 0px'}} onClick={ev=>ev.stopPropagation()}><InsPaidCell entry={e} onUpdate={handleUpdate}/></td>
                  <td style={{padding:'3px 1px 3px 0px',textAlign:'right'}} onClick={ev=>ev.stopPropagation()}>
                    <InlineCell value={e.paymentErrorLoss} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="paymentErrorLoss" width={34}/>
                  </td>
                  <td style={{padding:'3px 2px',textAlign:'right'}} onClick={ev=>ev.stopPropagation()}>
                    <InlineCell value={e.insuranceNonpaymentLoss} onUpdate={(field,val)=>handleUpdate(e.id,field,val)} field="insuranceNonpaymentLoss" width={34}/>
                  </td>
                  {showNotes&&<td style={{padding:'3px 3px',maxWidth:80,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:9,color:'#6D28D9'}} title={e.attn}>{e.attn||''}</td>}
                  {showNotes&&<td style={{padding:'3px 3px',maxWidth:80,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:9,color:'#64748B',fontStyle:'italic'}} title={e.notes}>{e.notes||''}</td>}
                  <td style={{padding:'3px 3px'}}><StatusCell entry={e}/></td>
                  <td style={{padding:'3px 3px'}} onClick={ev=>ev.stopPropagation()}>
                    <div style={{display:'flex',gap:2}}>
                      <button onClick={()=>setEditEntry(e)} style={{background:'#EFF6FF',border:'none',borderRadius:3,padding:'2px 5px',cursor:'pointer',color:'#1D4ED8',fontSize:8}}>Edit</button>
                      <button onClick={async()=>{if(!window.confirm('Delete?'))return;await deleteDoc(doc(db,'billingEntries',e.id));handleDelete(e.id);}} style={{background:'#FEE2E2',border:'none',borderRadius:3,padding:'2px 4px',cursor:'pointer',color:'#DC2626',fontSize:8}}>✕</button>
                    </div>
                  </td>
                </tr>
                {isExp&&(
                  <tr style={{borderBottom:'1px solid #F0F4F8',background:'#F8FAFC'}}>
                    <td colSpan={25} style={{padding:'5px 10px 6px 22px'}}>
                      <div style={{display:'flex',gap:4,flexWrap:'wrap',alignItems:'center'}}>
                        {['exam','cl','optos','oct','dfe','ov','myopia'].filter(k=>parseFloat(e[k]||e.routine)>0).map(k=>(
                          <span key={k} style={{background:k==='myopia'?'#F5F3FF':'#F1F5F9',color:k==='myopia'?'#8B5CF6':'#475569',borderRadius:4,padding:'2px 6px',fontSize:10}}>
                            {k==='exam'?'Exam':k.charAt(0).toUpperCase()+k.slice(1)}: {fmt$(e[k]||e.routine)}
                          </span>
                        ))}
                        {e.otherType&&<span style={{background:'#EFF6FF',color:'#1D4ED8',borderRadius:4,padding:'2px 6px',fontSize:10}}>{e.otherType}: {fmt$(e.otherAmt)}</span>}
                        {e.otherType2&&<span style={{background:'#EFF6FF',color:'#1D4ED8',borderRadius:4,padding:'2px 6px',fontSize:10}}>{e.otherType2}: {fmt$(e.otherAmt2)}</span>}
                        {hasDiscount&&<span style={{background:'#FEF3C7',color:'#92400E',borderRadius:4,padding:'2px 6px',fontSize:10}}>Discount: {fmt$(e.discountAmount)}{e.discountType?` (${e.discountType})`:''}</span>}
                        {e.claimNumber&&<span style={{background:'#FFF7ED',color:'#C2410C',borderRadius:4,padding:'2px 6px',fontSize:10}}>Claim: {e.claimNumber}</span>}
                        {e.notes&&<span style={{color:'#64748B',fontSize:10,fontStyle:'italic'}}>Note: {e.notes}</span>}
                        {e.attn&&(
                          <span style={{color:'#6D28D9',fontSize:10}}>
                            ATTN: {e.attn}
                            {e.attnResolved
                              ?<span style={{color:'#10B981',marginLeft:4}}>✓ Resolved{e.attnComment?`: ${e.attnComment}`:''}</span>
                              :<button onClick={ev=>{ev.stopPropagation();setResolveEntry(e);}} style={{marginLeft:6,background:'#F5F3FF',border:'none',borderRadius:4,padding:'1px 7px',cursor:'pointer',fontSize:9,color:'#6D28D9',fontWeight:700}}>Resolve</button>}
                          </span>
                        )}
                      </div>
                    </td>
                  </tr>
                )}
              </React.Fragment>
            );
          })}
          {items.length===0&&<tr><td colSpan={20} style={{padding:20,textAlign:'center',color:'#94A3B8',fontSize:11}}>No entries.</td></tr>}
        </tbody>
      </table>
      {hasMore&&(
        <div style={{textAlign:'center',padding:'10px 0',borderTop:'1px solid #F0F4F8'}}>
          <button onClick={()=>setPage(p=>p+1)}
            style={{background:'#F1F5F9',border:'1px solid #E2E8F0',borderRadius:6,padding:'6px 20px',cursor:'pointer',fontSize:11,color:'#64748B',fontWeight:600}}>
            Show all {items.length - visible.length} remaining
          </button>
        </div>
      )}
    </div>);
  };


  // ── Per-location daily totals for week views ────────────────────────────────

// ── All view drill-down (Month → Week → Days) ────────────────────────────────
  const AllDrillDown = ({items, expandedRows, setExpandedRows, ReviewTable}) => {
    // Group by month then week
    const byMonth = useMemo(()=>{
      const months = {};
      for (const e of items) {
        if (!e.date) continue;
        const dt = new Date(e.date+'T12:00:00');
        const wi = getFiscalWeekInfo(e.date);
        const mKey = getFiscalMonthKey(e.date);
        const mDt = new Date(mKey+'-15T00:00:00');
        const mLabel = `${MONTH_NAMES[mDt.getMonth()]} ${mDt.getFullYear()}`;
        const wNum = wi.week;
        const wKey = `${mKey}-W${String(wNum).padStart(2,'0')}`;
        if (!months[mKey]) months[mKey] = {label:mLabel, weeks:{}, entries:[]};
        months[mKey].entries.push(e);
        if (!months[mKey].weeks[wKey]) months[mKey].weeks[wKey] = {label:`W${wNum}: ${getWeekLabel(e.date)}`, entries:[]};
        months[mKey].weeks[wKey].entries.push(e);
      }
      return months;
    }, [items]);

    const tot = (arr) => {
      const pt = arr.reduce((s,e)=>s+(parseFloat(e.ptPaid)||0),0);
      const ins = arr.reduce((s,e)=>s+(parseFloat(e.insurancePaid1)||0)||(parseFloat(e.ins)||0),0);
      const total = arr.reduce((s,e)=>{const p=parseFloat(e.ptPaid)||0;const ip=parseFloat(e.insurancePaid1)||0;const ie=parseFloat(e.ins)||0;return s+p+(ip>0?ip:ie);},0);
      return {pt,ins,total};
    };

    return (
      <div>
        {Object.entries(byMonth).map(([mKey,month])=>{
          const mOpen = expandedRows[mKey] === true;
          const mt = tot(month.entries);
          return (
            <div key={mKey} style={{marginBottom:6}}>
              <button onClick={()=>setExpandedRows(prev=>({...prev,[mKey]:!mOpen}))}
                style={{background:'#3B6B8A',color:'white',border:'none',borderRadius:mOpen?'7px 7px 0 0':7,padding:'6px 14px',cursor:'pointer',fontSize:11,fontWeight:700,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span>{month.label} <span style={{fontWeight:400,opacity:0.7,fontSize:10}}>({month.entries.length} pts)</span></span>
                <span style={{display:'flex',gap:8,alignItems:'center',fontSize:9,fontWeight:600}}>
                  {LOCATIONS.filter(l=>month.entries.some(e=>e.location===l)).map(l=>{
                    const lt=tot(month.entries.filter(e=>e.location===l));
                    return <span key={l} style={{color:LOC_COLORS[l]||'#6EE7B7',opacity:0.9}}>{l}: {fmt$(lt.total)}</span>;
                  })}
                  <span style={{fontWeight:800,fontSize:12,color:'white'}}>{fmt$(mt.total)}</span>
                  <span style={{opacity:0.6,fontWeight:400}}>{mOpen?'▼':'▶'}</span>
                </span>
              </button>
              {mOpen&&(
                <div style={{border:'1px solid #E2E8F0',borderTop:'none',borderRadius:'0 0 7px 7px',marginBottom:2}}>
                  {Object.entries(month.weeks).map(([wKey,week])=>{
                    const wOpen = expandedRows[wKey] === true;
                    const wt = tot(week.entries);
                    return (
                      <div key={wKey}>
                        <button onClick={()=>setExpandedRows(prev=>({...prev,[wKey]:!wOpen}))}
                          style={{background:'#F1F5F9',color:'#1B3A5C',border:'none',borderBottom:'1px solid #E2E8F0',padding:'5px 14px',cursor:'pointer',fontSize:10,fontWeight:600,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                          <span>{week.label} <span style={{fontWeight:400,color:'#94A3B8',fontSize:9}}>({week.entries.length})</span></span>
                          <span style={{display:'flex',gap:8,alignItems:'center',fontSize:9,fontWeight:600}}>
                            {LOCATIONS.filter(l=>week.entries.some(e=>e.location===l)).map(l=>{
                              const lt=tot(week.entries.filter(e=>e.location===l));
                              return <span key={l} style={{color:LOC_COLORS[l]||'#1B3A5C'}}>{l}: {fmt$(lt.total)}</span>;
                            })}
                            <span style={{color:'#1B3A5C',fontWeight:800,fontSize:11}}>{fmt$(wt.total)}</span>
                            <span style={{color:'#94A3B8',fontWeight:400}}>{wOpen?'▼':'▶'}</span>
                          </span>
                        </button>
                        {wOpen&&<ReviewTable items={week.entries} pageSize={week.entries.length}/>}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };

  const WeekDayView=({items})=>{
    // Group by date
    const byDate = useMemo(()=>{
      const groups = {};
      for (const e of items) {
        if (!groups[e.date]) groups[e.date] = [];
        groups[e.date].push(e);
      }
      return groups;
    }, [items]);

    const entryTotal = (e) => {
      const p = parseFloat(e.ptPaid)||0;
      const ip = parseFloat(e.insurancePaid1)||0;
      const ie = parseFloat(e.ins)||0;
      return p + (ip>0?ip:ie);
    };

    const locTotals = (entries, loc) => {
      const locs = loc ? entries.filter(e=>e.location===loc) : entries;
      return {
        pt:    locs.reduce((s,e)=>s+(parseFloat(e.ptPaid)||0),0),
        ins:   locs.reduce((s,e)=>s+(parseFloat(e.insurancePaid1)||0)||(parseFloat(e.ins)||0),0),
        total: locs.reduce((s,e)=>s+entryTotal(e),0),
        count: locs.length,
      };
    };

    const todayStr = today();
    const todayWk = getFiscalWeekInfo(todayStr);
    const dates = Object.keys(byDate).sort((a,b)=>b.localeCompare(a));

    return (
      <div>
        {dates.map(date=>{
          const dayEntries = byDate[date];
          const dayKey = `day-${date}`;
          const isToday = date === today();
        const entryWk = getFiscalWeekInfo(date);
        const isCurrentWeek = entryWk.week === todayWk.week && entryWk.year === todayWk.year;
        const isOpen = expandedRows[dayKey] !== undefined ? expandedRows[dayKey] : isCurrentWeek;
          const dayTot = locTotals(dayEntries);

          // Per-location breakdown
          const locData = LOCATIONS.map(loc=>{
            const t = locTotals(dayEntries, loc);
            return t.count > 0 ? {loc, ...t} : null;
          }).filter(Boolean);

          return (
            <div key={date} style={{marginBottom:8}}>
              {/* Day header */}
              <button onClick={()=>setExpandedRows(prev=>({...prev,[dayKey]:!isOpen}))}
                style={{background:'#3B6B8A',color:'white',border:'none',borderRadius:isOpen?'7px 7px 0 0':7,padding:'6px 14px',cursor:'pointer',fontSize:11,fontWeight:700,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span>{dayOfWeek(date)} — {fmtDate(date)} <span style={{fontWeight:400,opacity:0.7,fontSize:10}}>({dayEntries.length} pts)</span></span>
                <span style={{display:'flex',gap:8,alignItems:'center',fontSize:10,fontWeight:600}}>
                  <span style={{fontWeight:800,fontSize:12,color:'#93C5FD'}}>{fmt$(dayTot.total)}</span>
                  <span style={{opacity:0.6,fontWeight:400}}>{isOpen?'▼':'▶'}</span>
                </span>
              </button>

              {isOpen&&(
                <div style={{border:'1px solid #E2E8F0',borderTop:'none',borderRadius:'0 0 7px 7px',overflow:'hidden'}}>
                  {/* Per-location summary strip — compact, left-aligned */}
                  <div style={{background:'#F8FAFC',padding:'4px 12px',borderBottom:'1px solid #E2E8F0',display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
                    {locData.map(({loc,total,count})=>(
                      <div key={loc} style={{display:'flex',alignItems:'center',gap:4}}>
                        <span style={{fontSize:9,fontWeight:700,background:LOC_COLORS[loc]+'25',color:LOC_COLORS[loc],borderRadius:3,padding:'1px 6px'}}>{loc}</span>
                        <span style={{fontSize:8,color:'#94A3B8'}}>{count}p</span>
                        <span style={{fontSize:9,fontWeight:700,color:'#1B3A5C'}}>{fmt$(total)}</span>
                      </div>
                    ))}
                  </div>
                  <ReviewTable items={dayEntries}/>
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }


  if(loading) return <div style={{padding:40,textAlign:'center',color:'#94A3B8'}}>Loading...</div>;

  return (
    <div style={{maxWidth:1300,margin:'0 auto',padding:'10px 12px'}}>
      <style>{focusStyle}</style>
      <div style={{display:'flex',gap:6,marginBottom:10,flexWrap:'wrap',justifyContent:'space-between',alignItems:'center'}}>
        <div style={{display:'flex',gap:6}}>
          {flaggedCount>0&&(
            <button onClick={()=>{setMode('review');setFilterStatus(filterStatus==='flagged'?'':'flagged');}}
              style={{background:'#FEE2E2',border:'1px solid #FCA5A5',borderRadius:6,padding:'4px 10px',fontSize:11,color:'#DC2626',fontWeight:600,cursor:'pointer'}}>
              ⚠ {flaggedCount} flagged
            </button>
          )}
        </div>
        <div style={{display:'flex',gap:6}}>
          <button onClick={()=>setShowCalendar(true)} style={{background:'#F1F5F9',color:'#1B3A5C',border:'1px solid #E2E8F0',borderRadius:6,padding:'5px 10px',cursor:'pointer',fontSize:11,fontWeight:600}}>📅 Schedule</button>
          <button onClick={()=>setShowExport(true)} style={{background:'#10B981',color:'white',border:'none',borderRadius:6,padding:'5px 12px',cursor:'pointer',fontSize:11,fontWeight:600}}>Export Excel</button>
        </div>
      </div>

      <div style={{display:'flex',gap:1,marginBottom:10,borderBottom:'2px solid #F0F4F8'}}>
        {[['today','Today'],['entry','Entry'],['review','Review'],['days','Days'],['losses','Losses'],['eob','EOB']].map(([m,label])=>(
          <button key={m} onClick={()=>setMode(m)}
            style={{background:'none',border:'none',borderBottom:mode===m?'2px solid #1B3A5C':'2px solid transparent',color:mode===m?'#1B3A5C':'#94A3B8',padding:'6px 14px',cursor:'pointer',fontSize:12,fontWeight:mode===m?700:400,fontFamily:"'DM Sans',sans-serif",marginBottom:-2}}>
            {label}
          </button>
        ))}
      </div>

      {mode==='today'&&(
        <div>
          <EntryTopBar isToday={true}/>
          {pasteSaving&&(
            <div style={{background:'#FEF3C7',border:'1px solid #FCD34D',borderRadius:6,padding:'6px 12px',marginBottom:8,fontSize:11,color:'#92400E',fontWeight:600,display:'flex',alignItems:'center',gap:6}}>
              <span>⏳ Saving {pasteCount} patient rows to database...</span>
            </div>
          )}
          {showPaste&&(
            <div style={{background:'#EFF6FF',border:'1px solid #BFDBFE',borderRadius:8,padding:'10px 12px',margin:'0 0 10px 0'}}>
              <p style={{fontSize:11,color:'#1D4ED8',fontWeight:600,marginBottom:6}}>Paste TAB schedule — patient names extracted in order</p>
              <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)}
                rows={5} placeholder="Paste schedule here..."
                style={{width:'100%',fontSize:11,fontFamily:'monospace',border:'1px solid #BFDBFE',borderRadius:6,padding:'6px 8px',resize:'vertical',boxSizing:'border-box'}}/>
              <div style={{display:'flex',gap:8,alignItems:'center',marginTop:8}}>
                <button onClick={()=>{const n=parseScheduleText(pasteText);if(n.length>0){setPasteCount(n.length);setPasteSaving(true);setTodayInitCount(0);setTimeout(()=>{setTodayInitCount(n.length);setPasteNames(n);},50);}setShowPaste(false);setPasteText('');}}
                  style={{background:'#1D4ED8',color:'white',border:'none',borderRadius:6,padding:'6px 14px',cursor:'pointer',fontSize:12,fontWeight:600}}>
                  Generate rows
                </button>
                <button onClick={()=>{setShowPaste(false);setPasteText('');}}
                  style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'6px 10px',cursor:'pointer',fontSize:12,color:'#64748B'}}>Cancel</button>
              </div>
            </div>
          )}
          <EntryTable key="today" entries={entries} lockedDate={today()} lockedLoc={todayLoc} lockedDoctor={todayDoctor}
            feeSettings={feeSettings} otherItems={otherItems} onSave={handleSave} onDelete={handleDelete}
            queue={queue} onSwap={swapFromQueue} onRemove={removeFromQueue} initCount={todayInitCount}
                pasteNames={pasteNames} onPasteConsumed={()=>{setPasteNames([]);setPasteSaving(false);}}/>
        </div>
      )}

      {mode==='entry'&&(
        <div>
          <EntryTopBar isToday={false}/>
          {showPaste&&(
            <div style={{background:'#EFF6FF',border:'1px solid #BFDBFE',borderRadius:8,padding:'10px 12px',margin:'0 0 10px 0'}}>
              <p style={{fontSize:11,color:'#1D4ED8',fontWeight:600,marginBottom:6}}>Paste TAB schedule — patient names extracted in order</p>
              <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)}
                rows={5} placeholder="Paste schedule here..."
                style={{width:'100%',fontSize:11,fontFamily:'monospace',border:'1px solid #BFDBFE',borderRadius:6,padding:'6px 8px',resize:'vertical',boxSizing:'border-box'}}/>
              <div style={{display:'flex',gap:8,alignItems:'center',marginTop:8}}>
                <button onClick={()=>{const n=parseScheduleText(pasteText);if(n.length>0){setPasteCount(n.length);setPasteSaving(true);setEntryInitCount(0);setTimeout(()=>{setEntryInitCount(n.length);setPasteNames(n);},50);}setShowPaste(false);setPasteText('');}}
                  style={{background:'#1D4ED8',color:'white',border:'none',borderRadius:6,padding:'6px 14px',cursor:'pointer',fontSize:12,fontWeight:600}}>
                  Generate rows
                </button>
                <button onClick={()=>{setShowPaste(false);setPasteText('');}}
                  style={{background:'#F1F5F9',border:'none',borderRadius:6,padding:'6px 10px',cursor:'pointer',fontSize:12,color:'#64748B'}}>Cancel</button>
              </div>
            </div>
          )}
          <SquarePanel
            date={entryDate}
            loc={entryLoc}
            entries={entries.filter(e=>e.date===entryDate&&e.location===entryLoc)}
            onEntryUpdated={upd=>setEntries(prev=>prev.map(e=>e.id===upd.id?upd:e))}
          />
          <EntryTable key="entry" entries={entries} lockedDate={entryDate} lockedLoc={entryLoc} lockedDoctor={entryDoctor}
            feeSettings={feeSettings} otherItems={otherItems} onSave={handleSave} onDelete={handleDelete}
            queue={queue} onSwap={swapFromQueue} onRemove={removeFromQueue} onQueueCurrent={queueCurrent} initCount={entryInitCount}
                pasteNames={pasteNames} onPasteConsumed={()=>{setPasteNames([]);setPasteSaving(false);}}/>
        </div>
      )}

      {mode==='review'&&(
        <div>
          <div style={{display:'flex',gap:5,marginBottom:8,flexWrap:'wrap',alignItems:'center',background:'white',padding:'7px 10px',borderRadius:8,border:'1px solid #E2E8F0'}}>
            <div style={{display:'flex',gap:3,flexWrap:'wrap'}}>
              {[['thisweek','This Wk'],['lastweek','Last Wk'],['thismonth','This Mo'],['twomonths','2 Mo'],['all','All']].map(([v,l])=>(
                <button key={v} onClick={()=>{setReviewFilter(v);setFilterDate('');if(v==='custom'&&!customMonth)setCustomMonth(new Date().toISOString().slice(0,7));}}
                  style={{background:reviewFilter===v&&!filterDate?'#1B3A5C':'#F1F5F9',color:reviewFilter===v&&!filterDate?'white':'#64748B',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:9,fontWeight:reviewFilter===v&&!filterDate?700:400}}>
                  {l}
                </button>
              ))}
              <div style={{display:'flex',alignItems:'center',gap:0}}>
                <button onClick={()=>{setReviewFilter('custom');setFilterDate('');}}
                  style={{background:reviewFilter==='custom'&&!filterDate?'#1B3A5C':'#F1F5F9',color:reviewFilter==='custom'&&!filterDate?'white':'#64748B',border:'none',borderRadius:'5px 0 0 5px',padding:'3px 8px',cursor:'pointer',fontSize:9,fontWeight:reviewFilter==='custom'&&!filterDate?700:400}}>
                  Month:
                </button>
                <input type="month" value={customMonth} 
                  onChange={e=>{setCustomMonth(e.target.value);setReviewFilter('custom');setFilterDate('');}}
                  onFocus={()=>{if(!customMonth)setCustomMonth(new Date().toISOString().slice(0,7));}}
                  style={{...inp,padding:'2px 5px',fontSize:9,borderRadius:'0 5px 5px 0',width:108,borderColor:reviewFilter==='custom'&&!filterDate?'#1B3A5C':'#E2E8F0'}}/>
              </div>
            </div>
            {/* Specific date picker */}
            <div style={{display:'flex',alignItems:'center',gap:4}}>
              <span style={{fontSize:9,color:'#94A3B8'}}>Day:</span>
              <input type="date" value={filterDate} onChange={e=>{setFilterDate(e.target.value);if(e.target.value)setCustomMonth('');}}
                style={{...inp,padding:'3px 6px',fontSize:9,borderColor:filterDate?'#2E7D8C':'#E2E8F0'}}/>
              {filterDate&&<button onClick={()=>setFilterDate('')} style={{background:'none',border:'none',color:'#94A3B8',cursor:'pointer',fontSize:11,padding:'0 2px'}}>✕</button>}
            </div>
            <input placeholder="Search..." value={searchName} onChange={e=>setSearchName(e.target.value)} style={{...inp,width:120,fontSize:10}}/>
            <select value={filterLoc} onChange={e=>setFilterLoc(e.target.value)} style={{...inp,fontSize:10}}><option value="">All Locs</option>{LOCATIONS.map(l=><option key={l} value={l}>{l}</option>)}</select>
            <select value={filterDoctor} onChange={e=>setFilterDoctor(e.target.value)} style={{...inp,fontSize:10}}>
              <option value="">All Drs</option>
              {[...doctorList].sort().map(d=><option key={d} value={d}>{d}</option>)}
              {/* Former doctors — appear in data but not in active list */}
              {(()=>{
                const former=[...new Set(entries.map(e=>e.doctorId).filter(d=>d&&!doctorList.includes(d)))].sort();
                if(!former.length) return null;
                return [
                  <option key="__sep" disabled>── Former ──</option>,
                  ...former.map(d=><option key={d} value={d}>{d}</option>)
                ];
              })()}
            </select>
            <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{...inp,fontSize:10}}>
              <option value="">All Status</option><option value="pending">Pending</option><option value="partial">Partial</option><option value="flagged">Flagged</option><option value="done">Done</option>
            </select>
            <select value={filterCash} onChange={e=>setFilterCash(e.target.value)} style={{...inp,fontSize:10}}><option value="">Cash: all</option><option value="expected">Cash expected</option></select>
            <select value={filterIns} onChange={e=>setFilterIns(e.target.value)} style={{...inp,fontSize:10}}><option value="">Ins: all</option><option value="pending">Pend</option><option value="confirmed">Confirmed</option><option value="bad">Bad</option><option value="bad-done">Bad/Done</option></select>
            {(filterLoc||filterDoctor||filterStatus||filterCash||filterIns||searchName||filterDate)&&<button onClick={()=>{setFilterLoc('');setFilterDoctor('');setFilterStatus('');setFilterCash('');setFilterIns('');setSearchName('');setFilterDate('');setReviewFilter('twomonths');setSearchInput('');}} style={{background:'none',border:'none',color:'#94A3B8',cursor:'pointer',fontSize:10}}>Clear</button>}
            <span style={{fontSize:9,color:'#94A3B8'}}>{filtered.length}</span>
            {filtered.filter(e=>e.attn&&!e.attnResolved).length>0&&(
              <button onClick={async()=>{
                const toResolve=filtered.filter(e=>e.attn&&!e.attnResolved);
                if(!window.confirm(`Resolve all ${toResolve.length} ATTN items in current filter?`))return;
                for(const e of toResolve){
                  const u={...e,attnResolved:true,updatedAt:new Date().toISOString()};
                  await setDoc(doc(db,'billingEntries',e.id),u);
                  setEntries(prev=>prev.map(x=>x.id===e.id?u:x));
                }
              }} style={{background:'#F5F3FF',border:'1px solid #DDD6FE',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:9,color:'#6D28D9',fontWeight:600,whiteSpace:'nowrap'}}>
                Resolve All ATTNs in filter ({filtered.filter(e=>e.attn&&!e.attnResolved).length})
              </button>
            )}
            <div style={{marginLeft:'auto',display:'flex',gap:4}}>
              <button onClick={()=>setShowSvcs(v=>!v)} style={{background:showSvcs?'#1B3A5C':'#F1F5F9',color:showSvcs?'white':'#64748B',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:9,fontWeight:600}}>Svcs</button>
              <button onClick={()=>setShowNotes(v=>!v)} style={{background:showNotes?'#1B3A5C':'#F1F5F9',color:showNotes?'white':'#64748B',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:9,fontWeight:600}}>ATTN/Notes</button>

            </div>
          </div>

          {showColSettings&&(
            <div style={{background:'#F8FAFC',border:'1px solid #E2E8F0',borderRadius:8,padding:'10px 14px',marginBottom:8,display:'flex',gap:12,flexWrap:'wrap',alignItems:'center'}}>
              <span style={{fontSize:11,fontWeight:700,color:'#1B3A5C'}}>Column Widths (px):</span>
              {[['patient','Patient',83],['date','Date',60],['loc','Loc',32],['doctor','Doctor',45],['ptPaid','Pt Paid',52],['ins','Ins Exp',52],['total','Total',52]].map(([k,label,def])=>(
                <label key={k} style={{display:'flex',flexDirection:'column',gap:2,alignItems:'center'}}>
                  <span style={{fontSize:9,color:'#94A3B8'}}>{label}</span>
                  <input type="number" value={cw('review',k,def)} min={20} max={300}
                    onChange={e=>setColWidth('review',k,e.target.value)}
                    style={{width:52,padding:'2px 4px',border:'1px solid #E2E8F0',borderRadius:4,fontSize:11,textAlign:'center'}}/>
                </label>
              ))}
              <button onClick={()=>{setColWidths(p=>{const u={...p};delete u.review;localStorage.setItem('sparkColWidths',JSON.stringify(u));return u;})}
              } style={{background:'#FEE2E2',border:'none',borderRadius:5,padding:'3px 8px',cursor:'pointer',fontSize:10,color:'#DC2626'}}>Reset</button>
            </div>
          )}
          {/* Day summary card when a specific date is selected */}
          {filterDate&&filtered.length>0&&(()=>{
            const totPt=filtered.reduce((s,e)=>s+(parseFloat(e.ptPaid)||0),0);
            const totIns=filtered.reduce((s,e)=>s+(parseFloat(e.ins)||0),0);
            const totInsPaid=filtered.reduce((s,e)=>s+(parseFloat(e.insurancePaid1)||0),0);
            const totTotal=filtered.reduce((s,e)=>{const p=parseFloat(e.ptPaid)||0;const ip=parseFloat(e.insurancePaid1)||0;const ie=parseFloat(e.ins)||0;return s+p+(ip>0?ip:ie);},0);
            const totErr=filtered.reduce((s,e)=>s+(parseFloat(e.paymentErrorLoss)||0),0);
            return (
              <div style={{display:'flex',gap:10,marginBottom:10,flexWrap:'wrap'}}>
                {[
                  ['Patients',filtered.length,'#1B3A5C',false],
                  ['Pt Collected',totPt,'#16A34A',true],
                  ['Ins Expected',totIns,'#1D4ED8',true],
                  ['Ins Paid',totInsPaid,'#2E7D8C',true],
                  ['Total',totTotal,'#1B3A5C',true],
                  ...(totErr>0?[['Err Loss',totErr,'#EF4444',true]]:[]),
                ].map(([label,val,color,isMoney])=>(
                  <div key={label} style={{background:'white',border:'1px solid #E2E8F0',borderRadius:8,padding:'8px 14px',minWidth:90}}>
                    <div style={{fontSize:9,color:'#94A3B8',fontWeight:700,textTransform:'uppercase',marginBottom:3}}>{label}</div>
                    <div style={{fontSize:16,fontWeight:800,color}}>{isMoney?fmt$(val):val}</div>
                  </div>
                ))}
              </div>
            );
          })()}

          {searchName&&filtered.length>0?(
            <div>
              <p style={{fontSize:11,color:'#64748B',marginBottom:8,padding:'0 4px'}}>{filtered.length} result{filtered.length!==1?'s':''} for "{searchName}"</p>
              <ReviewTable items={filtered}/>
            </div>
          ):reviewFilter==='custom'&&groupedByWeek&&!filterDate?(
            <div>
              {Object.entries(groupedByWeek).map(([wKey,week])=>{
                const wPt=week.entries.reduce((s,e)=>s+(parseFloat(e.ptPaid)||0),0);
                const wIns=week.entries.reduce((s,e)=>s+(parseFloat(e.ins)||0),0);
                const wInsPaid=week.entries.reduce((s,e)=>s+(parseFloat(e.insurancePaid1)||0),0);
                const wTotal=week.entries.reduce((s,e)=>{const p=parseFloat(e.ptPaid)||0;const ip=parseFloat(e.insurancePaid1)||0;const ie=parseFloat(e.ins)||0;return s+p+(ip>0?ip:ie);},0);
                return (
                  <div key={wKey} style={{marginBottom:10}}>
                    <button onClick={()=>setExpandedRows(prev=>({...prev,[wKey]:!prev[wKey]}))}
                      style={{background:'#F1F5F9',color:'#1B3A5C',border:'none',borderRadius:7,padding:'5px 14px',cursor:'pointer',fontSize:11,fontWeight:700,width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
                      <span>{wKey} — {week.label} <span style={{fontWeight:400,color:'#94A3B8',fontSize:10}}>({week.entries.length} pts)</span></span>
                      <span style={{display:'flex',gap:12,alignItems:'center',fontSize:10,fontWeight:600}}>
                        <span style={{color:'#1B3A5C',fontWeight:800}}>{fmt$(wTotal)}</span>
                        <span style={{color:'#94A3B8',fontWeight:400}}>{expandedRows[wKey]!==false?'▼':'▶'}</span>
                      </span>
                    </button>
                    {expandedRows[wKey]!==false&&<ReviewTable items={week.entries}/>}
                  </div>
                );
              })}
            </div>
          ):(reviewFilter==='thisweek'||reviewFilter==='lastweek'||reviewFilter==='thismonth'||reviewFilter==='twomonths'||reviewFilter==='custom'?(
            <WeekDayView items={filtered}/>
          ):reviewFilter==='all'?(
            <AllDrillDown items={filtered} expandedRows={expandedRows} setExpandedRows={setExpandedRows} ReviewTable={ReviewTable}/>
          ):<ReviewTable items={filtered}/>)}
        </div>
      )}

      {mode==='days'&&(
        <DaysTab entries={entries}
          onStartDay={(date,loc,doctor)=>{setEntryDate(date);setEntryLoc(loc);setEntryDoctor(doctor||'');setEntryInitCount(0);setMode('entry');}}
          onViewDay={(date,loc)=>{setReviewFilter('custom');setCustomMonth(date.slice(0,7));setFilterLoc(loc);setFilterDate('');setMode('review');}}
          onAddToQueue={addToQueue}
          scheduleTemplate={scheduleTemplate}
          onEntryUpdated={upd=>{setEntries(prev=>prev.map(e=>e.id===upd.id?upd:e));}}
        />
      )}

      {mode==='losses'&&<LossReport entries={entries} doctorList={doctorList}/>}
      {mode==='eob'&&<EOBScreen entries={entries} onApplied={updatedEntries=>{
        setEntries(prev=>prev.map(e=>{const u=updatedEntries.find(u=>u.id===e.id);return u||e;}));
      }}/>}
      {showExport&&<ExportModal entries={entries} onClose={()=>setShowExport(false)} exportFn={exportExcel}/>}
      {editEntry&&<EditModal entry={editEntry} onSave={updated=>{handleSave(updated);setEditEntry(null);}} onClose={()=>setEditEntry(null)} doctorLogs={doctorLogs} feeSettings={feeSettings} otherItems={otherItems}/>}
      {resolveEntry&&<ResolveModal entry={resolveEntry} onSave={updated=>{handleSave(updated);setResolveEntry(null);}} onClose={()=>setResolveEntry(null)}/>}
      {false&&<EOBImportModal entries={entries} onClose={()=>setShowEOB(false)} onApplied={async(updatedEntries)=>{
        // Moved to EOB tab
        if(updatedEntries&&updatedEntries.length>0){
          setEntries(prev=>{
            const map=Object.fromEntries(prev.map(e=>[e.id,e]));
            for(const u of updatedEntries) map[u.id]=u;
            return Object.values(map);
          });
        }
      }}/> }
      {showCalendar&&<CalendarModal onClose={()=>{setShowCalendar(false);loadScheduleTemplate().then(setScheduleTemplate);}} doctorList={doctorList}/>}
    </div>
  );
}

// ── EOB Screen (tab in billing sheet) ────────────────────────────────────────

function PendingClaimRow({claim, idx, batch, entries, searchOpen, setSearchOpen, searchText, setSearchText, approving, dismissClaim, approveClaim, guessPayorSlot, onNavigate}) {
  const key = batch.id+'_'+idx;
  const suggestedEntry = claim.suggestedEntryId ? entries.find(e=>e.id===claim.suggestedEntryId) : null;
  const [linkedEntryId, setLinkedEntryId] = React.useState(claim.suggestedEntryId||null);
  const linkedEntry = linkedEntryId ? entries.find(e=>e.id===linkedEntryId) : suggestedEntry;
  // Use whichever entry is actively linked — fall back to suggested
  const activeEntry = linkedEntry || suggestedEntry;
  // ins1Amt/ins2Amt may not be set on old spreadsheet imports; fall back to ins total for P1
  const insTotal = activeEntry ? parseFloat(activeEntry.ins)||0 : 0;
  const ins1Amt  = activeEntry ? (parseFloat(activeEntry.ins1Amt)||0) : 0;
  const ins2Amt  = activeEntry ? (parseFloat(activeEntry.ins2Amt)||0) : 0;
  const ins3Amt  = activeEntry ? (parseFloat(activeEntry.ins3Amt)||0) : 0;
  // insExp: use split amounts if present, otherwise fall back to total ins
  const insExp   = ins1Amt > 0 ? ins1Amt + ins2Amt + ins3Amt : insTotal;
  const guessedSlot = guessPayorSlot(claim, activeEntry);
  const [selectedSlot, setSelectedSlot] = React.useState(guessedSlot);
  const linkKey = batch.id+'_'+idx+'_link';
  const isSearchOpen = searchOpen[linkKey];
  const sText = searchText[linkKey]||'';
  const searchResults = sText.length > 1
    ? entries.filter(e=>e.patientName?.toLowerCase().includes(sText.toLowerCase())).slice(0,6)
    : [];
  const [showDetail, setShowDetail] = React.useState(false);
  const [detailNote, setDetailNote] = React.useState('');
  const [showSlotSplit, setShowSlotSplit] = React.useState(false);
  const [slotAmounts, setSlotAmounts] = React.useState({p1:'',p2:'',p3:''});
  const inp = {padding:'4px 6px',border:'1px solid #E2E8F0',borderRadius:5,fontSize:10,fontFamily:"'DM Sans',sans-serif"};

  return (
    <tr style={{borderBottom:'1px solid #F0F4F8',background:idx%2===0?'white':'#FAFBFC',verticalAlign:'top'}}>
      <td style={{padding:'5px 6px',maxWidth:140}}>
        <div style={{fontSize:10,fontWeight:600,color:claim.status==='unmatched'?'#64748B':'#7E22CE'}}>
          {claim.status==='unmatched'?'? No match':'~ Suggested'}
        </div>
        <div style={{fontSize:9,color:'#94A3B8',marginTop:1}}>{claim.reason||claim.status}</div>
      </td>
      <td style={{padding:'5px 6px',fontWeight:600,whiteSpace:'nowrap'}}>{claim.memberName}</td>
      <td style={{padding:'5px 6px',whiteSpace:'nowrap',color:'#64748B'}}>{claim.dos}</td>
      <td style={{padding:'5px 6px',textAlign:'right',fontWeight:700,color:'#1B3A5C'}}>${claim.amount?.toFixed(2)}</td>
      <td style={{padding:'5px 6px'}}>
        <span style={{background:batch.source==='eyemed'||batch.source==='eyemed_text'?'#FEF3C7':'#EFF6FF',
          color:batch.source==='eyemed'||batch.source==='eyemed_text'?'#92400E':'#1D4ED8',
          borderRadius:4,padding:'1px 5px',fontSize:9,fontWeight:700}}>
          {batch.source==='eyemed'||batch.source==='eyemed_text'?'EyeMed':'Assign.'}
        </span>
      </td>
      <td style={{padding:'5px 6px',minWidth:130}}>
        {linkedEntry ? (
          <div>
            <span onClick={()=>setShowDetail(true)}
              style={{color:'#7E22CE',cursor:'pointer',fontWeight:600,textDecoration:'underline',fontSize:11}}
              title="Click to view entry details">
              {linkedEntry.patientName}
            </span>
            <span style={{fontSize:9,color:'#94A3B8',marginLeft:4}}>({linkedEntry.date})</span>
            <button onClick={()=>{setLinkedEntryId(null);setSearchOpen(p=>({...p,[linkKey]:true}));}}
              style={{marginLeft:4,background:'none',border:'none',cursor:'pointer',fontSize:9,color:'#94A3B8'}}>✎</button>
          </div>
        ) : (
          <span style={{fontStyle:'italic',color:'#94A3B8',fontSize:10}}>— not linked</span>
        )}
        <button onClick={()=>setSearchOpen(p=>({...p,[linkKey]:!p[linkKey]}))}
          style={{background:'#EFF6FF',border:'1px solid #BFDBFE',borderRadius:4,padding:'1px 5px',cursor:'pointer',fontSize:9,color:'#1D4ED8',marginTop:2,display:'block'}}>
          {isSearchOpen?'Cancel':'Link Patient'}
        </button>
        {isSearchOpen&&(
          <div style={{marginTop:3}}>
            <input autoFocus value={sText}
              onChange={e=>setSearchText(p=>({...p,[linkKey]:e.target.value}))}
              placeholder="Search name..." style={{...inp,width:'100%'}}/>
            {searchResults.length>0&&(
              <div style={{background:'white',border:'1px solid #E2E8F0',borderRadius:4,maxHeight:110,overflowY:'auto',marginTop:2,zIndex:10,position:'relative'}}>
                {searchResults.map(e=>(
                  <div key={e.id} onClick={()=>{setLinkedEntryId(e.id);setSearchOpen(p=>({...p,[linkKey]:false}));setSearchText(p=>({...p,[linkKey]:''}));}}
                    style={{padding:'3px 7px',cursor:'pointer',fontSize:11,borderBottom:'1px solid #F0F4F8'}}
                    onMouseEnter={ev=>ev.currentTarget.style.background='#F0F9FF'}
                    onMouseLeave={ev=>ev.currentTarget.style.background='white'}>
                    {e.patientName} <span style={{color:'#94A3B8',fontSize:9}}>({e.date})</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </td>
      <td style={{padding:'5px 6px',textAlign:'right',color:insExp>0?'#1B3A5C':'#94A3B8'}}>
        {insExp>0?`$${insExp.toFixed(2)}`:'—'}
      </td>
      {[1,2,3].map(slot=>{
        const slotPayor = slot===1?activeEntry?.payor1:slot===2?activeEntry?.payor2:activeEntry?.payor3;
        // For old entries with only ins (no ins1Amt/ins2Amt), show total on P1
        const slotAmt = slot===1?(ins1Amt||insTotal):slot===2?ins2Amt:ins3Amt;
        if(!slotPayor) return <td key={slot} style={{padding:'5px 6px',color:'#E2E8F0',fontSize:10}}>—</td>;
        // Trigger slot split when there are 2+ payors but amounts not yet broken out
        const needsSplit = activeEntry?.payor2 && ins1Amt===0 && ins2Amt===0 && insTotal>0;
        return (
          <td key={slot} style={{padding:'5px 6px'}}>
            <button onClick={()=>{setSelectedSlot(slot);if(needsSplit)setShowSlotSplit(true);}}
              style={{background:selectedSlot===slot?'#1B3A5C':'#F1F5F9',
                color:selectedSlot===slot?'white':'#64748B',
                border:selectedSlot===slot?'none':'1px solid #E2E8F0',
                borderRadius:5,padding:'2px 5px',cursor:'pointer',fontSize:9,fontWeight:selectedSlot===slot?700:400,
                display:'block',width:'100%',textAlign:'left'}}>
              <div>{slotPayor}</div>
              {slotAmt>0&&<div style={{opacity:0.7}}>${slotAmt.toFixed(2)}</div>}
            </button>
          </td>
        );
      })}
      <td style={{padding:'5px 6px'}}>
        {activeEntry&&(
          <button onClick={()=>approveClaim(batch.id,idx,activeEntry.id,selectedSlot)}
            disabled={approving[key]}
            style={{background:'#10B981',color:'white',border:'none',borderRadius:5,padding:'4px 8px',cursor:'pointer',fontSize:10,fontWeight:700,whiteSpace:'nowrap'}}>
            {approving[key]?'...':'✓ Approve'}
          </button>
        )}
      </td>
      <td style={{padding:'5px 6px'}}>
        <button onClick={()=>dismissClaim(batch.id,idx)}
          style={{background:'#FEE2E2',border:'none',borderRadius:5,padding:'4px 6px',cursor:'pointer',fontSize:10,color:'#DC2626'}}>
          ✕
        </button>
      </td>
      {/* Detail modal + slot split modal rendered outside tr via React portal-like pattern */}
      {showDetail&&activeEntry&&(
        <td colSpan={0} style={{padding:0,border:'none'}}>
          <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.4)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center'}}
            onClick={()=>setShowDetail(false)}>
            <div style={{background:'white',borderRadius:14,padding:24,maxWidth:480,width:'90%',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}
              onClick={e=>e.stopPropagation()}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
                <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:'#1B3A5C',margin:0}}>{activeEntry.patientName}</h3>
                <button onClick={()=>setShowDetail(false)} style={{background:'none',border:'none',cursor:'pointer',fontSize:18,color:'#94A3B8'}}>✕</button>
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:12,fontSize:12}}>
                <div><span style={{color:'#94A3B8'}}>Date: </span>{activeEntry.date}</div>
                <div><span style={{color:'#94A3B8'}}>Location: </span>{activeEntry.location}</div>
                <div><span style={{color:'#94A3B8'}}>Doctor: </span>{activeEntry.doctorId}</div>
                <div><span style={{color:'#94A3B8'}}>Ins Exp: </span>{insExp>0?`$${insExp.toFixed(2)}`:'—'}</div>
                <div><span style={{color:'#94A3B8'}}>P1: </span>{activeEntry.payor1||'—'}</div>
                <div><span style={{color:'#94A3B8'}}>P2: </span>{activeEntry.payor2||'—'}</div>
                {activeEntry.payor3&&<div><span style={{color:'#94A3B8'}}>P3: </span>{activeEntry.payor3}</div>}
                <div><span style={{color:'#94A3B8'}}>EOB $: </span>${claim.amount?.toFixed(2)}</div>
              </div>
              <div style={{marginBottom:12}}>
                <label style={{fontSize:11,color:'#64748B',fontWeight:600,display:'block',marginBottom:4}}>Notes for this match:</label>
                <textarea value={detailNote} onChange={e=>setDetailNote(e.target.value)} rows={3}
                  placeholder="e.g. Wrong DOS entered by doctor, deferred payment from prior visit..."
                  style={{width:'100%',border:'1px solid #E2E8F0',borderRadius:6,padding:'6px 8px',fontSize:12,resize:'vertical',boxSizing:'border-box'}}/>
              </div>
              <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
                <button onClick={()=>setShowDetail(false)}
                  style={{background:'#F1F5F9',border:'none',borderRadius:7,padding:'8px 16px',cursor:'pointer',fontSize:12,color:'#64748B'}}>Cancel</button>
                <button onClick={async()=>{
                  const {setDoc,doc}=await import('firebase/firestore');
                  if(detailNote){const upd={...activeEntry,notes:(activeEntry.notes?activeEntry.notes+' | ':'')+detailNote,updatedAt:new Date().toISOString()};await setDoc(doc(db,'billingEntries',activeEntry.id),upd);}
                  setShowDetail(false);
                  approveClaim(batch.id,idx,activeEntry.id,selectedSlot);
                }} style={{background:'#10B981',color:'white',border:'none',borderRadius:7,padding:'8px 16px',cursor:'pointer',fontSize:12,fontWeight:700}}>
                  ✓ Approve Match
                </button>
                <button onClick={()=>{setShowDetail(false);onNavigate(activeEntry.patientName);}}
                  style={{background:'#EFF6FF',border:'none',borderRadius:7,padding:'8px 16px',cursor:'pointer',fontSize:12,color:'#1D4ED8'}}>
                  Open in Review →
                </button>
              </div>
            </div>
          </div>
        </td>
      )}
      {showSlotSplit&&activeEntry&&(
        <td colSpan={0} style={{padding:0,border:'none'}}>
          <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.4)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center'}}
            onClick={()=>setShowSlotSplit(false)}>
            <div style={{background:'white',borderRadius:14,padding:20,maxWidth:360,width:'90%',boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}}
              onClick={e=>e.stopPropagation()}>
              <h3 style={{fontFamily:"'DM Serif Display',serif",fontSize:15,color:'#1B3A5C',marginBottom:4}}>Set Individual Insurance Amounts</h3>
              <p style={{fontSize:11,color:'#94A3B8',marginBottom:14}}>
                Total Ins Exp: ${insTotal.toFixed(2)}. Set individual amounts for each payor — they should add up to the total.
              </p>
              {['payor1','payor2','payor3'].map((pk,si)=>{
                const payor=activeEntry[pk];
                if(!payor)return null;
                const slotKey=['p1','p2','p3'][si];
                return(
                  <div key={pk} style={{display:'flex',alignItems:'center',gap:8,marginBottom:8}}>
                    <span style={{fontSize:12,fontWeight:600,minWidth:80}}>{payor}</span>
                    <input type="text" inputMode="decimal" value={slotAmounts[slotKey]||''}
                      onChange={e=>setSlotAmounts(prev=>({...prev,[slotKey]:e.target.value}))}
                      placeholder="$0.00" style={{flex:1,padding:'6px 8px',border:'1px solid #E2E8F0',borderRadius:6,fontSize:12}}/>
                    {si===selectedSlot-1&&<span style={{fontSize:10,color:'#10B981',fontWeight:700}}>← this EOB</span>}
                  </div>
                );
              })}
              <div style={{fontSize:11,color:
                Math.abs((parseFloat(slotAmounts.p1)||0)+(parseFloat(slotAmounts.p2)||0)+(parseFloat(slotAmounts.p3)||0)-insTotal)<0.01
                ?'#10B981':'#EF4444',marginBottom:12}}>
                Sum: ${((parseFloat(slotAmounts.p1)||0)+(parseFloat(slotAmounts.p2)||0)+(parseFloat(slotAmounts.p3)||0)).toFixed(2)} / ${insTotal.toFixed(2)}
              </div>
              <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
                <button onClick={()=>setShowSlotSplit(false)}
                  style={{background:'#F1F5F9',border:'none',borderRadius:7,padding:'7px 14px',cursor:'pointer',fontSize:12,color:'#64748B'}}>Cancel</button>
                <button onClick={async()=>{
                  const {setDoc,doc}=await import('firebase/firestore');
                  const upd={...activeEntry,
                    ins1Amt:slotAmounts.p1||'',ins2Amt:slotAmounts.p2||'',ins3Amt:slotAmounts.p3||'',
                    updatedAt:new Date().toISOString()};
                  await setDoc(doc(db,'billingEntries',activeEntry.id),upd);
                  // Mutate local copy so slot amounts display immediately without full reload
                  activeEntry.ins1Amt=upd.ins1Amt; activeEntry.ins2Amt=upd.ins2Amt; activeEntry.ins3Amt=upd.ins3Amt;
                  setShowSlotSplit(false);
                }} style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:7,padding:'7px 14px',cursor:'pointer',fontSize:12,fontWeight:700}}>
                  Save Amounts
                </button>
              </div>
            </div>
          </div>
        </td>
      )}
    </tr>
  );
}

function EOBScreen({entries, onApplied}) {
  const [showImport, setShowImport] = React.useState(false);
  const [pending, setPending] = React.useState([]);
  const [loading, setLoading] = React.useState(true);
  const [eobRefs, setEobRefs] = React.useState({vsp:'',vspDate:'',em:'',emDate:''});
  const [expanded, setExpanded] = React.useState({});
  const [rematching, setRematching] = React.useState(false);
  const [searchOpen, setSearchOpen] = React.useState({}); // batchId+idx -> true
  const [searchText, setSearchText] = React.useState({});
  const [approving, setApproving] = React.useState({});

  const loadPending = React.useCallback(async () => {
    try {
      const { getDocs, collection, orderBy, query, getDoc, doc } = await import('firebase/firestore');
      const snap = await getDocs(query(collection(db, 'pendingEOB'), orderBy('importedAt', 'desc')));
      const batches = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setPending(batches.filter(b => b.claims?.length > 0));
      const settingsSnap = await getDoc(doc(db, 'billingSettings', 'eobTracker'));
      if (settingsSnap.exists()) setEobRefs(settingsSnap.data().refs || {vsp:'',vspDate:'',em:'',emDate:''});
    } catch(e) { console.error(e); }
    setLoading(false);
  }, []);

  React.useEffect(() => { loadPending(); }, [loadPending]);

  const saveRefs = async (updated) => {
    const { setDoc, doc } = await import('firebase/firestore');
    await setDoc(doc(db, 'billingSettings', 'eobTracker'), { refs: updated, updatedAt: new Date().toISOString() });
    setEobRefs(updated);
  };

  const dismissClaim = async (batchId, claimIdx) => {
    const { setDoc, deleteDoc, doc } = await import('firebase/firestore');
    const batch = pending.find(b => b.id === batchId);
    if (!batch) return;
    const updatedClaims = batch.claims.filter((_, i) => i !== claimIdx);
    if (updatedClaims.length === 0) {
      await deleteDoc(doc(db, 'pendingEOB', batchId));
    } else {
      await setDoc(doc(db, 'pendingEOB', batchId), { ...batch, claims: updatedClaims });
    }
    setPending(prev => prev.map(b => b.id === batchId ? { ...b, claims: updatedClaims } : b).filter(b => b.claims?.length > 0));
  };

  const dismissBatch = async (batchId) => {
    if (!window.confirm('Dismiss all claims in this batch?')) return;
    const { deleteDoc, doc } = await import('firebase/firestore');
    await deleteDoc(doc(db, 'pendingEOB', batchId));
    setPending(prev => prev.filter(b => b.id !== batchId));
  };

  const approveClaim = async (batchId, claimIdx, entryId, payorSlot) => {
    const key = batchId + '_' + claimIdx;
    setApproving(prev => ({...prev, [key]: true}));
    try {
      const { setDoc, doc } = await import('firebase/firestore');
      const batch = pending.find(b => b.id === batchId);
      const claim = batch?.claims[claimIdx];
      const entry = entries.find(e => e.id === entryId);
      if (!claim || !entry) return;
      if (payorSlot === 2) entry.insurancePaid2 = String(claim.amount);
      else if (payorSlot === 3) entry.insurancePaid3 = String(claim.amount);
      else entry.insurancePaid1 = String(claim.amount);
      // Check if all payors now paid
      const allPaid = (!entry.payor1||entry.payor1==='Self'||!!entry.insurancePaid1) &&
                      (!entry.payor2||!!entry.insurancePaid2) &&
                      (!entry.payor3||!!entry.insurancePaid3);
      const upd = {...entry, insPaidState: allPaid ? 'confirmed' : 'partial', updatedAt: new Date().toISOString()};
      await setDoc(doc(db, 'billingEntries', entry.id), upd);
      onApplied([upd]);
      await dismissClaim(batchId, claimIdx);
    } catch(e) { alert('Error: ' + e.message); }
    setApproving(prev => ({...prev, [key]: false}));
  };

  // Determine which payor slot the EOB amount likely matches
  const guessPayorSlot = (claim, entry) => {
    if (!entry) return 1;
    const amt = claim.amount;
    const ins1 = parseFloat(entry.ins1Amt) || parseFloat(entry.ins) || 0;
    const ins2 = parseFloat(entry.ins2Amt) || 0;
    const ins3 = parseFloat(entry.ins3Amt) || 0;
    const tol = 2;
    if (ins1 > 0 && Math.abs(amt - ins1) <= tol) return 1;
    if (ins2 > 0 && Math.abs(amt - ins2) <= tol) return 2;
    if (ins3 > 0 && Math.abs(amt - ins3) <= tol) return 3;
    return 1; // default
  };

  const totalPending = pending.reduce((s, b) => s + (b.claims?.length || 0), 0);
  const inp = {padding:'5px 8px',border:'1px solid #E2E8F0',borderRadius:6,fontSize:11,fontFamily:"'DM Sans',sans-serif"};

  return (
    <div style={{maxWidth:1200,margin:'0 auto',padding:'10px 12px'}}>
      {/* Header */}
      <div style={{display:'flex',gap:10,alignItems:'flex-start',marginBottom:16,flexWrap:'wrap'}}>
        <button onClick={()=>setShowImport(true)}
          style={{background:'#1B3A5C',color:'white',border:'none',borderRadius:8,padding:'9px 20px',cursor:'pointer',fontSize:13,fontWeight:700,flexShrink:0}}>
          📥 Import EOB Remittance
        </button>
        {totalPending > 0 && (
          <button onClick={async()=>{
            setRematching(true);
            try {
              const {matchClaimsExternal} = await import('./EOBImport');
              const {setDoc,deleteDoc,doc} = await import('firebase/firestore');
              let totalConfirmed = 0;
              for (const batch of pending) {
                const claims = batch.claims.map(c=>({
                  claimNum:c.claimNum,memberName:c.memberName,
                  last:c.memberName?.split(',')[0]?.trim()||'',
                  first:c.memberName?.split(',')[1]?.trim()||'',
                  doctor:c.doctor||'',dos:c.dos,amount:c.amount,negative:false,source:batch.source
                }));
                const results = matchClaimsExternal(claims, entries);
                const confirmed = results.filter(r=>r.status==='matched'&&r.insPaidState==='confirmed');
                const still = batch.claims.filter((_,i)=>!confirmed.find(r=>r.claim===claims[i]));
                for (const r of confirmed) {
                  const upd={...r.entry,insurancePaid1:String(r.claim.amount),insPaidState:'confirmed',updatedAt:new Date().toISOString()};
                  await setDoc(doc(db,'billingEntries',r.entry.id),upd);
                  onApplied([upd]);
                  totalConfirmed++;
                }
                if (still.length===0) await deleteDoc(doc(db,'pendingEOB',batch.id));
                else await setDoc(doc(db,'pendingEOB',batch.id),{...batch,claims:still});
              }
              await loadPending();
              alert(`Rematch complete. ${totalConfirmed} new matches confirmed.`);
            } catch(e){alert('Rematch error: '+e.message);}
            setRematching(false);
          }} disabled={rematching}
          style={{background:'#7C3AED',color:'white',border:'none',borderRadius:8,padding:'9px 16px',cursor:'pointer',fontSize:12,fontWeight:600,flexShrink:0}}>
            {rematching?'Rematching...':'🔄 Rematch '+totalPending+' pending'}
          </button>
        )}
        {/* Two EOB ref fields */}
        <div style={{display:'flex',gap:12,alignItems:'center',marginLeft:'auto',flexWrap:'wrap'}}>
          <div style={{display:'flex',flexDirection:'column',gap:2}}>
            <span style={{fontSize:9,color:'#94A3B8',textTransform:'uppercase',fontWeight:700}}>Last VSP/Assign. Plan</span>
            <div style={{display:'flex',gap:4}}>
              <input value={eobRefs.vsp||''} onChange={e=>setEobRefs(p=>({...p,vsp:e.target.value}))}
                onBlur={()=>saveRefs(eobRefs)} placeholder="Ref #" style={{...inp,width:90}}/>
              <input type="date" value={eobRefs.vspDate||''} onChange={e=>saveRefs({...eobRefs,vspDate:e.target.value})}
                style={{...inp,width:120}}/>
            </div>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:2}}>
            <span style={{fontSize:9,color:'#94A3B8',textTransform:'uppercase',fontWeight:700}}>Last EyeMed</span>
            <div style={{display:'flex',gap:4}}>
              <input value={eobRefs.em||''} onChange={e=>setEobRefs(p=>({...p,em:e.target.value}))}
                onBlur={()=>saveRefs(eobRefs)} placeholder="Ref #" style={{...inp,width:90}}/>
              <input type="date" value={eobRefs.emDate||''} onChange={e=>saveRefs({...eobRefs,emDate:e.target.value})}
                style={{...inp,width:120}}/>
            </div>
          </div>
        </div>
      </div>

      {/* Pending list */}
      {loading ? (
        <div style={{textAlign:'center',color:'#94A3B8',padding:20}}>Loading...</div>
      ) : totalPending === 0 ? (
        <div style={{background:'#F0FDF4',borderRadius:10,padding:20,textAlign:'center',color:'#166534',fontSize:13}}>
          ✓ No pending EOB claims. All imports resolved.
        </div>
      ) : (
        <div>
          <p style={{fontSize:12,color:'#94A3B8',marginBottom:10}}>
            {totalPending} claim{totalPending!==1?'s':''} pending across {pending.length} import{pending.length!==1?'s':''}
          </p>
          {pending.map(batch=>(
            <div key={batch.id} style={{border:'1px solid #E2E8F0',borderRadius:10,marginBottom:10,overflow:'hidden'}}>
              <div style={{background:'#F8FAFC',padding:'8px 12px',display:'flex',justifyContent:'space-between',alignItems:'center',cursor:'pointer'}}
                onClick={()=>setExpanded(prev=>({...prev,[batch.id]:!prev[batch.id]}))}>
                <div style={{display:'flex',gap:8,alignItems:'center'}}>
                  <span style={{fontWeight:700,fontSize:13,color:'#1B3A5C'}}>
                    {new Date(batch.importedAt).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}
                  </span>
                  <span style={{fontSize:10,color:'#94A3B8',textTransform:'uppercase'}}>{batch.source}</span>
                  <span style={{background:'#FEE2E2',color:'#DC2626',borderRadius:4,padding:'1px 6px',fontSize:11,fontWeight:700}}>
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

              {expanded[batch.id]&&(
                <div style={{overflowX:'auto'}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                    <thead>
                      <tr style={{borderBottom:'1px solid #E2E8F0',background:'#FAFBFC'}}>
                        {['Status/Reason','EOB Patient','DOS','EOB $','Type','Suggested Match','Ins Exp','P1','P2','P3','Approve','Dismiss'].map(h=>(
                          <th key={h} style={{padding:'5px 6px',textAlign:'left',fontSize:9,fontWeight:700,color:'#94A3B8',textTransform:'uppercase',whiteSpace:'nowrap'}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {batch.claims.map((claim,idx)=>(
                        <PendingClaimRow key={claim.claimNum||batch.id+'_'+idx} claim={claim} idx={idx} batch={batch}
                          entries={entries} searchOpen={searchOpen} setSearchOpen={setSearchOpen}
                          searchText={searchText} setSearchText={setSearchText}
                          approving={approving} dismissClaim={dismissClaim} approveClaim={approveClaim}
                          guessPayorSlot={guessPayorSlot}
                          onNavigate={name=>{window._eobNavigate&&window._eobNavigate(name);}}
                        />
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {showImport&&(
        <EOBImportModal
          entries={entries}
          onClose={()=>{setShowImport(false);loadPending();}}
          onApplied={(updatedEntries, importMeta)=>{
            onApplied(updatedEntries);
            // Auto-save disbursement date and ref ID for the appropriate payor type
            if (importMeta?.date) {
              const isEyeMed = importMeta.source === 'eyemed' || importMeta.source === 'eyemed_text';
              const updated = isEyeMed
                ? { ...eobRefs, em: importMeta.refId || eobRefs.em, emDate: importMeta.date }
                : { ...eobRefs, vsp: importMeta.refId || eobRefs.vsp, vspDate: importMeta.date };
              saveRefs(updated);
            }
          }}
        />
      )}
    </div>
  );
}
