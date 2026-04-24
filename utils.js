// ── Constants ──────────────────────────────────────────────────────────────
export const LOCATIONS = ['SC', 'F', 'WC', 'SV'];
export const LOC_FULL = { SC: 'Santa Clara', F: 'Fremont', WC: 'Walnut Creek', SV: 'Sunnyvale' };
export const LOC_COLORS = { SC: '#2E7D8C', F: '#1B3A5C', WC: '#C9A84C', SV: '#10B981' };

export const PAYORS = [
  'VSP', 'EyeMed', 'MES Vision', 'UHC/Spectera', 'Davis Vision',
  'Superior Vision', 'NVA', 'Avesis', 'FEP', 'VBA', 'Heritage',
  'Colonial Penn', 'Self Pay', 'Other'
];

export const SERVICE_FIELDS = [
  { key: 'routine',   label: 'Routine',   hint: 'Comprehensive exam' },
  { key: 'cl',        label: 'CL',        hint: 'Contact lens exam' },
  { key: 'optos',     label: 'Optos',     hint: 'Optomap' },
  { key: 'dfe',       label: 'DFE',       hint: 'Dilated fundus exam' },
  { key: 'ov',        label: 'OV',        hint: 'Office visit' },
  { key: 'oct',       label: 'OCT',       hint: 'OCT scan' },
  { key: 'topo',      label: 'Topo',      hint: 'Topography' },
  { key: 'other',     label: 'Other',     hint: 'Other services' },
  { key: 'myopia',    label: 'Myopia',    hint: 'Myopia control (excluded from avg calc)' },
  { key: 'lasik',     label: 'LASIK',     hint: 'LASIK referral' },
  { key: 'materials', label: 'Materials', hint: 'Frames, lenses, contacts' },
];

export const STATUS_OPTIONS = ['pending', 'completed', 'flagged'];

// ── Utilities ──────────────────────────────────────────────────────────────
export function uid() { return Math.random().toString(36).slice(2, 9); }
export function today() { return new Date().toISOString().slice(0, 10); }
export function fmtDate(d) {
  if (!d) return '—';
  return new Date(d + 'T00:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}
export function fmtDay(d) {
  if (!d) return '';
  return new Date(d + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'long' });
}
export function fmt$(n) {
  if (!n && n !== 0) return '—';
  return '$' + parseFloat(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
export function getWeek(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const start = new Date(d.getFullYear(), 0, 1);
  return Math.ceil(((d - start) / 86400000 + start.getDay() + 1) / 7);
}

// Empty entry template
export function emptyEntry(doctorId, doctorName) {
  return {
    id: uid(),
    date: today(),
    location: '',
    patientName: '',
    doctorId: doctorId || '',
    doctorName: doctorName || '',
    routine: '', cl: '', optos: '', dfe: '', ov: '', oct: '',
    topo: '', other: '', myopia: '', lasik: '', materials: '',
    payor1: '', payor2: '',
    cash: '',
    ins: '',
    claimNumber: '',
    notes: '',
    // Biller-only fields
    status: 'pending',
    insurancePaid1: '', insurancePaid2: '',
    paymentErrorLoss: '', insuranceNonpaymentLoss: '',
    attn: '',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}
