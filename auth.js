import { db } from './firebase';
import { doc, getDoc, setDoc, collection, getDocs, deleteDoc } from 'firebase/firestore';
import { createUserWithEmailAndPassword, signInWithEmailAndPassword, updatePassword } from 'firebase/auth';
import { auth } from './firebase';

// ── User roles ─────────────────────────────────────────────────────────────
// master: Christine — sees everything
// biller: Dr. Kha — sees all logs + main sheet
// doctor: individual doctor — sees only their log

export const MASTER_EMAIL = 'cyang.od@gmail.com';

// Initial doctor accounts — you'll set passwords via admin panel
export const INITIAL_USERS = [
  { id: 'kha',    name: 'Dr. Kha',    email: 'kha@thesparkoptometry.app',    role: 'biller' },
  { id: 'pan',    name: 'Dr. Pan',    email: 'pan@thesparkoptometry.app',    role: 'doctor' },
  { id: 'fan',    name: 'Dr. Fan',    email: 'fan@thesparkoptometry.app',    role: 'doctor' },
  { id: 'luong',  name: 'Dr. Luong',  email: 'luong@thesparkoptometry.app',  role: 'doctor' },
  { id: 'kaneta', name: 'Dr. Kaneta', email: 'kaneta@thesparkoptometry.app', role: 'doctor' },
];

export async function loadUserProfile(uid) {
  const snap = await getDoc(doc(db, 'billingUsers', uid));
  return snap.exists() ? snap.data() : null;
}

export async function saveUserProfile(uid, data) {
  await setDoc(doc(db, 'billingUsers', uid), data, { merge: true });
}

export async function loadAllUsers() {
  const snap = await getDocs(collection(db, 'billingUsers'));
  return snap.docs.map(d => ({ uid: d.id, ...d.data() }));
}

export async function deleteUser(uid) {
  await deleteDoc(doc(db, 'billingUsers', uid));
}

export async function signIn(email, password) {
  return await signInWithEmailAndPassword(auth, email, password);
}

export async function changePassword(newPassword) {
  if (!auth.currentUser) throw new Error('Not signed in');
  await updatePassword(auth.currentUser, newPassword);
}
