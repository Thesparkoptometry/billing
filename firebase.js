import { initializeApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyDpWF7_2czyzZaZ-g77yPOE1gxwh5C7wZg",
  authDomain: "spark-doctor-portal.firebaseapp.com",
  projectId: "spark-doctor-portal",
  storageBucket: "spark-doctor-portal.firebasestorage.app",
  messagingSenderId: "989715476181",
  appId: "1:989715476181:web:22352ed6e4b083b8ffbc5c"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
