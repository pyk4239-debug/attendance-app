import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

const firebaseConfig = {
  apiKey: "AIzaSyCyQpqprnTfvoFywru5qDh2WorTtM1_zjY",
  authDomain: "attendance-app-f10cb.firebaseapp.com",
  projectId: "attendance-app-f10cb",
  storageBucket: "attendance-app-f10cb.firebasestorage.app",
  messagingSenderId: "18938480382",
  appId: "1:18938480382:web:b7dffb73e92fa41de887af"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const storage = getStorage(app);

