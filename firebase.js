// firebase.js
// Initializes Firebase, signs in anonymously, and exports useful references.
// It mirrors the behavior from the original single-file version.
// Usage: import { initFirebase } from './firebase.js'; const { db, auth, monitoringCollectionRef, helpers, userId } = await initFirebase();

import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
import { getAuth, signInAnonymously } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
import { getFirestore, collection, doc, onSnapshot, addDoc, updateDoc, deleteDoc, setLogLevel } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

export async function initFirebase() {
  let db = null;
  let auth = null;
  let userId = null;
  let monitoringCollectionRef = null;

  try {
    const appId = typeof __app_id !== 'undefined' ? __app_id : 'sched-master-default';
    const firebaseConfig = JSON.parse(typeof __firebase_config !== 'undefined' ? __firebase_config : '{}');
    const app = initializeApp(firebaseConfig);
    db = getFirestore(app);
    auth = getAuth(app);
    setLogLevel('debug');

    await signInAnonymously(auth);
    userId = auth.currentUser?.uid || null;

    const monitoringPath = `/artifacts/${appId}/public/data/monitoring`;
    monitoringCollectionRef = collection(db, monitoringPath);

    return {
      db,
      auth,
      userId,
      monitoringCollectionRef,
      helpers: { collection, doc, onSnapshot, addDoc, updateDoc, deleteDoc }
    };
  } catch (error) {
    console.error("Firebase initialization failed in firebase.js:", error);
    // rethrow so caller can handle
    throw error;
  }
}
