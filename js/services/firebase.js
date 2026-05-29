// Firebase helpers: centralizes Firestore/Auth setup and exposes shared handles.
// Requires Firebase SDK + initialization script to be loaded first.
(function (window) {
  if (window.AppFirebase) return;
  if (!window.firebase) {
    throw new Error('Firebase SDK not loaded; ensure firebase-app/auth/firestore scripts load before services/firebase.js');
  }

  const db = firebase.firestore();
  const auth = firebase.auth();
  const FieldValue = firebase.firestore.FieldValue;
  const Timestamp = firebase.firestore.Timestamp;

  // Only log errors to keep console clean
  try { firebase.firestore.setLogLevel('error'); } catch (_) { /* ignore */ }

  // Offline persistence (reduces reads/bandwidth on reloads). Must be called before any reads.
  try {
    firebase.firestore().enablePersistence({ synchronizeTabs: true })
      .catch(err => {
        if (err && (err.code === 'failed-precondition' || err.code === 'unimplemented')) {
          console.warn('Firestore persistence not enabled:', err.code);
        } else {
          console.warn('Firestore persistence error:', err?.message || err);
        }
      });
  } catch (e) {
    console.warn('enablePersistence threw:', e);
  }

  const serverTimestamp = () => FieldValue.serverTimestamp();

  window.AppFirebase = Object.freeze({
    db,
    auth,
    FieldValue,
    Timestamp,
    serverTimestamp,
  });
})(window);
