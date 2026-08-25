// Firebase config - EDIT ONCE, no need to copy every time you upload HTML
// For production, replace with your own Firebase project config (free, no card)
// If you host on Firebase Hosting, this file is ignored - auto-config via /__/firebase/init.js
// Demo config below works for testing (shared public demo) - replace with your own for private trips

window.FIREBASE_CONFIG = {
  apiKey: "AIzaSyDyVy6R9czC0ps6leC4LUhZ-Ew9KqtCwp8",
  authDomain: "expensesplit-gc.firebaseapp.com",
  projectId: "expensesplit-gc",
  storageBucket: "expensesplit-gc.firebasestorage.app",
  messagingSenderId: "1044897195717",
  appId: "1:1044897195717:web:a0b0f8b3b2a4716c1532ac"
};

// Admin password to create trips - change this
window.ADMIN_CREATE_PASSWORD = "ADMIN2026";

// Set to false to completely hide Create Trip from public (only via ?admin=1 URL)
window.ALLOW_PUBLIC_CREATE = false;
