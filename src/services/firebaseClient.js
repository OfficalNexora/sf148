import { getApp, getApps, initializeApp } from 'firebase/app';
import { getFirestore, initializeFirestore } from 'firebase/firestore';

const firebaseConfig = {
    apiKey: import.meta.env.VITE_FIREBASE_API_KEY || 'AIzaSyBULMeZfo2gLMFgcyT1TdDQlEnHI5mg6Z4',
    authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || 'form137-sync.firebaseapp.com',
    projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || 'form137-sync',
    storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || 'form137-sync.firebasestorage.app',
    messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || '863019085410',
    appId: import.meta.env.VITE_FIREBASE_APP_ID || '1:863019085410:web:1096b14b4318acf3282b3b'
};

let appInstance = null;
let dbInstance = null;

export function getFirebaseApp() {
    if (!appInstance) {
        appInstance = getApps().length ? getApp() : initializeApp(firebaseConfig);
    }
    return appInstance;
}

export function getFirebaseDb() {
    if (!dbInstance) {
        const app = getFirebaseApp();
        try {
            dbInstance = initializeFirestore(app, {
                experimentalForceLongPolling: true
            });
        } catch (err) {
            dbInstance = getFirestore(app);
        }
    }
    return dbInstance;
}
