/**
 * syncService.js — Online real-time sync via Firestore (free Spark plan).
 * No Firebase Storage required — sync data is stored directly in Firestore documents.
 *
 * Teacher workflow:  submitSync(name, data)  →  Firestore doc created
 * Admin workflow:    listenForPendingSyncs(cb) → onSnapshot fires instantly
 *                   mergeSyncRequest(req)      → deep-merge into local db
 *                   markAsMerged(id)           → sets status = 'merged'
 */

import { initializeApp } from 'firebase/app';
import {
    getFirestore,
    collection,
    addDoc,
    updateDoc,
    doc,
    query,
    where,
    onSnapshot,
    serverTimestamp,
    deleteDoc
} from 'firebase/firestore';

// ── Firebase Config ──────────────────────────────────────────────────────────
const firebaseConfig = {
    apiKey: "AIzaSyBULMeZfo2gLMFgcyT1TdDQlEnHI5mg6Z4",
    authDomain: "form137-sync.firebaseapp.com",
    projectId: "form137-sync",
    storageBucket: "form137-sync.firebasestorage.app",
    messagingSenderId: "863019085410",
    appId: "1:863019085410:web:1096b14b4318acf3282b3b"
};

let app = null;
let db = null;

function getDB() {
    if (!db) {
        app = initializeApp(firebaseConfig);
        db = getFirestore(app);
    }
    return db;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Compress a JSON object into a base64 string to reduce Firestore document size.
 * This is a lightweight stringify; we don't gzip since we're in-browser.
 */
function encode(data) {
    return btoa(unescape(encodeURIComponent(JSON.stringify(data))));
}

function decode(str) {
    return JSON.parse(decodeURIComponent(escape(atob(str))));
}

// Firestore hard limit is 1MB per document. We'll warn above 800KB.
const MAX_PAYLOAD_BYTES = 800_000;

// ── Public API ────────────────────────────────────────────────────────────────

const syncService = {
    /**
     * Teacher: submit their entire local database as a sync request.
     * @param {string} teacherName - Display name of the sender
     * @param {object} syncData    - { structure, records[] }
     * @returns {{ success: boolean, error?: string }}
     */
    async submitSync(teacherName, syncData) {
        try {
            const encoded = encode(syncData);
            const byteSize = new Blob([encoded]).size;

            if (byteSize > MAX_PAYLOAD_BYTES) {
                return {
                    success: false,
                    error: `Sync file too large (${Math.round(byteSize / 1024)}KB). Maximum is 800KB. Remove some records and try again.`
                };
            }

            const firestore = getDB();
            await addDoc(collection(firestore, 'syncRequests'), {
                teacherName: teacherName || 'Unknown Teacher',
                status: 'pending',
                createdAt: serverTimestamp(),
                sizeBytes: byteSize,
                data: encoded
            });

            return { success: true };
        } catch (err) {
            console.error('submitSync error:', err);
            return { success: false, error: err.message };
        }
    },

    /**
     * Admin: subscribe to pending sync requests in real time.
     * @param {function} callback - called with array of { id, teacherName, createdAt, sizeBytes }
     * @returns {function} unsubscribe function
     */
    listenForPendingSyncs(callback) {
        try {
            const firestore = getDB();
            // Simple query — no orderBy, so no composite index needed.
            // We sort the results client-side instead.
            const q = query(
                collection(firestore, 'syncRequests'),
                where('status', '==', 'pending')
            );

            return onSnapshot(q, (snapshot) => {
                const requests = snapshot.docs.map(d => ({
                    id: d.id,
                    teacherName: d.data().teacherName,
                    createdAt: d.data().createdAt?.toDate?.() || new Date(),
                    sizeBytes: d.data().sizeBytes || 0,
                    _encoded: d.data().data
                }));
                // Sort newest first client-side — no index required
                requests.sort((a, b) => b.createdAt - a.createdAt);
                callback(requests);
            }, (err) => {
                console.error('Firestore listener error:', err);
                callback([]);
            });
        } catch (err) {
            console.error('listenForPendingSyncs error:', err);
            return () => { }; // no-op unsubscribe
        }
    },

    /**
     * Admin: decode and return the raw sync data from a request.
     * @param {object} request - one item from listenForPendingSyncs callback
     * @returns {{ structure, records[] } | null}
     */
    decodeSyncRequest(request) {
        try {
            return decode(request._encoded);
        } catch (err) {
            console.error('decodeSyncRequest error:', err);
            return null;
        }
    },

    /**
     * Admin: mark a sync request as merged in Firestore.
     * @param {string} docId
     */
    async markAsMerged(docId) {
        try {
            const firestore = getDB();
            await updateDoc(doc(firestore, 'syncRequests', docId), {
                status: 'merged',
                mergedAt: serverTimestamp()
            });
            return { success: true };
        } catch (err) {
            console.error('markAsMerged error:', err);
            return { success: false, error: err.message };
        }
    },

    /**
     * Admin: permanently delete a sync request document.
     * @param {string} docId
     */
    async dismissSync(docId) {
        try {
            const firestore = getDB();
            await deleteDoc(doc(firestore, 'syncRequests', docId));
            return { success: true };
        } catch (err) {
            console.error('dismissSync error:', err);
            return { success: false, error: err.message };
        }
    }
};

export default syncService;
