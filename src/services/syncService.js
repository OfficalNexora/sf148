/**
 * Online real-time sync via Firestore.
 */

import {
    addDoc,
    collection,
    deleteDoc,
    doc,
    onSnapshot,
    query,
    serverTimestamp,
    updateDoc,
    where
} from 'firebase/firestore';
import { getFirebaseDb } from './firebaseClient';

function getDB() {
    return getFirebaseDb();
}

function encode(data) {
    return btoa(unescape(encodeURIComponent(JSON.stringify(data))));
}

function decode(str) {
    return JSON.parse(decodeURIComponent(escape(atob(str))));
}

const MAX_PAYLOAD_BYTES = 800_000;

const syncService = {
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

    listenForPendingSyncs(callback) {
        try {
            const firestore = getDB();
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
                requests.sort((a, b) => b.createdAt - a.createdAt);
                callback(requests);
            }, (err) => {
                console.error('Firestore listener error:', err);
                callback([]);
            });
        } catch (err) {
            console.error('listenForPendingSyncs error:', err);
            return () => { };
        }
    },

    decodeSyncRequest(request) {
        try {
            return decode(request._encoded);
        } catch (err) {
            console.error('decodeSyncRequest error:', err);
            return null;
        }
    },

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
