import {
  doc,
  getDoc,
  setDoc,
  deleteDoc,
  collection,
  getDocs,
} from 'firebase/firestore';
import { firestoreDb } from '../../db/firebase.ts';
import { OutlookConnectionDoc } from './outlook.types.ts';
import { OUTLOOK_COLLECTION_NAME } from './outlook.constants.ts';

// In-memory cache for fast access and fallback
const outlookCache = new Map<string, OutlookConnectionDoc>();

function sanitizeForFirestore<T extends Record<string, any>>(obj: T): Record<string, any> {
  const result: Record<string, any> = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value !== undefined) {
      if (value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
        result[key] = sanitizeForFirestore(value);
      } else if (Array.isArray(value)) {
        result[key] = value.filter((v) => v !== undefined);
      } else {
        result[key] = value;
      }
    }
  }
  return result;
}

export async function getOutlookConnection(teacherId: string): Promise<OutlookConnectionDoc | null> {
  if (!teacherId) return null;

  try {
    const docRef = doc(firestoreDb, OUTLOOK_COLLECTION_NAME, `conn_${teacherId}`);
    const snap = await getDoc(docRef);
    if (snap.exists()) {
      const data = snap.data() as OutlookConnectionDoc;
      outlookCache.set(teacherId, data);
      return data;
    }
  } catch (err) {
    console.warn(`[Outlook Repository] Firestore read fallback for teacher ${teacherId}:`, err);
  }

  return outlookCache.get(teacherId) || null;
}

export async function saveOutlookConnection(data: OutlookConnectionDoc): Promise<void> {
  if (!data.teacherId) return;

  outlookCache.set(data.teacherId, data);

  try {
    const docRef = doc(firestoreDb, OUTLOOK_COLLECTION_NAME, `conn_${data.teacherId}`);
    await setDoc(docRef, sanitizeForFirestore(data));
  } catch (err) {
    console.error(`[Outlook Repository] Error persisting connection for teacher ${data.teacherId}:`, err);
  }
}

export async function deleteOutlookConnection(teacherId: string): Promise<void> {
  if (!teacherId) return;

  outlookCache.delete(teacherId);

  try {
    const docRef = doc(firestoreDb, OUTLOOK_COLLECTION_NAME, `conn_${teacherId}`);
    await deleteDoc(docRef);
  } catch (err) {
    console.error(`[Outlook Repository] Error deleting connection for teacher ${teacherId}:`, err);
  }
}
