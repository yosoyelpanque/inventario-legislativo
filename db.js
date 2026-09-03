const DB_NAME = 'inventario-legislativo-db';
const DB_VERSION = 1;

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains('state')) db.createObjectStore('state');
      if (!db.objectStoreNames.contains('photos')) db.createObjectStore('photos');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function transaction(storeName, mode, action) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    let result;
    try { result = action(store); } catch (error) { db.close(); reject(error); return; }
    tx.oncomplete = () => { db.close(); resolve(result?.result); };
    tx.onerror = () => { db.close(); reject(tx.error); };
    tx.onabort = () => { db.close(); reject(tx.error); };
  });
}

export const storage = {
  load: () => transaction('state', 'readonly', (store) => store.get('main')),
  save: (state) => transaction('state', 'readwrite', (store) => store.put(state, 'main')),
  photo: (key) => transaction('photos', 'readonly', (store) => store.get(key)),
  savePhoto: (key, blob) => transaction('photos', 'readwrite', (store) => store.put(blob, key)),
  deletePhoto: (key) => transaction('photos', 'readwrite', (store) => store.delete(key)),
  clear: async () => {
    const db = await openDb();
    db.close();
    return new Promise((resolve, reject) => {
      const request = indexedDB.deleteDatabase(DB_NAME);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error('Cierre las otras pestañas de Inventario Legislativo e inténtelo de nuevo.'));
    });
  }
};
