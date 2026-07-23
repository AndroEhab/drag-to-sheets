/**
 * IndexedDB-backed store for FileSystemFileHandle objects.
 *
 * FileSystemFileHandle cannot be serialized to JSON (chrome.storage),
 * but CAN be stored in IndexedDB via the structured clone algorithm.
 * This module provides a simple key-value store for handles keyed by
 * a generated ID, plus helpers to verify/request write permission.
 */

// eslint-disable-next-line no-unused-vars
const FileHandleStore = (() => {
  'use strict';

  const DB_NAME = 'drag-to-sheets-handles';
  const DB_VERSION = 1;
  const STORE_NAME = 'file-handles';

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error('Failed to open handle store'));
    });
  }

  function withStore(mode, fn) {
    return openDb().then((db) => {
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, mode);
        const store = tx.objectStore(STORE_NAME);
        const result = fn(store);
        tx.oncomplete = () => { db.close(); resolve(result); };
        tx.onerror = () => { db.close(); reject(tx.error); };
        tx.onabort = () => { db.close(); reject(tx.error || new Error('Transaction aborted')); };
      });
    });
  }

  /** Generate a short random ID for a handle entry. */
  function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  /**
   * Store a FileSystemFileHandle and return its generated ID.
   * @param {FileSystemFileHandle} handle
   * @returns {Promise<string>} The handle ID
   */
  async function saveHandle(handle) {
    const id = generateId();
    await withStore('readwrite', (store) => {
      store.put(handle, id);
    });
    return id;
  }

  /**
   * Retrieve a FileSystemFileHandle by ID.
   * @param {string} id
   * @returns {Promise<FileSystemFileHandle|null>}
   */
  async function getHandle(id) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const request = tx.objectStore(STORE_NAME).get(id);
      request.onsuccess = () => { db.close(); resolve(request.result || null); };
      request.onerror = () => { db.close(); reject(request.error); };
    });
  }

  /**
   * Delete a stored handle by ID.
   * @param {string} id
   */
  async function deleteHandle(id) {
    await withStore('readwrite', (store) => {
      store.delete(id);
    });
  }

  /**
   * Verify write permission on a handle, requesting if needed.
   * Must be called from a user-gesture context the first time.
   * @param {FileSystemFileHandle} handle
   * @returns {Promise<boolean>} true if permission granted
   */
  async function verifyWritePermission(handle) {
    if (!handle) return false;
    try {
      const opts = { mode: 'readwrite' };
      if ((await handle.queryPermission(opts)) === 'granted') return true;
      if ((await handle.requestPermission(opts)) === 'granted') return true;
    } catch (_) {
      // Permission denied or API not available
    }
    return false;
  }

  /**
   * Write data to a FileSystemFileHandle.
   * @param {FileSystemFileHandle} handle
   * @param {Blob|ArrayBuffer|string} contents
   */
  async function writeToHandle(handle, contents) {
    const writable = await handle.createWritable();
    try {
      await writable.write(contents);
    } finally {
      await writable.close();
    }
  }

  /**
   * Store a directory handle for creating new files.
   * @param {FileSystemDirectoryHandle} dirHandle
   * @returns {Promise<string>} The handle ID
   */
  async function saveDirHandle(dirHandle) {
    const id = 'dir:' + generateId();
    await withStore('readwrite', (store) => {
      store.put(dirHandle, id);
    });
    return id;
  }

  return {
    saveHandle,
    getHandle,
    deleteHandle,
    verifyWritePermission,
    writeToHandle,
    saveDirHandle,
    generateId,
  };
})();
