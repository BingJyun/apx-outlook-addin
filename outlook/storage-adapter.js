/**
 * APX.AI Outlook 儲存 Adapter。
 * 封裝 IndexedDB（優先）與 RoamingSettings（fallback），抽象儲存邏輯。
 * 重用 /shared/storage-core，暴露 window.apxStorage 與 Gmail adapter 行為一致。
 * 7 天 expiry，JSDoc 詳細，所有非同步用 async/await。
 * 依賴 Office.js（RoamingSettings）與原生 IndexedDB。
 */

(function() {
  'use strict';

  /**
   * IndexedDB 資料庫名稱。
   * @constant {string}
   */
  const DB_NAME = window.constants.DB_NAME;

  /**
   * IndexedDB 物件存放區名稱。
   * @constant {string}
   */
  const STORE_NAME = window.constants.STORE_NAME;

  /**
   * IndexedDB 版本號。
   * @constant {number}
   */
  const DB_VERSION = window.constants.DB_VERSION;

  /**
   * 取得 IndexedDB 連線。
   * @returns {Promise<IDBDatabase>} IndexedDB 資料庫實例。
   * @throws {Error} 若 IndexedDB 不支援或開啟失敗。
   */
  const getIndexedDB = async () => {
    if (!window.indexedDB) {
      throw new Error('IndexedDB not supported');
    }
    return new Promise((resolve, reject) => {
      const request = window.indexedDB.open(DB_NAME, DB_VERSION);
      request.onerror = () => reject(new Error('IndexedDB open failed'));
      request.onsuccess = () => resolve(request.result);
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
    });
  };

  /**
   * 取得 RoamingSettings。
   * @returns {Promise<Object>} RoamingSettings 設定物件。
   * @throws {Error} 若 Office 環境無效。
   */
  const getRoamingSettings = async () => {
    if (!Office || !Office.context || !Office.context.roamingSettings) {
      throw new Error('RoamingSettings not available');
    }
    return Office.context.roamingSettings;
  };

  /**
   * 儲存資料至 IndexedDB。
   * @param {string} key - 鍵。
   * @param {*} value - 值。
   * @returns {Promise<void>}
   */
  const saveToIndexedDB = async (key, value) => {
    const db = await getIndexedDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([STORE_NAME], 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.put(value, key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('IndexedDB save failed'));
    });
  };

  /**
   * 從 IndexedDB 載入資料。
   * @param {string} key - 鍵。
   * @returns {Promise<*>} 值或 null。
   */
  const loadFromIndexedDB = async (key) => {
    const db = await getIndexedDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([STORE_NAME], 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(new Error('IndexedDB load failed'));
    });
  };

  /**
   * 從 IndexedDB 移除資料。
   * @param {string} key - 鍵。
   * @returns {Promise<void>}
   */
  const removeFromIndexedDB = async (key) => {
    const db = await getIndexedDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([STORE_NAME], 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.delete(key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('IndexedDB remove failed'));
    });
  };

  /**
   * 儲存資料至 RoamingSettings。
   * @param {string} key - 鍵。
   * @param {*} value - 值。
   * @returns {Promise<void>}
   */
  const saveToRoamingSettings = async (key, value) => {
    const settings = await getRoamingSettings();
    settings.set(key, value);
    return new Promise((resolve, reject) => {
      settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error('RoamingSettings save failed'));
        }
      });
    });
  };

  /**
   * 從 RoamingSettings 載入資料。
   * @param {string} key - 鍵。
   * @returns {Promise<*>} 值或 null。
   */
  const loadFromRoamingSettings = async (key) => {
    const settings = await getRoamingSettings();
    return new Promise((resolve, reject) => {
      settings.get(key, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || null);
        } else {
          reject(new Error('RoamingSettings load failed'));
        }
      });
    });
  };

  /**
   * 從 RoamingSettings 移除資料。
   * @param {string} key - 鍵。
   * @returns {Promise<void>}
   */
  const removeFromRoamingSettings = async (key) => {
    const settings = await getRoamingSettings();
    settings.remove(key);
    return new Promise((resolve, reject) => {
      settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error('RoamingSettings remove failed'));
        }
      });
    });
  };

  /**
   * 儲存資料至儲存系統（優先 IndexedDB，fallback RoamingSettings）。
   * @param {string} key - 鍵。
   * @param {*} value - 值。
   * @returns {Promise<void>}
   */
  const save = async (key, value) => {
    try {
      await saveToIndexedDB(key, value);
    } catch {
      await saveToRoamingSettings(key, value);
    }
  };

  /**
   * 從儲存系統載入資料（優先 IndexedDB，fallback RoamingSettings）。
   * @param {string} key - 鍵。
   * @returns {Promise<*>} 值或 null。
   */
  const load = async (key) => {
    try {
      return await loadFromIndexedDB(key);
    } catch {
      return await loadFromRoamingSettings(key);
    }
  };

  /**
   * 從儲存系統移除資料（優先 IndexedDB，fallback RoamingSettings）。
   * @param {string} key - 鍵。
   * @returns {Promise<void>}
   */
  const remove = async (key) => {
    try {
      await removeFromIndexedDB(key);
    } catch {
      await removeFromRoamingSettings(key);
    }
  };

  /**
   * 儲存資料並加上時間戳記。
   * @param {string} key - 鍵。
   * @param {*} data - 資料。
   * @returns {Promise<void>}
   */
  const saveWithExpiry = async (key, data) => {
    const dataWithTimestamp = window.storageCore.addTimestamp(data);
    await save(key, dataWithTimestamp);
  };

  /**
   * 載入資料並檢查是否過期，若過期則移除並返回 null。
   * @param {string} key - 鍵。
   * @returns {Promise<*>} 資料或 null。
   */
  const loadWithExpiry = async (key) => {
    const data = await load(key);
    if (data && window.storageCore.isAuthExpired(data.timestamp)) {
      await remove(key);
      return null;
    }
    return data;
  };

  /**
   * 儲存認證資料。
   * @param {string} account - 帳號。
   * @param {string} password - 密碼。
   * @returns {Promise<void>}
   */
  const saveCredentials = async (account, password) => {
    const authData = window.storageCore.saveCredentialsData(account, password);
    await saveWithExpiry(window.constants.STORAGE_KEYS.AUTH, authData);
  };

  /**
   * 驗證私鑰並更新認證資料。
   * @param {string} pemContent - PEM 內容。
   * @returns {Promise<void>}
   */
  const verifyPrivateKey = async (pemContent) => {
    const authData = await loadWithExpiry(window.constants.STORAGE_KEYS.AUTH);
    const updatedAuth = window.storageCore.verifyPrivateKeyLogic(authData, pemContent);
    await saveWithExpiry(window.constants.STORAGE_KEYS.AUTH, updatedAuth);
  };

  /**
   * 載入認證資料。
   * @returns {Promise<Object|null>} 認證資料或 null。
   */
  const loadAuth = async () => {
    return await loadWithExpiry(window.constants.STORAGE_KEYS.AUTH);
  };

  /**
   * 移除認證資料。
   * @returns {Promise<void>}
   */
  const removeAuth = async () => {
    await remove(window.constants.STORAGE_KEYS.AUTH);
  };

  /**
   * 儲存伺服器 URL。
   * @param {string} url - 伺服器 URL。
   * @returns {Promise<void>}
   */
  const saveServerUrl = async (url) => {
    await saveWithExpiry(window.constants.STORAGE_KEYS.SERVER_URL, { url });
  };

  /**
   * 載入伺服器 URL。
   * @returns {Promise<Object|null>} 伺服器 URL 資料或 null。
   */
  const loadServerUrl = async () => {
    return await loadWithExpiry(window.constants.STORAGE_KEYS.SERVER_URL);
  };

  /**
   * 移除伺服器 URL。
   * @returns {Promise<void>}
   */
  const removeServerUrl = async () => {
    await remove(window.constants.STORAGE_KEYS.SERVER_URL);
  };

  // 暴露至全域
  window.apxStorage = {
    saveCredentials,
    verifyPrivateKey,
    load: loadAuth,
    remove: removeAuth,
    saveServerUrl,
    loadServerUrl,
    removeServerUrl,
  };
})();