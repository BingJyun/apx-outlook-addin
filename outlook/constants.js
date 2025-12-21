/**
 * APX.AI Outlook 專屬常數。
 * 封裝 Outlook 專屬的常數，避免 magic string/number。
 * 重用 /shared/constants.js 的共用部分，Outlook 專屬的放這裡。
 */

(function() {
  'use strict';

  /**
   * IndexedDB 資料庫名稱。
   * @constant {string}
   */
  const DB_NAME = 'ApxOutlookStorage';

  /**
   * IndexedDB 物件存放區名稱。
   * @constant {string}
   */
  const STORE_NAME = 'keyValueStore';

  /**
   * IndexedDB 版本號。
   * @constant {number}
   */
  const DB_VERSION = 1;

  // 暴露至全域（與 shared/constants.js 合併）
  window.constants = {
    ...window.constants, // 保留 shared 的常數
    DB_NAME,
    STORE_NAME,
    DB_VERSION,
  };
})();