/**
 * APX.AI 全域常數與中英 i18n。
 * 禁止任何 magic string/number。
 * 設計為 Outlook 重用。
 */

(function() {
  /**
   * API 端點常數（統一，未來 Outlook 可擴展）。
   */
  const API_ENDPOINTS = {
    UPLOAD: '/gmailapi/upload',
    DOWNLOAD: '/gmailapi/download',
    DOWNLOAD_STATUS: '/gmailapi/download-status/{taskId}',
    DOWNLOAD_FILE: '/gmailapi/download-file/{taskId}',
    CLEANUP: '/gmailapi/cleanup/{taskId}',
  };

  /**
   * 預設與逾時常數。
   */
  const DEFAULTS = {
    API_BASE_URL: 'https://apxpoc.ioneit.com',
    MAX_POLLS: 60,
    POLL_INTERVAL_MS: 2000,
    MAX_FILE_SIZE_BYTES: 25 * 1024 * 1024, // 25MB
    DELETE_AFTER_DAYS: '180',
    ENABLE_ENCRYPTION: 'false',
    SEVEN_DAYS_MS: 7 * 24 * 60 * 60 * 1000,
    DOWNLOAD_PATH: './temp/server-override',
    SUCCESS_CLOSE_DELAY: 1500,
    DIALOG_HEIGHT_PERCENT: 80, // Taskpane 對話框高度百分比
    DIALOG_WIDTH_PERCENT: 35,  // Taskpane 對話框寬度百分比
  };

  /**
   * Storage 鍵名。
   */
  const STORAGE_KEYS = {
    AUTH: 'apx_auth',
    SERVER_URL: 'apx_server_url',
  };

  /**
   * Gmail DOM selectors（移除 content.js magic strings）。
   */
  const GMAIL_SELECTORS = {
    TOOLBAR: "div[aria-label='工具列']",
    FILE_INPUT: 'input[type="file"][name="Filedata"]',
    COMPOSE_BODY: 'div[aria-label="郵件內文"]',
    RECIPIENT_FIELDS: 'div[name="to"], div[name="cc"], div[name="bcc"]',
    MESSAGE_BODY: "div[aria-label='Message Body']",
  };

  /**
   * i18n 中英（台灣用語），從 storage 載語言碼切換。
   * @param {string} key - 訊息鍵。
   * @param {string} language - 'zhTW' | 'en-US'。
   * @returns {string} 訊息。
   */
  const getMessage = (key, language = 'zhTW') => {
    const messages = {
      zhTW: {
        // 錯誤
        UPLOAD_FAILED: '上傳失敗：HTTP {status} - {error}',
        DOWNLOAD_INIT_FAILED: '下載初始化失敗：HTTP {status} - {error}',
        EMPTY_SERVER_URL: '請輸入伺服器 URL',
        AUTH_EXPIRED: '認證資訊已過期或無效，請重新登入。',
        NO_PRIVATE_KEY_FILE: '請選擇私鑰檔案',
        EMPTY_PRIVATE_KEY_PASSWORD: '請輸入私鑰密碼',
        NO_RECIPIENT: '無法讀取有效的收件人 Email，請確認已在收件人欄位填寫。',
        FILE_TOO_LARGE: '檔案 "{name}" ({size} MB) 已超過 25MB 上限，將透過 APX.AI 安全傳送。',
        TIMEOUT: '下載逾時，伺服器處理過久。',
        NO_TASK_ID: '伺服器未回傳任務ID。',
        NO_LOGIN_DATA: '無登入資料，請先登入',
        DOWNLOAD_AUTH_FAILED: '下載失敗：帳號、密碼或私鑰檔案有誤',
        TASKPANE_OPEN_FAILED: '無法開啟安全傳送視窗，請檢查網路或重新嘗試。',
        // UI
        PROCESSING: '處理中...',
        UPLOADING: '檔案上傳中...',
        DOWNLOADING: '檔案準備完成，正在下載...',
        CLEANUP: '正在清理伺服器任務...',
        SUCCESS: '上傳成功！',
        DOWNLOAD_SUCCESS: '檔案 "{name}" 已成功下載。頁面可以關閉。',
        SERVER_PROCESSING: '伺服器處理中... (狀態：{status}, {attempt}/{max})',
        GMAIL_BUTTON_TEXT: '🔐 用 APX.AI 傳送',
        DOWNLOAD_FILL_FIELDS: '請填寫帳號、密碼，選擇私鑰檔案並輸入私鑰密碼。',
        // UI 標籤
        SERVER_URL_TITLE: '設定伺服器 URL',
        SERVER_URL_DESCRIPTION: '請輸入 APX.AI 伺服器網址',
        SERVER_URL_PLACEHOLDER: '例如：https://apxpoc.ioneit.com',
        CONTINUE_BUTTON: '繼續',
        LOGIN_TITLE: '登入 APX.AI',
        LOGIN_ACCOUNT_PLACEHOLDER: '帳號',
        LOGIN_PASSWORD_PLACEHOLDER: '密碼',
        LOGIN_BUTTON: '登入',
        PRIVATE_KEY_TITLE: '驗證您的私鑰',
        PRIVATE_KEY_PASSWORD_PLACEHOLDER: '私鑰密碼',
        VERIFY_KEY_BUTTON: '完成驗證',
        MAIN_TITLE: '使用 APX.AI 安全傳送',
        RECIPIENT_LABEL: '收件人：',
        UPLOAD_BUTTON: '上傳並產生連結',
        LOGOUT_BUTTON: '登出',
        REFRESH_BUTTON: '重新整理',
        LOADING_TEXT: '載入中',
        PROCESSING_TEXT: '處理中...',
        ERROR_MESSAGE: '錯誤訊息',
      },
      enUS: {
        // Errors
        UPLOAD_FAILED: 'Upload failed: HTTP {status} - {error}',
        DOWNLOAD_INIT_FAILED: 'Download init failed: HTTP {status} - {error}',
        EMPTY_SERVER_URL: 'Please enter server URL',
        AUTH_EXPIRED: 'Auth expired or invalid, please log in again.',
        NO_PRIVATE_KEY_FILE: 'Please select private key file',
        EMPTY_PRIVATE_KEY_PASSWORD: 'Please enter private key password',
        NO_RECIPIENT: 'Cannot read valid recipient email. Please fill in the recipient field.',
        FILE_TOO_LARGE: 'File "{name}" ({size} MB) exceeds 25MB limit, use APX.AI secure send.',
        TIMEOUT: 'Download timeout, server too slow.',
        NO_TASK_ID: 'No task ID returned from server.',
        NO_LOGIN_DATA: 'No login data, please log in first.',
        DOWNLOAD_AUTH_FAILED: 'Download failed: account, password, or private key file incorrect',
        TASKPANE_OPEN_FAILED: 'Failed to open secure send window, please check network or try again.',
        // UI
        PROCESSING: 'Processing...',
        UPLOADING: 'Uploading file...',
        DOWNLOADING: 'File ready, downloading...',
        CLEANUP: 'Cleaning server task...',
        SUCCESS: 'Upload success!',
        DOWNLOAD_SUCCESS: 'File "{name}" downloaded. Page can be closed.',
        SERVER_PROCESSING: 'Server processing... (status: {status}, {attempt}/{max})',
        GMAIL_BUTTON_TEXT: '🔐 Send with APX.AI',
        DOWNLOAD_FILL_FIELDS: 'Please fill in account, password, select private key file, and enter private key password.',
        // UI labels
        SERVER_URL_TITLE: 'Set Server URL',
        SERVER_URL_DESCRIPTION: 'Please enter the APX.AI server URL',
        SERVER_URL_PLACEHOLDER: 'e.g., https://apxpoc.ioneit.com',
        CONTINUE_BUTTON: 'Continue',
        LOGIN_TITLE: 'Login to APX.AI',
        LOGIN_ACCOUNT_PLACEHOLDER: 'Account',
        LOGIN_PASSWORD_PLACEHOLDER: 'Password',
        LOGIN_BUTTON: 'Login',
        PRIVATE_KEY_TITLE: 'Verify Your Private Key',
        PRIVATE_KEY_PASSWORD_PLACEHOLDER: 'Private Key Password',
        VERIFY_KEY_BUTTON: 'Complete Verification',
        MAIN_TITLE: 'Use APX.AI Secure Send',
        RECIPIENT_LABEL: 'Recipients:',
        UPLOAD_BUTTON: 'Upload and Generate Link',
        LOGOUT_BUTTON: 'Logout',
        REFRESH_BUTTON: 'Refresh',
        LOADING_TEXT: 'Loading',
        PROCESSING_TEXT: 'Processing...',
        ERROR_MESSAGE: 'Error Message',
      },
    };
    const msg = messages[language]?.[key] || key;
    return msg;
  }

  /**
   * 樣式常數（避免 HTML magic）。
   */
  const STYLES = {
    PASSWORD_TOGGLE_HIDDEN: 'bi bi-eye-slash',
    PASSWORD_TOGGLE_VISIBLE: 'bi bi-eye',
  };

  // Global 暴露（Chrome 相容，Outlook import）
  window.constants = {
    API_ENDPOINTS,
    DEFAULTS,
    STORAGE_KEYS,
    GMAIL_SELECTORS,
    getMessage,
    STYLES,
  };
})();