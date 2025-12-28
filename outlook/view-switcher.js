/**
 * APX.AI Outlook View Switcher。
 * 集中管理所有 View 狀態切換與初始導航邏輯。
 * 啟動流程：檢查 storage-core 完整認證（serverUrl + auth + keyFileBase64 + 未過期），依結果導向對應 View。
 * 收件人使用 Office.context.mailbox.item.to.getAsync，取第一位顯示 memberReceiveAcc。
 * 公開 showView(viewName) API 供 handler 呼叫。
 */

(function() {
  'use strict';

  /**
   * View 名稱常數（避免 magic string）。
   * @enum {string}
   */
  const VIEWS = {
    SERVER_INPUT: 'serverInputView',
    LOGIN: 'loginView',
    PRIVATE_KEY: 'privateKeyView',
    MAIN: 'mainView',
    LOADING: 'loadingView',
    ERROR: 'errorView',
  };

  /**
   * DOM 元素 ID 常數（避免 magic string）。
   * @enum {string}
   */
  const ELEMENTS = {
    RECIPIENT_DISPLAY: 'recipientDisplay',
    ERROR_MESSAGE: 'errorMessage',
  };

  /**
   * 隱藏所有 View。
   */
  const hideAllViews = () => {
    document.querySelectorAll('[data-view]').forEach(view => {
      view.style.display = 'none';
    });
  };

  /**
   * 顯示指定 View。
   * @param {string} viewName - View 名稱。
   */
  const showView = (viewName) => {
    hideAllViews();
    const view = document.querySelector(`[data-view="${viewName}"]`);
    if (view) {
      view.style.display = 'block';
    }
  };

  /**
   * 顯示錯誤 View 並設定訊息。
   * @param {string} messageKey - 錯誤訊息鍵。
   */
  const showError = (messageKey) => {
    const errorElement = document.getElementById(ELEMENTS.ERROR_MESSAGE);
    if (errorElement) {
      errorElement.textContent = window.constants.getMessage(messageKey, 'zhTW');
    }
    showView(VIEWS.ERROR);
  };

  /**
   * 讀取收件人資訊並顯示。
   * @returns {Promise<boolean>} 是否成功讀取。
   */
  const loadRecipient = async () => {
    try {
      const item = Office.context.mailbox.item;
      const recipients = await new Promise((resolve, reject) => {
        item.to.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(new Error('Failed to get recipients'));
          }
        });
      });

      if (!recipients || recipients.length === 0) {
        showError('NO_RECIPIENT');
        return false;
      }

      const firstRecipient = recipients[0];
      const email = firstRecipient.emailAddress || firstRecipient;
      const memberReceiveAcc = email.split('@')[0];
      const displayElement = document.getElementById(ELEMENTS.RECIPIENT_DISPLAY);
      if (displayElement) {
        displayElement.textContent = memberReceiveAcc;
      }
      return true;
    } catch {
      showError('NO_RECIPIENT');
      return false;
    }
  };

  /**
   * 檢查儲存狀態並導向對應 View。
   * @returns {Promise<void>}
   */
  const checkStorageAndNavigate = async () => {
    try {
      // 檢查 serverUrl
      const serverUrlData = await window.apxStorage.loadServerUrl();
      if (!serverUrlData || !serverUrlData.url) {
        showView(VIEWS.SERVER_INPUT);
        return;
      }

      // 檢查 auth
      const authData = await window.apxStorage.load();
      if (!authData || !authData.account || !authData.password) {
        showView(VIEWS.LOGIN);
        return;
      }

      // 檢查 isAuthenticated
      if (!authData.isAuthenticated) {
        showView(VIEWS.PRIVATE_KEY);
        return;
      }

      // 載入收件人並顯示 mainView
      const recipientLoaded = await loadRecipient();
      if (recipientLoaded) {
        showView(VIEWS.MAIN);
      }
    } catch {
      showError('AUTH_EXPIRED');
    }
  };

  /**
   * 初始化 View Switcher。
   * 包在 Office.initialize 內執行。
   */
  Office.initialize = async () => {
    // 初始顯示 loading
    showView(VIEWS.LOADING);

    // 檢查並導航
    await checkStorageAndNavigate();
  };

  // 暴露 API
  window.viewSwitcher = {
    showView,
    showError,
  };
})();