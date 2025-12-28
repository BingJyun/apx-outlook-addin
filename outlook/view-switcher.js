/**
 * APX.AI Outlook View Switcher。
 * 單一職責：集中管理 View 切換、初始導航、收件人讀取、Office 主題適配。
 * 所有 Office.js 呼叫包在 Office.initialize 內。
 * 公開 API：showView(viewName)、showError(messageKey)。
 * @module outlook/view-switcher
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
   * @private
   */
  const hideAllViews = () => {
    document.querySelectorAll('[data-view]').forEach((view) => {
      view.style.display = 'none';
    });
  };

  /**
   * 顯示指定 View。
   * @param {string} viewName - View 名稱（來自 VIEWS）。
   * @public
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
   * @param {string} messageKey - 錯誤訊息鍵值（來自 constants.MESSAGES）。
   * @public
   */
  const showError = (messageKey) => {
    const errorElement = document.getElementById(ELEMENTS.ERROR_MESSAGE);
    if (errorElement) {
      errorElement.textContent = window.constants.getMessage(messageKey, 'zhTW');
    }
    showView(VIEWS.ERROR);
  };

  /**
   * 讀取收件人資訊並顯示於 recipientDisplay。
   * 使用 Office.context.mailbox.item.to.getAsync，取第一位 email 的本地部分。
   * @returns {Promise<boolean>} 成功讀取為 true，否則 false。
   * @private
   */
  const loadRecipient = async () => {
    try {
      const item = Office.context.mailbox.item;
      /** @type {Office.Recipients} */
      const recipients = await new Promise((resolve, reject) => {
        item.to.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(new Error(`收件人讀取失敗：${result.error.message}`));
          }
        });
      });

      if (!recipients || recipients.length === 0) {
        window.errorHandler.showError('NO_RECIPIENT');
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
      window.errorHandler.showError('NO_RECIPIENT');
      return false;
    }
  };

  /**
   * 檢查 storage 狀態並導向對應 View。
   * 順序：serverUrl → auth → isAuthenticated → loadRecipient → mainView。
   * @returns {Promise<void>}
   * @private
   */
  const checkStorageAndNavigate = async () => {
    try {
      // 檢查 serverUrl
      const serverUrlData = await window.apxStorage.loadServerUrl();
      if (!serverUrlData?.url) {
        showView(VIEWS.SERVER_INPUT);
        return;
      }

      // 檢查 auth
      const authData = await window.apxStorage.load();
      if (!authData?.account || !authData?.password) {
        showView(VIEWS.LOGIN);
        return;
      }

      // 檢查 private key 驗證
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
      if (window.errorHandler?.handleAuthError) {
        window.errorHandler.handleAuthError('AUTH_EXPIRED');
      } else {
        showError('AUTH_EXPIRED');
      }
    }
  };

  /**
   * 套用 Office 主題至 Taskpane body。
   * 確保與 Outlook 主題同步（light/dark）。
   * @private
   */
  const applyTheme = () => {
    if (Office.context?.officeTheme) {
      const theme = Office.context.officeTheme;
      document.body.style.backgroundColor = theme.bodyBackgroundColor;
      document.body.style.color = theme.bodyForegroundColor;
    }
  };

  /**
   * 監聽 Office 主題變更事件。
   * 使用正確的 Office.js API：addHandlerAsync(Office.EventType.ThemeChanged)。
   * @private
   * @returns {Promise<void>}
   */
  const listenForThemeChanges = async () => {
    try {
      if (Office.context?.officeTheme?.addHandlerAsync) {
        await Office.context.officeTheme.addHandlerAsync(
          Office.EventType.ThemeChanged,
          applyTheme
        );
      }
    } catch {
      // 靜默處理：主題監聽失敗不影響核心功能
    }
  };

  /**
   * 初始化 View Switcher。
   * 包含：loading → storage 檢查 → View 導航 + 主題套用/監聽。
   */
  Office.initialize = async () => {
    // 初始 loading
    showView(VIEWS.LOADING);

    // 套用初始主題並監聽變更
    applyTheme();
    await listenForThemeChanges();

    // 檢查並導航
    await checkStorageAndNavigate();
  };

  // 暴露公開 API
  window.viewSwitcher = {
    showView,
    showError,
  };
})();