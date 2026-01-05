/**
 * APX.AI Outlook View Switcher。
 * 單一職責：集中管理 View 切換、初始導航、收件人讀取、Office 主題適配、附件監聽。
 * 所有 Office.js 呼叫包在 Office.initialize 內。
 * 公開 API：showView(viewName)、showError(messageKey)、showSuccess(messageKey)、getRecipient()。
 * @module outlook/view-switcher
 */

(function() {
  'use strict';

  /**
   * 將 Office.js callback 風格轉換為 Promise。
   * @param {Function} officeCall - Office.js API 呼叫函數。
   * @returns {Promise} Promise 化的結果。
   * @private
   */
  const promisifyOfficeCall = (officeCall) => {
    return new Promise((resolve, reject) => {
      officeCall((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(`Office API 呼叫失敗：${result.error.message}`));
        }
      });
    });
  };

  /**
   * 初始化所有 [data-key] 元素的文字內容。
   * 使用 constants.getMessage(key, 'zhTW') 設定 textContent。
   * @returns {void}
   * @private
   */
  const initTexts = () => {
    const elements = document.querySelectorAll('[data-key]');
    elements.forEach((element) => {
      const key = element.getAttribute('data-key');
      if (key) {
        element.textContent = window.constants.getMessage(key, 'zhTW');
      }
    });
  };

  /**
   * 初始化所有 [data-placeholder-key] 元素的 placeholder。
   * 使用 constants.getMessage(key, 'zhTW') 設定 placeholder。
   * @returns {void}
   * @private
   */
  const initPlaceholders = () => {
    const elements = document.querySelectorAll('[data-placeholder-key]');
    elements.forEach((element) => {
      const key = element.getAttribute('data-placeholder-key');
      if (key) {
        element.placeholder = window.constants.getMessage(key, 'zhTW');
      }
    });
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
   * @param {string} viewName - View 名稱（來自 constants.VIEWS）。
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
   * 顯示成功 View 並設定訊息。
   * 直接使用 constants.getMessage(key, 'zhTW') 設定 textContent。
   * @param {string} messageKey - 成功訊息鍵值（來自 constants.MESSAGES）。
   * @public
   */
  const showSuccess = (messageKey) => {
    const successElement = document.querySelector('[data-key="SUCCESS_MESSAGE"]');
    if (successElement) {
      successElement.textContent = window.constants.getMessage(messageKey, 'zhTW');
    }
    showView(window.constants.VIEWS.SUCCESS);
  };

  /**
   * 讀取收件人資訊並顯示於 recipientDisplay。
   * 顯示完整 email。
   * 使用 Office.context.mailbox.item.to.getAsync，取第一位收件人。
   * @returns {Promise<string|null>} memberReceiveAcc（本地部分）或 null。
   * @throws 無收件人時顯示錯誤訊息。
   * @private
   */
  const loadRecipient = async () => {
    try {
      const item = Office.context.mailbox.item;
      /** @type {Office.Recipients} */
      const recipients = await promisifyOfficeCall((callback) => item.to.getAsync(callback));

      if (!recipients || recipients.length === 0) {
        return null;
      }

      const firstRecipient = recipients[0];
      const email = firstRecipient.emailAddress || firstRecipient;
      const memberReceiveAcc = email.split('@')[0];
      const displayElement = document.getElementById('recipientDisplay');
      if (displayElement) {
        displayElement.textContent = email; // 顯示完整 email
      }
      return memberReceiveAcc;
    } catch {
      window.errorHandler.showError('NO_RECIPIENT');
      return null;
    }
  };

  /**
   * 公開 API：取得收件人本地部分（後端需求）。
   * @returns {Promise<string|null>} memberReceiveAcc 或 null。
   * @public
   */
  const getRecipient = async () => {
    return await loadRecipient();
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
        showView(window.constants.VIEWS.SERVER_INPUT);
        return;
      }

      // 檢查 auth
      const authData = await window.apxStorage.load();
      if (!authData?.account || !authData?.password) {
        showView(window.constants.VIEWS.LOGIN);
        return;
      }

      // 檢查 private key 驗證
      if (!authData.isAuthenticated) {
        showView(window.constants.VIEWS.PRIVATE_KEY);
        return;
      }

      // 載入收件人並顯示 mainView
      const memberReceiveAcc = await loadRecipient();
      if (memberReceiveAcc) {
        showView(window.constants.VIEWS.MAIN);
      }
    } catch {
      window.errorHandler.handleAuthError('AUTH_EXPIRED');
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
        await promisifyOfficeCall((callback) => Office.context.officeTheme.addHandlerAsync(Office.EventType.ThemeChanged, applyTheme, callback));
      }
    } catch {
      // 靜默處理：主題監聽失敗不影響核心功能
    }
  };

  /**
   * 附件變更事件處理。
   * 檢查單一附件是否超過閾值，若是則移除並開啟 Taskpane。
   * @param {Office.AttachmentsChangedEventArgs} _event - 事件引數（未使用）。
   * @returns {Promise<void>}
   * @private
   */
  const onAttachmentsChanged = async (_event) => {
    const item = Office.context.mailbox.item;
    const attachments = await promisifyOfficeCall((callback) => item.attachments.getAsync(callback));

    if (attachments.length === 1) {
      const attachment = attachments[0];
      if (attachment.size > window.constants.DEFAULTS.MAX_FILE_SIZE_BYTES) {
        await promisifyOfficeCall((callback) => item.removeAttachmentAsync(attachment.id, callback));
        const taskpaneUrl = `${window.location.protocol}//${window.location.host}/taskpane.html`;
        await new Promise((resolve, reject) => {
          Office.context.ui.displayDialogAsync(taskpaneUrl, {
            height: window.constants.DEFAULTS.DIALOG_HEIGHT_PERCENT,
            width: window.constants.DEFAULTS.DIALOG_WIDTH_PERCENT,
            displayInIframe: true
          }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(asyncResult.error);
            }
          });
        });
      }
    }
  };

  /**
   * 初始化 View Switcher。
   * 包含：loading → 初始化文字/佔位符 → storage 檢查 → View 導航 + 主題套用/監聽 + 附件監聽。
   */
  Office.initialize = async () => {
    // 初始 loading
    showView(window.constants.VIEWS.LOADING);

    // 初始化文字和佔位符
    initTexts();
    initPlaceholders();

    // 套用初始主題並監聽變更
    applyTheme();
    await listenForThemeChanges();

    // 設定附件監聽
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, onAttachmentsChanged);

    // 檢查並導航
    await checkStorageAndNavigate();
  };

  // 暴露公開 API
  window.viewSwitcher = {
    showView,
    getRecipient,
    showSuccess,
  };
})();