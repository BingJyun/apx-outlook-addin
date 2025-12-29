/**
 * APX.AI Outlook Auth Handler。
 * 專責處理登入、登出、私鑰驗證與 storage 存取邏輯。
 * 驗證成功後呼叫 window.viewSwitcher 切換到 Main View。
 * 所有文字來自 constants.getMessage(key, 'zhTW')。
 */

(function() {
  'use strict';

  /**
   * 綁定 Server URL 繼續按鈕事件。
   * 儲存 serverUrl 並導向登入 View。
   */
  const bindContinueBtn = () => {
    const continueBtn = document.getElementById('continueBtn');
    if (!continueBtn) {return;}

    continueBtn.addEventListener('click', async () => {
      const serverUrlInput = document.getElementById('serverUrlInput');
      if (!serverUrlInput) {return;}

      const url = serverUrlInput.value.trim();
      if (!url) {
        window.errorHandler.showError('EMPTY_SERVER_URL');
        return;
      }

      try {
        await window.apxStorage.saveServerUrl(url);
        window.viewSwitcher.showView('loginView');
      } catch {
        window.errorHandler.handleAuthError('AUTH_EXPIRED');
      }
    });
  };

  /**
   * 綁定登入按鈕事件。
   * 儲存 credentials 並導向私鑰驗證 View。
   */
  const bindLoginBtn = () => {
    const loginBtn = document.getElementById('loginBtn');
    if (!loginBtn) {return;}

    loginBtn.addEventListener('click', async () => {
      const loginAcc = document.getElementById('loginAcc');
      const loginPwd = document.getElementById('loginPwd');
      if (!loginAcc || !loginPwd) {return;}

      const account = loginAcc.value.trim();
      const password = loginPwd.value;

      if (!account || !password) {
        window.errorHandler.showError('NO_LOGIN_DATA');
        return;
      }

      try {
        await window.apxStorage.saveCredentials(account, password);
        window.viewSwitcher.showView('privateKeyView');
      } catch {
        window.errorHandler.handleAuthError('AUTH_EXPIRED');
      }
    });
  };

  /**
   * 綁定私鑰驗證按鈕事件。
   * 驗證私鑰並導向 Main View。
   */
  const bindVerifyKeyBtn = () => {
    const verifyKeyBtn = document.getElementById('verifyKeyBtn');
    if (!verifyKeyBtn) {return;}

    verifyKeyBtn.addEventListener('click', async () => {
      const pemFileInput = document.getElementById('pemFileInput');
      const pemPwdInput = document.getElementById('pemPwdInput');
      if (!pemFileInput || !pemPwdInput) {return;}

      const file = pemFileInput.files[0];
      const pemPwd = pemPwdInput.value.trim();

      if (!file) {
        window.errorHandler.showError('NO_PRIVATE_KEY_FILE');
        return;
      }

      if (!pemPwd) {
        window.errorHandler.showError('EMPTY_PRIVATE_KEY_PASSWORD');
        return;
      }

      try {
        const pemContent = await file.text();
        await window.apxStorage.verifyPrivateKey(pemContent);
        window.viewSwitcher.showView('mainView');
      } catch {
        window.errorHandler.handleAuthError('AUTH_EXPIRED');
      }
    });
  };

  /**
   * 綁定登出按鈕事件。
   * 清除認證 storage 並導向登入 View。
   */
  const bindLogoutBtn = () => {
    const logoutBtn = document.getElementById('logoutBtn');
    if (!logoutBtn) {return;}

    logoutBtn.addEventListener('click', async () => {
      try {
        await window.apxStorage.remove();
        window.viewSwitcher.showView('loginView');
      } catch {
        window.errorHandler.handleAuthError('AUTH_EXPIRED');
      }
    });
  };

  /**
   * 綁定密碼顯示/隱藏 toggle 事件。
   * @param {string} inputId - 密碼輸入框 ID。
   * @param {string} toggleId - Toggle 按鈕 ID。
   * @param {string} iconId - Icon ID。
   */
  const bindPasswordToggle = (inputId, toggleId, iconId) => {
    const toggleBtn = document.getElementById(toggleId);
    const icon = document.getElementById(iconId);
    if (!toggleBtn || !icon) {return;}

    toggleBtn.addEventListener('click', () => {
      const input = document.getElementById(inputId);
      if (!input) {return;}

      const isVisible = input.type === 'text';
      input.type = isVisible ? 'password' : 'text';
      icon.className = isVisible ? window.constants.STYLES.PASSWORD_TOGGLE_HIDDEN : window.constants.STYLES.PASSWORD_TOGGLE_VISIBLE;
    });
  };

  /**
   * 初始化所有事件綁定。
   * 假設在 view-switcher 之後載入，DOM 已準備好。
   */
  const init = () => {
    bindContinueBtn();
    bindLoginBtn();
    bindVerifyKeyBtn();
    bindLogoutBtn();
    bindPasswordToggle('loginPwd', 'loginPwdToggle', 'loginPwdIcon');
    bindPasswordToggle('pemPwdInput', 'pemPwdToggle', 'pemPwdIcon');
  };

  // DOM 載入後初始化
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();