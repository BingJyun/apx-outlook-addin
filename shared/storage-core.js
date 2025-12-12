/**
 * APX.AI 儲存核心邏輯（抽象層）。
 * 7天 expiry、saveCredentials/verifyPrivateKey。
 * Outlook/Gmail adapter import此，無平台字眼。
 * 用 window.constants/window.utils。
 */

(function() {
  /**
   * 檢查過期（純邏輯）。
   * @param {number} timestamp
   * @returns {boolean}
   */
  const isAuthExpired = (timestamp) => {
    if (!timestamp) return true;
    return (Date.now() - timestamp) > window.constants.DEFAULTS.SEVEN_DAYS_MS;
  };

  /**
   * 驗證私鑰 + 更新認證（純邏輯）。
   * @param {object} auth - 現有認證
   * @param {string} pemContent - PEM
   * @returns {object} 更新後認證
   * @throws {Error}
   */
  const verifyPrivateKeyLogic = (auth, pemContent) => {
    if (!auth?.account) throw new Error(window.constants.getMessage('NO_LOGIN_DATA', 'zh-TW'));
    const keyFileBase64 = window.utils.getPrivateKeyBase64(pemContent);
    return {
      ...auth,
      keyFileBase64,
      isAuthenticated: true,
    };
  };

  // Global 暴露（adapter呼叫）
  window.storageCore = {
    isAuthExpired,
    verifyPrivateKeyLogic,
    saveCredentialsData: (account, password) => ({
      account,
      password,
      keyFileBase64: null,
      isAuthenticated: false,
    }),
    addTimestamp: (data) => ({ ...data, timestamp: Date.now() }),
  };
})();
