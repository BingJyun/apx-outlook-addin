/**
 * APX.AI Outlook Error Handler。
 * 單一職責：集中處理所有錯誤顯示與認證清除邏輯。
 * 所有錯誤透過 viewSwitcher.showError 導向 Error View。
 * 認證錯誤自動清除 window.apxStorage。
 * 錯誤文字來自 constants.getMessage(key, 'zhTW')。
 * 無任何 UI 主題或 Office.initialize 邏輯。
 * @module outlook/error-handler
 */
(function() {
  'use strict';

  /**
   * 顯示錯誤 View 並設定訊息。
   * 統一入口，所有錯誤由此呼叫。
   * 直接設定 [data-key="ERROR_MESSAGE"] 以支援 i18n。
   * @param {string} messageKey - 錯誤訊息鍵值（來自 constants.MESSAGES）。
   * @throws {Error} 若 viewSwitcher 未準備好。
   */
  const showError = (messageKey) => {
    if (window.viewSwitcher && typeof window.viewSwitcher.showError === 'function') {
      window.viewSwitcher.showError(messageKey);
    } else {
      // Fallback：直接設定元素（避免循環依賴）
      const errorElement = document.querySelector('[data-key="ERROR_MESSAGE"]');
      if (errorElement) {
        errorElement.textContent = window.constants.getMessage(messageKey, 'zhTW');
      }
      // 切換到 error view（假設有 window.viewSwitcher）
      if (window.viewSwitcher) {
        window.viewSwitcher.showView('errorView');
      }
    }
  };

  /**
   * 處理認證相關錯誤：清除 storage 並顯示錯誤。
   * 用於 AUTH_EXPIRED、DOWNLOAD_AUTH_FAILED 等。
   * @param {string} messageKey - 錯誤訊息鍵值。
   * @returns {Promise<void>}
   */
  const handleAuthError = async (messageKey) => {
    try {
      if (window.apxStorage && typeof window.apxStorage.remove === 'function') {
        await window.apxStorage.remove();
      }
    } catch {
      // 靜默處理：storage 清除失敗不影響錯誤顯示
    }
    showError(messageKey);
  };

  // 暴露公開 API
  window.errorHandler = {
    /**
     * @see showError
     */
    showError,
    /**
     * @see handleAuthError
     */
    handleAuthError,
  };
})();