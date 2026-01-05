/**
 * APX.AI Outlook Error Handler。
 * 單一職責：集中處理所有錯誤顯示與認證清除邏輯。
 * 認證錯誤自動清除 window.apxStorage。
 * 錯誤文字來自 constants.getMessage(key, 'zhTW')。
 * @module outlook/error-handler
 */
(function() {
  'use strict';

  /**
   * 顯示錯誤 View 並設定訊息。
   * 統一入口，所有錯誤由此呼叫。
   * 直接設定 [data-key="ERROR_MESSAGE"] 以支援 i18n，並檢查 viewSwitcher 可用性後切換 View。
   * @param {string} messageKey - 錯誤訊息鍵值（來自 constants.MESSAGES）。
   */
  const showError = (messageKey) => {
    // 設定錯誤訊息文字
    const errorElement = document.querySelector('[data-key="ERROR_MESSAGE"]');
    if (errorElement) {
      errorElement.textContent = window.constants.getMessage(messageKey, 'zhTW');
    }
    // 切換到錯誤 View（若 viewSwitcher 可用）
    if (window.viewSwitcher && typeof window.viewSwitcher.showView === 'function') {
      window.viewSwitcher.showView(window.constants.VIEWS.ERROR);
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