/**
 * APX.AI Outlook Error Handler。
 * 集中處理所有錯誤。
 * 所有錯誤導向 Error View（使用 viewSwitcher.showError）。
 * 認證錯誤必須清除 storage-core（呼叫 apxStorage.remove）。
 * 錯誤文字統一來自 constants.getMessage(key, 'zhTW')。
 * 嚴禁 console.log（僅允許 console.error 用於嚴重不可預期錯誤，MVP 階段可暫留但最終移除）。
 * Office Theme 適配（確保 Taskpane 隨 Outlook 主題切換 light/dark）。
 */

(function() {
  'use strict';

  /**
   * 顯示錯誤 View 並設定訊息。
   * 統一入口，避免重複邏輯。
   * @param {string} messageKey - 錯誤訊息鍵（來自 constants）。
   */
  const showError = (messageKey) => {
    window.viewSwitcher.showError(messageKey);
  };

  /**
   * 處理認證相關錯誤，清除 storage 並顯示錯誤。
   * 專門用於需要清除認證的錯誤（如 AUTH_EXPIRED）。
   * @param {string} messageKey - 錯誤訊息鍵。
   * @returns {Promise<void>} 非同步清除 storage。
   */
  const handleAuthError = async (messageKey) => {
    try {
      await window.apxStorage.remove();
    } catch {
      // Silent fail: storage removal failure during auth error handling
    }
    showError(messageKey);
  };

  /**
   * 套用 Office Theme 至 Taskpane。
   * 監聽主題變更並動態調整背景與文字顏色。
   */
  const applyTheme = () => {
    const theme = Office.context.officeTheme;
    if (theme) {
      document.body.style.backgroundColor = theme.bodyBackgroundColor;
      document.body.style.color = theme.bodyForegroundColor;
      // 延伸：可調整其他元素如按鈕、輸入框，但目前聚焦核心
    }
  };

  /**
   * 監聽 Office Theme 變更事件。
   * 確保 Taskpane 隨 Outlook 主題即時切換。
   */
  const listenForThemeChanges = () => {
    if (Office.context.officeTheme && typeof Office.context.officeTheme.addEventHandler === 'function') {
      Office.context.officeTheme.addEventHandler('changed', applyTheme);
    }
  };

  /**
   * 初始化 Error Handler。
   * 包在 Office.initialize 內執行，確保 Office.js 準備好。
   */
  Office.initialize = () => {
    applyTheme();
    listenForThemeChanges();
  };

  // 暴露統一 API
  window.errorHandler = {
    showError,
    handleAuthError,
  };
})();