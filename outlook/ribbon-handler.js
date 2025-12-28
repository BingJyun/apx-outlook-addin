/**
 * APX.AI Outlook Ribbon Handler。
 * 專責處理手動觸發 Taskpane。
 * Ribbon button 僅負責開啟 Taskpane。
 * 不得包含任何 UI 邏輯。
 */

(function() {
  'use strict';

  /**
   * 開啟 Taskpane。
   * 使用 Office.js displayDialogAsync。
   * @returns {Promise<void>}
   */
  const openTaskpane = async () => {
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
          reject(new Error('Taskpane 開啟失敗'));
        }
      });
    });
  };

  /**
   * Ribbon button 點擊事件處理。
   * @param {Office.RibbonControl} _control - Ribbon control（未使用）。
   * @returns {Promise<void>}
   */
  const onRibbonButtonClick = async (_control) => {
    try {
      await openTaskpane();
    } catch {
      window.errorHandler.showError('TASKPANE_OPEN_FAILED');
    }
  };

  // 暴露至全域（Office.js 調用）
  window.ribbonHandler = {
    onRibbonButtonClick,
  };
})();