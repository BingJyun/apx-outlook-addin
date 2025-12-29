/**
 * APX.AI Outlook Add-in Link Inserter
 * 負責將安全下載連結插入郵件本文並關閉 Taskpane。
 * 遵循單一職責原則，只處理連結插入與關閉邏輯。
 */

/**
 * 插入安全下載連結到郵件本文並關閉 Taskpane。
 * 此函數僅負責連結插入，不涉及上傳、驗證或儲存邏輯。
 *
 * @async
 * @param {string} fileName - 上傳檔案的名稱，用於連結格式。
 * @param {string} baseUrl - 伺服器 base URL，用於建構下載連結。
 * @throws {Error} 如果插入連結失敗。
 */
async function insertDownloadLink(fileName, baseUrl) {
  try {
    // 建構下載連結（僅使用 server baseUrl，無 taskId）
    const linkHtml = `<br><br>---<br>此檔案透過 APX.AI 安全傳送：<br><b>${fileName}</b> - <a href="${baseUrl}" target="_blank">點此到 APX.AI 下載</a>（建議使用 Chrome 瀏覽器開啟）`;

    // 使用 Office.js 插入連結到郵件本文
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.prependAsync(
        linkHtml,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(`插入連結失敗: ${asyncResult.error.message}`));
          }
        }
      );
    });

    // 顯示成功訊息（短暫）
    window.viewSwitcher.showSuccess('SUCCESS');

    // 延遲後關閉 Taskpane
    await window.utils.sleep(window.constants.DEFAULTS.SUCCESS_CLOSE_DELAY);
    Office.context.ui.closeContainer();

  } catch (error) {
    window.errorHandler.showError('UPLOAD_FAILED');
    throw error; // 重新拋出以便上層處理
  }
}

// 暴露全域函數（僅此一個公開 API）
window.linkInserter = {
  insertDownloadLink,
};