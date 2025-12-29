/**
 * APX.AI Outlook Upload Handler。
 * 專責處理檔案上傳邏輯與即時狀態更新。
 * 上傳僅呼叫 /shared/apiService.uploadFile。
 * 成功後通知 link-inserter 插入 baseUrl（不處理插入細節）。
 */

(function() {
  'use strict';

  /**
   * DOM 元素 ID 常數（避免 magic string）。
   * @enum {string}
   */
  const ELEMENTS = {
    FILE_INPUT: 'fileInput',
    UPLOAD_BTN: 'uploadBtn',
    UPLOAD_STATUS: 'uploadStatus',
  };

  /**
   * 處理上傳流程。
   * @returns {Promise<void>}
   */
  const handleUpload = async () => {
    const fileInput = document.getElementById(ELEMENTS.FILE_INPUT);
    const uploadStatus = document.getElementById(ELEMENTS.UPLOAD_STATUS);
    const file = fileInput.files[0];

    if (!file) {
      window.errorHandler.showError('NO_FILE_SELECTED');
      return;
    }

    // 載入必要資料
    const authData = await window.apxStorage.load();
    if (!authData || !authData.account || !authData.password || !authData.keyFileBase64) {
      window.errorHandler.showError('NO_LOGIN_DATA');
      return;
    }

    const serverUrlData = await window.apxStorage.loadServerUrl();
    if (!serverUrlData || !serverUrlData.url) {
      window.errorHandler.showError('NO_SERVER_URL');
      return;
    }

    const memberReceiveAcc = await window.viewSwitcher.getRecipient();
    if (!memberReceiveAcc) {
      window.errorHandler.showError('NO_RECIPIENT');
      return;
    }

    // 更新狀態：上傳中
    uploadStatus.textContent = window.constants.getMessage('UPLOADING', 'zhTW');

    // 呼叫 API 上傳
    await window.apiService.uploadFile({
      file: file,
      account: authData.account,
      password: authData.password,
      keyFileBase64: authData.keyFileBase64,
      memberReceiveAcc: memberReceiveAcc,
    });

    // 更新狀態：成功
    uploadStatus.textContent = window.constants.getMessage('SUCCESS', 'zhTW');

    // 通知 link-inserter 插入連結
    window.linkInserter.insertDownloadLink(file.name, serverUrlData.url);
  };

  /**
   * 設定事件監聽器。
   */
  const setEventListeners = () => {
    const uploadBtn = document.getElementById(ELEMENTS.UPLOAD_BTN);
    if (uploadBtn) {
      uploadBtn.addEventListener('click', async () => {
        try {
          await handleUpload();
        } catch {
          window.errorHandler.showError('UPLOAD_FAILED');
        }
      });
    }
  };

  /**
   * 初始化 Upload Handler。
   * 在 DOM 載入後設定監聽器。
   */
  document.addEventListener('DOMContentLoaded', () => {
    setEventListeners();
  });

  // 暴露 API（若未來需外部呼叫）
  window.uploadHandler = {
    handleUpload,
  };
})();