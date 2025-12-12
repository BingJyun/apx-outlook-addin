/**
 * APX.AI API 服務。
 * 負責上傳/下載初始化。
 * Outlook 重用。
 */

(function() {
  const { DEFAULTS, API_ENDPOINTS } = window.constants;
  const { getApiErrorMessage } = window.utils;

  /**
   * 上傳檔案。
   * @param {object} uploadData - {file, account, password, keyFileBase64, memberReceiveAcc}
   * @returns {Promise<object>} API 結果
   * @throws {Error} 上傳失敗
   */
  const uploadFile = async (uploadData) => {
    const baseUrl = (await window.apxStorage.loadServerUrl())?.url || DEFAULTS.API_BASE_URL;
    const formData = new FormData();
    formData.append('baseUrl', baseUrl);
    formData.append('account', uploadData.account);
    formData.append('password', uploadData.password);
    formData.append('keyFilePath', uploadData.keyFileBase64);
    formData.append('memberReceiveAcc', uploadData.memberReceiveAcc);
    formData.append('enableEncryption', DEFAULTS.ENABLE_ENCRYPTION);
    formData.append('description', `Uploaded from APX Plugin: ${uploadData.file.name}`);
    formData.append('deleteAfterDays', DEFAULTS.DELETE_AFTER_DAYS);
    formData.append('file', uploadData.file);

    const response = await fetch(`${baseUrl}${API_ENDPOINTS.UPLOAD}`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(window.constants.getMessage('UPLOAD_FAILED', 'zh-TW')
        .replace('{status}', response.status)
        .replace('{error}', errorText));
    }

    const result = await response.json();
    if (!result.Success) {
      throw new Error(getApiErrorMessage(result));
    }

    return result;
  };

  /**
   * 初始化下載。
   * @param {object} downloadData - {account, password, keyFile, fileNo}
   * @returns {Promise<string>} taskId
   * @throws {Error} 初始化失敗
   */
  const initiateDownload = async (downloadData) => {
    const baseUrl = (await window.apxStorage.loadServerUrl())?.url || DEFAULTS.API_BASE_URL;
    const body = {
      baseUrl,
      account: downloadData.account,
      password: downloadData.password,
      fileNo: downloadData.fileNo,
      keyFilePath: downloadData.keyFile,
      downloadPath: DEFAULTS.DOWNLOAD_PATH,
    };

    const response = await fetch(`${baseUrl}${API_ENDPOINTS.DOWNLOAD}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(window.constants.getMessage('DOWNLOAD_INIT_FAILED', 'zh-TW')
        .replace('{status}', response.status)
        .replace('{error}', errorText));
    }

    const result = await response.json();
    if (!result.Success) {
      throw new Error(getApiErrorMessage(result));
    }

    const taskId = result.Data?.taskId;
    if (!taskId) {
      throw new Error(window.constants.getMessage('NO_TASK_ID', 'zh-TW'));
    }

    return taskId;
  };

  // Global 暴露
  window.apiService = {
    uploadFile,
    initiateDownload,
  };
})();
