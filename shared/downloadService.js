/**
 * APX.AI 下載服務。
 * 負責下載輪詢、檔案下載與清理。
 * Outlook 重用。與 apiService.initiateDownload 搭配使用。
 * 職責：從 taskId 開始完成下載流程（不含初始化）。
 */

(function() {
  const { DEFAULTS, API_ENDPOINTS } = window.constants;
  const { getApiErrorMessage } = window.utils;
  const { sleep } = window.utils; // 假設 utils 已抽 sleep；若無，後續補

  /**
   * 輪詢下載狀態直到完成。
   * @param {string} taskId - 伺服器任務 ID。
   * @param {function(string):void} onProgress - 進度回調。
   * @returns {Promise<{status: string, originalFileName?: string}>} 完成時的狀態資料。
   * @throws {Error} 輪詢超時、錯誤狀態或其他 API 失敗。
   */
  const pollForCompletion = async (taskId, onProgress) => {
    onProgress("伺服器準備檔案中，請稍候...");
    const pollUrl = `${(await window.apxStorage.loadServerUrl())?.url || DEFAULTS.API_BASE_URL}${API_ENDPOINTS.DOWNLOAD_STATUS.replace('{taskId}', taskId)}`;
    const MAX_POLLS = 60; // 60 次輪詢 * 2 秒 = 2 分鐘超時
    const POLL_INTERVAL = 2000; // 2 秒間隔

    for (let i = 0; i < MAX_POLLS; i++) {
      await sleep(POLL_INTERVAL);

      const pollResponse = await fetch(pollUrl);

      if (!pollResponse.ok) {
        const errorBody = await pollResponse.text();
        let errorMsg;
        try {
          const errorJson = JSON.parse(errorBody);
          errorMsg = getApiErrorMessage(errorJson) || `輪詢 HTTP 錯誤！狀態：${pollResponse.status}`;
        } catch (e) {
          errorMsg = `輪詢 HTTP 錯誤！狀態：${pollResponse.status}，訊息：${errorBody}`;
        }
        throw new Error(errorMsg);
      }

      const pollResult = await pollResponse.json();
      const pollStatus = pollResult.Data?.status;
      if (pollStatus === "completed") {
        const originalFileName = pollResult.Data?.originalFileName || pollResult.Data?.baseFileName || pollResult.Data?.fileName;
        return { status: pollStatus, originalFileName };
      } else if (pollStatus === "error") {
        // 立即在進度顯示錯誤訊息，讓用戶友好知道
        onProgress(window.constants.getMessage('DOWNLOAD_AUTH_FAILED', 'zhTW'));
        throw new Error(window.constants.getMessage('DOWNLOAD_AUTH_FAILED', 'zhTW'));
      }
      onProgress(`伺服器處理中... (狀態：${pollStatus || 'pending'}, ${i + 1}/${MAX_POLLS})`);
    }

    throw new Error("下載超時。伺服器準備檔案時間過長。");
  };

  /**
   * 下載已準備好的檔案。
   * @param {string} taskId - 伺服器任務 ID。
   * @param {string} expectedFileName - 預期的檔案名稱（從輪詢取得）。
   * @returns {Promise<{fileBlob: Blob, fileName: string}>} 檔案 Blob 與原始檔名。
   * @throws {Error} 下載 API 失敗。
   */
  const downloadFileByTaskId = async (taskId, expectedFileName) => {
    const baseUrl = (await window.apxStorage.loadServerUrl())?.url || DEFAULTS.API_BASE_URL;
    const downloadUrl = `${baseUrl}${API_ENDPOINTS.DOWNLOAD_FILE.replace('{taskId}', taskId)}`;

    const downloadResponse = await fetch(downloadUrl);

    if (!downloadResponse.ok) {
      const errorBody = await downloadResponse.text();
      let errorMsg;
      try {
        const errorJson = JSON.parse(errorBody);
        errorMsg = getApiErrorMessage(errorJson) || `檔案下載 HTTP 錯誤！狀態：${downloadResponse.status}`;
      } catch (e) {
        errorMsg = `檔案下載 HTTP 錯誤！狀態：${downloadResponse.status}，訊息：${errorBody}`;
      }
      throw new Error(errorMsg);
    }

    const fileBlob = await downloadResponse.blob();
    let fileName = expectedFileName || `apx-download-${taskId}`;

    const contentDisposition = downloadResponse.headers.get("Content-Disposition");
    const originalFileNameMatch = contentDisposition?.match(/originalFileName="([^"]+)"/);
    if (originalFileNameMatch && originalFileNameMatch[1]) {
      fileName = decodeURIComponent(originalFileNameMatch[1]);
    } else {
      // 使用預期的檔案名稱（從輪詢解析），若無則預設
      console.warn("無法從 Content-Disposition 提取原始檔名，使用預期名稱或預設值。");
    }

    return { fileBlob, fileName };
  };

  /**
   * 清理伺服器任務。
   * @param {string} taskId - 伺服器任務 ID。
   * @returns {Promise<void>}
   */
  const cleanupTask = async (taskId) => {
    const baseUrl = (await window.apxStorage.loadServerUrl())?.url || DEFAULTS.API_BASE_URL;
    const cleanupUrl = `${baseUrl}${API_ENDPOINTS.CLEANUP.replace('{taskId}', taskId)}`;

    try {
      const cleanupResponse = await fetch(cleanupUrl, {
        method: "DELETE",
      });
      if (!cleanupResponse.ok) {
        console.error(`任務 ${taskId} 清理失敗：HTTP 狀態 ${cleanupResponse.status}`);
      } else {
        console.log(`任務 ${taskId} 清理成功。`);
      }
    } catch (cleanupError) {
      console.error(`任務 ${taskId} 清理錯誤：`, cleanupError);
    }
  };

  /**
   * 完整下載流程：輪詢 + 下載 + 清理。
   * @param {string} taskId - 伺服器任務 ID。
   * @param {function(string):void} onProgress - 進度回調。
   * @returns {Promise<{fileBlob: Blob, fileName: string}>} 檔案 Blob 與原始檔名。
   * @throws {Error} 任何步驟失敗。
   */
  const completeDownload = async (taskId, onProgress) => {
    let serverTaskId = taskId;
    let fileName = `apx-download-${taskId}`;

    try {
      // 步驟 1: 輪詢直到完成
      const pollResult = await pollForCompletion(serverTaskId, onProgress);
      if (pollResult.originalFileName) {
        // 解析原始檔名（去除時間戳與序號，例如 "1765440575036-896932846-apx-api-document.pdf" → "apx-api-document.pdf"）
        const nameParts = pollResult.originalFileName.split('-');
        if (nameParts.length >= 3) {
          fileName = nameParts.slice(2).join('-');
        } else {
          fileName = pollResult.originalFileName;
        }
      }

      onProgress("檔案準備完成，正在下載...");

      // 步驟 2: 下載檔案
      const { fileBlob, detectedFileName } = await downloadFileByTaskId(serverTaskId, fileName);
      if (detectedFileName && detectedFileName !== `apx-download-${serverTaskId}`) {
        fileName = detectedFileName;
      }

      // 步驟 3: 清理任務（只在成功時）
      onProgress("正在清理伺服器上的任務...");
      await cleanupTask(serverTaskId);

      return { fileBlob, fileName };

    } catch (err) {
      // 失敗時不清理，直接拋出
      throw err;
    }
  };

  // 全域暴露（與 apiService 一致）
  window.downloadService = {
    pollForCompletion,
    downloadFileByTaskId,
    cleanupTask,
    completeDownload,
  };
})();