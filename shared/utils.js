/**
 * APX.AI 工具函數。
 * Outlook 重用。
 */

(function() {
  /**
   * Sleep。
   * @param {number} ms
   * @returns {Promise<void>}
   */
  const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  /**
   * API 錯誤訊息解析。
   * @param {object} apiResult
   * @returns {string} 中文化錯誤。
   */
  const getApiErrorMessage = (apiResult) => {
    if (apiResult?.Errors && Array.isArray(apiResult.Errors) && apiResult.Errors.length > 0) {
      return apiResult.Errors.map(err => err.Message || err.message || JSON.stringify(err)).join('、');
    }
    return apiResult.Message || apiResult.message || '未知 API 錯誤。';
  };

  /**
   * PEM to Base64（storage 內嵌備份）。
   * @param {string} pemContent
   * @returns {string}
   */
  const getPrivateKeyBase64 = (pemContent) => btoa(pemContent);

  // Global 暴露
  window.utils = {
    sleep,
    getApiErrorMessage,
    getPrivateKeyBase64,
  };
})();
