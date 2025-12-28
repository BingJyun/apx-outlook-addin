# APX.AI Secure Transmission Outlook Add-in PRD

**文件版本**：v1.0（2025-12-20）  
**作者**：專案建築師（與 Grok 4.1 共同產出）  
**目標**：打造與 Gmail 版體驗零差異（甚至更好）、程式碼極度乾淨、完全符合企業資安規範的 Outlook Web Add-in。

## 1. 專案目標與核心原則（不可違背）

- **使用者體驗目標**：  
  使用者必須「完全感覺不出與 Gmail 版有任何差別，甚至更好」。所有視覺、文字、流程、錯誤訊息、按鈕行為必須像素級一致。

- **技術原則**：
  1. 客戶端絕對不做任何加密，所有加密交給 server。
  2. API 只使用兩個固定端點：`/gmailapi/upload` 與 `/gmailapi/download`（與 Gmail 版完全共用後端）。
  3. 所有共用邏輯必須 import 自 `/shared` 資料夾（apiService、downloadService、storage-core、utils、constants 等），禁止任何重複程式碼。
  4. 程式碼必須符合 Clean Code 規範（檔案 ≤ 300 行、見名知義、JSDoc、async/await、無 magic string）。
  5. Manifest 必須使用最新的 Office Add-ins manifest v1.1（XML 格式）。
  6. 所有 Office.js API 呼叫必須包在 `Office.initialize = () => {}` 內執行。

- **功能目標**：
  - 手動觸發：在 Compose 模式下提供 Ribbon button「用 APX.AI 安全傳送」。
  - 自動觸發：當使用者新增 ≥25MB 附件時，自動開啟 APX Taskpane 接管上傳。
  - 上傳完成後自動將安全下載連結插入郵件本文（格式與 Gmail 版完全相同）。
  - 7 天內免再次登入（使用 /shared/storage-core 封裝的帶過期機制）。

## 2. 使用者流程圖（文字版 Flowchart）

開始
  │
  ├─► 手動觸發：點擊 Ribbon button「用 APX.AI 安全傳送」
  │
  └─► 自動觸發：新增附件 → 偵測到單一附件 size ≥ 25MB（constants.DEFAULTS.MAX_FILE_SIZE_BYTES）→ 移除原附件避免雙重上傳

          ↓
    開啟 Taskpane（taskpane.html，寬度 350px）

          ↓
    initialize() → 檢查 storage-core

          ├─► serverUrl 存在 → 檢查 auth → isAuthenticated 且未過期 → Main View
          │
          └─► 依序切換 View（與 Gmail popup 完全相同順序）：
                  1. 無 serverUrl → Server URL Input View
                  2. 無有效帳密 → Login View
                  3. 無 keyFileBase64 或未驗證 → Private Key Verification View

          ↓ (Main View)
    讀取收件人 → 顯示於畫面 → 使用者選擇檔案 → 點擊「上傳並產生連結」

          ↓
    上傳 → 即時更新 uploadStatus 文字（與 Gmail uploadStatus.innerText 完全相同訊息）
          ↓
    apiService.uploadFile() 成功 → 取 server baseUrl 組成下載連結
          ↓
    Office.js item.body.prependAsync() 插入相同格式文字與連結
          ↓
    Taskpane 顯示短暫成功訊息 → 自動關閉（或提供關閉按鈕）

錯誤分支：
    → 任何錯誤 → 顯示紅色 Error View（訊息直接使用 constants.getMessage）
    → 認證相關錯誤 → 清除對應 storage → 導回對應 View

## 3. UI/UX 像素級對照表（根據 Gmail popup.html 精準還原）

### 通用要求（必須 100% 一致）
- Taskpane 寬度：350px（與 Gmail popup 一致）。
- Bootstrap 5.3.2 + Bootstrap Icons 1.11.3（直接 CDN 引用）。
- Logo：`<img src="./icons/icon80.png" style="height:100px;">` 置中。
- 主要按鈕：`class="btn btn-primary w-100"`。
- 密碼欄位：使用相同 input-group + password-toggle-btn 樣式（眼圖示右側，無邊框重疊）。
- 所有文字直接使用 `window.constants.getMessage(key, 'zhTW')`（固定 zhTW，不自動切換）。
- 支援 Office Theme 自動適配（背景、文字顏色隨 Office 切換）。
- icons 資料夾置於 root。

### Server URL Input View（對應 Gmail #serverInputView）
- Logo 100px 置中
- `<h5 class="mb-3 text-center">設定伺服器 URL</h5>`
- `<p class="small text-muted mb-3">請輸入 APX.AI 伺服器網址</p>`
- `<input type="url" id="serverUrlInput" class="form-control" placeholder="例如：https://apxpoc.ioneit.com">`
- `<button id="continueBtn" class="btn btn-primary w-100">繼續</button>`

### Login View（對應 Gmail #loginView）
- Logo 100px 置中
- `<h5 class="mb-3 text-center">登入 APX.AI</h5>`
- 帳號輸入：`<input type="text" id="loginAcc" class="form-control" placeholder="帳號">`
- 密碼輸入：input-group + eye toggle（id: loginPwd, loginPwdToggle, loginPwdIcon）
- `<button id="loginBtn" class="btn btn-primary w-100">登入</button>`

### Private Key Verification View（對應 Gmail #privateKeyView）
- Logo 100px 置中
- `<h5 class="mb-3 text-center">驗證您的私鑰</h5>`
- `.pem` 上傳：`<input type="file" id="pemFileInput" class="form-control" accept=".pem">`
- 私鑰密碼：input-group + eye toggle（id: pemPwdInput, pemPwdToggle, pemPwdIcon）
- `<button id="verifyKeyBtn" class="btn btn-primary w-100">完成驗證</button>`

### Main View（對應 Gmail #mainView）
- Logo 100px + 右上角兩個小按鈕：
  - `<button id="logoutBtn" class="btn btn-sm btn-outline-secondary position-absolute top-0 end-0">登出</button>`
  - `<button id="refreshBtn" class="btn btn-sm btn-outline-secondary position-absolute end-0" style="top: 38px;">重新整理</button>`
- `<h5 class="text-center m-0">使用 APX.AI 安全傳送</h5>`
- `<p class="small text-muted">收件人：<span id="recipientDisplay">自動讀取中...</span></p>`
- `<input type="file" id="fileInput" class="form-control">`
- （未來可顯示的閱後即毀選項目前隱藏）
- `<button id="uploadBtn" class="btn btn-primary w-100 mb-2">上傳並產生連結</button>`
- `<div id="uploadStatus" class="small text-break"></div>`

### Loading View（對應 Gmail #loadingView）
- 全畫面 spinner + `<p class="mt-2">處理中...</p>`

### Error View（對應 Gmail #errorView）
- `<div class="alert alert-danger">錯誤訊息</div>`

## 4. 功能規格對照表

- **收件人讀取**：
  - 使用 `Office.context.mailbox.item.to.getAsync()` 或 `item.getRecipientsAsync()` → 取第一位收件人 email → `memberReceiveAcc = email.split('@')[0]`（不支援多收件人）。

- **附件偵測與自動觸發**：
  - 在 `Office.initialize` 內監聽附件變化事件（如 `item.addHandlerAsync('attachmentsChanged')`）。
  - 若單一附件 size >= `constants.DEFAULTS.MAX_FILE_SIZE_BYTES`（25MB）→ 移除原附件（`item.removeAttachmentAsync()`）→ 自動開啟 Taskpane（只支援單檔）。
  - Outlook 預設大檔限：20-33MB（視 Exchange 伺服器），APX 接管避免超限錯誤。

- **連結插入格式**（簡化版，與 Gmail 版精神一致）：
  <br><br>---<br>此檔案透過 APX.AI 安全傳送：<br><b>$$ {fileName}</b> - <a href=" $${baseUrl}" target="_blank">點此到 APX.AI 下載</a>（建議使用 Chrome 瀏覽器開啟）
  - 使用 `item.body.prependAsync()` 插入。
  - 註：只用 server baseUrl（如 https://apxpoc.ioneit.com），使用者點開後自行 web 登入下載（無 taskId，簡化邏輯，避免下載頁 UI 依賴）。

- **儲存機制**：
  - 100% 使用 /shared/storage-core + Outlook 專屬 adapter（類似 gmail/storage-adapter.js）。
  - 優先 indexedDB（支援大資料與本地持久），fallback 到 roamingSettings（上限 32KB 足夠 Base64）。
  - Adapter 必須完全抽象，呼叫方式與 gmail/storage-adapter.js 一致（saveWithExpiry/loadWithExpiry/remove），暴露 window.apxStorage 物件：
    - saveCredentials(account, password): async, 用 saveWithExpiry 儲存 auth (account, password, timestamp)。
    - verifyPrivateKey(pemContent): async, 載入 auth，用 storage-core.verifyPrivateKeyLogic 更新 keyFileBase64 & isAuthenticated，saveWithExpiry。
    - load(): async, loadWithExpiry 取 auth，若過期移除返回 null。
    - remove(): async, 移除 auth key。
    - saveServerUrl(url): async, saveWithExpiry 儲存 {url}。
    - loadServerUrl(): async, loadWithExpiry 取 url。
    - removeServerUrl(): async, 移除 url key。
  - 內部：saveWithExpiry(key, data): async addTimestamp & indexedDB/roamingSettings set。
    loadWithExpiry(key): async get, check isAuthExpired, 若過期 remove 返回 null。
    remove(key): async indexedDB/roamingSettings remove。
  - 確保與 Gmail adapter 行為一致（7 天 expiry，JSDoc 詳細）。

- **文字訊息**：
  - 所有 UI 文字、錯誤訊息、狀態訊息直接使用 `constants.getMessage(key, 'zhTW')`（已包含中英，固定 zhTW）。

- **下載頁整合**：
  - 不支援 Add-in 注入 UI，使用者點連結後跳轉 server web 頁，自行登入下載（與 Gmail 非 extension 使用者一致）。

- **錯誤處理表格**（根據 constants.js zhTW 所有 key）：

| 錯誤類型         | 訊息 key                  | 處理方式                                      |
|------------------|---------------------------|-----------------------------------------------|
| 上傳失敗         | UPLOAD_FAILED            | 顯示 Error View，停留在 Main View，重試上傳。 |
| 下載初始化失敗   | DOWNLOAD_INIT_FAILED     | 顯示 Error View，若無下載 UI 僅 log。         |
| 認證過期         | AUTH_EXPIRED             | 清除 storage-core，導回 Login View。          |
| 無收件人         | NO_RECIPIENT             | 顯示 Error View，導回 Main View 重新讀取。     |
| 檔案太大         | FILE_TOO_LARGE           | 自動觸發 Taskpane（狀態訊息）。               |
| 下載逾時         | TIMEOUT                  | 顯示 Error View，若無下載 UI 僅 log。         |
| 無任務 ID        | NO_TASK_ID               | 顯示 Error View，重試 API（若需 taskId）。    |
| 無登入資料       | NO_LOGIN_DATA            | 導回 Login View。                             |
| 下載認證失敗     | DOWNLOAD_AUTH_FAILED     | 顯示 Error View，導回 Private Key View，若無下載 UI 僅 log。 |
| 處理中           | PROCESSING               | 狀態訊息，非錯誤。                            |
| 上傳中           | UPLOADING                | 狀態訊息，非錯誤。                            |
| 下載中           | DOWNLOADING              | 狀態訊息，非錯誤，若無下載 UI 僅 log。       |
| 清理中           | CLEANUP                  | 狀態訊息，非錯誤，若無下載 UI 僅 log。       |
| 上傳成功         | SUCCESS                  | 短暫顯示，關閉 Taskpane。                     |
| 下載成功         | DOWNLOAD_SUCCESS         | 狀態訊息，若無下載 UI 僅 log。                |
| 伺服器處理中     | SERVER_PROCESSING        | 狀態訊息，非錯誤，若無下載 UI 僅 log。       |
| 按鈕文字         | GMAIL_BUTTON_TEXT        | Ribbon button label（Outlook 用相同）。        |
| 下載填寫欄位     | DOWNLOAD_FILL_FIELDS     | 狀態訊息，若無下載 UI 僅 log。                |

## 5. Office 平台相容

- 優先 Outlook Online 相容，desktop 作為 bonus（測試 sideload）。
- 不支援 mobile Outlook。
- Ribbon icon：必備 16x16、32x32、80x80 px（PNG 格式）（參考 Microsoft docs：https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-icons）。

## 6. 測試與部署規格

- **部署步驟**：
  1. 用 Yeoman generator 建 Add-in 專案（yo office）。
  2. 更新 manifest.xml 與 code。
  3. Sideload 到 Outlook Web（Developer > Sideload）或 desktop（File > Manage Add-ins）。
  4. 未來上 Microsoft Partner Center 或企業 sideload。

- **測試方法**：
  - **開發人員手動測試**：用 VS Code Office Add-in debugger 跑 e2e（模擬 Compose、加 >25MB 附件、upload，檢查連結插入無延遲）。記錄 console error，驗證 async/await 無 callback。跑 ESLint strict mode 確保 clean。
  - **非程式人員測試**：提供 zip 檔 + 步驟文件（1. 開 Outlook Web，sideload manifest.xml；2. Compose 新郵件，點 Ribbon button；3. 加 >25MB 附件，檢查自動 Taskpane + 原附件移除；4. 上傳後驗證郵件本文連結；5. 截圖錯誤 + 重現步驟回報）。

## 7. 下一步計劃

1. 用 Continue.dev plan 功能讀 PRD & /shared，產生 roadmap。
2. 產 manifest.xml 完整版（單檔）。
3. 產 outlook/storage-adapter.js（單檔，參考 gmail/storage-adapter.js）。
4. 產 taskpane.html + taskpane.js（分檔，實作 View 切換）。
5. 逐一實作手動/自動觸發與連結插入。

（本文件已完善，可直接驅動開發）