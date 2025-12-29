# APX.AI Secure Transmission Outlook Add-in Development Roadmap (v1.1)

基於 PRD v1.0，總步驟 9，單檔/小步模式，每檔 <300 行。依賴 /shared 全重用，ESLint strict mode 全程。

## 1. Manifest & Icons Setup
**目標**  
建立 Outlook Add-in 的核心 manifest 與圖示資產。
**產出檔案**  
- manifest.xml  
- icons/ 資料夾（root，已有 icon16/32/80.png，確認引用正確）  
**實作規範（強制）**  
- v1.1 XML 格式  
- 定義 Ribbon button「用 APX.AI 安全傳送」（label 用 constants.GMAIL_BUTTON_TEXT）  
- Taskpane 寬度 350px  
- Compose 模式觸發  
**驗證**  
- Sideload Outlook Web，檢查 Ribbon button 顯示與 icons 載入  
**依賴**  
- 無

## 2. Storage Adapter
**目標**  
建立 Outlook 專屬的 storage 抽象層。  
**產出檔案**  
- outlook/storage-adapter.js  
**實作規範（強制）**  
- 使用 indexedDB/roamingSettings 封裝 shared/storage-core  
- 暴露 window.apxStorage（save/load/remove auth/serverUrl，7 天 expiry）  
- 詳細 JSDoc  
**驗證**  
- Mock Office.js，手動測試 save/load/remove 與 expiry  
**依賴**  
- Step 1（Office 環境）

## 3. Taskpane HTML & ESLint Config Setup
**目標**  
建立 Outlook Taskpane 的唯一 HTML 入口，並確立全專案的 ESLint 嚴格規範（package.json 已存在）。  
**產出檔案**  
- outlook/taskpane.html  
- .eslintrc.json  
**實作規範（強制）**  
- taskpane.html 為單一 HTML，寬度 350px  
- 僅使用 Bootstrap CDN，不引入其他 UI framework  
- View 必須完整包含：  
  - Server URL Input View  
  - Login View（帳號 / 密碼 / 密碼眼）  
  - Private Key Verification View（.pem + 密碼眼）  
  - Main View（收件人顯示、上傳區）  
  - Loading View  
  - Error View  
- 所有 View 僅用 data-view attribute 控制顯示，不得 inline JS  
- ESLint 使用 eslint:recommended 為基礎 + strict 規則（禁止 var、magic number、unused）  
**驗證**  
- 直接用瀏覽器開啟 taskpane.html，檢查像素與 Gmail 版一致  
- 執行 `npx eslint .` 必須 0 error  
**依賴**  
- Step 1（Taskpane 定義）

## 4. View Switcher
**目標**  
集中管理所有 View 狀態切換與初始導航邏輯。  
**產出檔案**  
- outlook/view-switcher.js  
**實作規範（強制）**  
- 所有 Office.js 呼叫必須包在 `Office.initialize`  
- View 切換只能透過本檔案公開 API  
- 啟動流程：  
  1. 檢查 storage-core 是否有完整認證  
  2. 依結果導向對應 View  
- 收件人資訊使用 `Office.context.mailbox.item.to.getAsync`  
- 不得直接操作 storage（僅呼叫 shared/storage）  
**驗證**  
- Sideload Outlook Web  
- 全登入流程 View 導航正確、無閃爍  
**依賴**  
- Step 2（storage）  
- Step 3（HTML）

## 5. Auth Handler
**目標**  
專責處理登入、登出、私鑰驗證與 storage 存取邏輯。  
**產出檔案**  
- outlook/auth-handler.js  
**實作規範（強制）**  
- 負責 saveCredentials、verifyPrivateKey、logout 等（清除 storage）  
- 所有 storage 操作透過 window.apxStorage  
- 驗證成功後通知 view-switcher 切換到 Main View  
- 錯誤拋出交給 error-handler  
**驗證**  
- Sideload，測試完整登入 → Main View、登出 → Login View  
**依賴**  
- Step 4（View Switcher）

## 6. Upload Handler
**目標**  
負責檔案上傳與即時狀態更新。  
**產出檔案**  
- outlook/upload-handler.js  
**實作規範（強制）**  
- 上傳僅呼叫 /shared/apiService.uploadFile  
- 上傳狀態文字必須來自 constants.getMessage  
- 成功後通知 link-inserter 插入 baseUrl  
- 不得操作 View 切換或 storage  
**驗證**  
- 已登入狀態下選擇檔案上傳，檢查 status 文字與 Gmail 一致  
**依賴**  
- Step 5（Auth Handler）

## 7. Link Inserter
**目標**  
僅負責將下載連結插入郵件並關閉 Taskpane。  
**產出檔案**  
- outlook/link-inserter.js  
**實作規範（強制）**  
- 僅插入 server baseUrl（不含 taskId）  
- 使用 `prependAsync`  
- 成功後顯示短暫成功訊息，再關閉 Taskpane  
- 不得進行任何上傳或驗證邏輯  
**驗證**  
- 郵件本文只出現 baseUrl  
- 無多餘文字或 HTML  
**依賴**  
- Step 6（Upload 完成）

## 8. Ribbon & Attachment Handlers
**目標**  
處理手動與自動觸發 Taskpane。  
**產出檔案**  
- outlook/ribbon-handler.js  
- outlook/attachment-handler.js  
**實作規範（強制）**  
- Ribbon button 僅負責開啟 Taskpane  
- attachmentsChanged：  
  - 偵測單一附件 ≥ 25MB（數值來自 constants）  
  - 自動移除原附件  
  - 開啟 Taskpane  
- 不得包含任何 UI 邏輯  
**驗證**  
- 點擊 Ribbon 正常開啟  
- 加入大檔案自動觸發  
**依賴**  
- Step 7（Taskpane 完整）

## 9. Error Handler & Global Integration
**目標**  
集中處理所有錯誤與全域整合收尾。  
**產出檔案**  
- outlook/error-handler.js  
**實作規範（強制）**  
- 所有錯誤導向 Error View  
- 認證錯誤必須清除 storage-core  
- 錯誤文字統一來自 constants.getMessage  
- 不得 console.log（僅允許 console.error，MVP 後移除）  
**驗證**  
- 模擬失敗流程，導回正確 View  
- 無殘留狀態  
**依賴**  
- Step 8（所有 handler）

## 10. 中期一致性盤點與小重構
**目標**  
確保 Outlook 所有新檔（outlook/*.js）全域一致、零技術債，為 MVP 打最乾淨基礎。  
**產出檔案**  
- 小修 outlook/*.js（view-switcher、auth-handler、upload-handler、link-inserter、ribbon-handler、attachment-handler、error-handler）  
**實作規範（強制）**  
- 所有錯誤顯示統一走 errorHandler.showError  
- 所有 View 切換統一走 viewSwitcher 公開 API  
- 所有 storage 操作統一走 window.apxStorage  
- Office.initialize 只定義一次（放在 view-switcher.js）  
- 移除所有殘留 console.log/error（或統一到 errorHandler）  
- 檢查所有 import 路徑正確、無重複 import  
- 檢查所有 DOM id 引用與 taskpane.html 匹配  
- 所有常數來自 /shared/constants  
- 產出後所有 outlook/*.js 必須 npx eslint 0 error  
**驗證**  
- 全專案 npx eslint outlook/ 0 error  
- 手動檢查錯誤流程全走 errorHandler  
**依賴**  
- Step 9（error-handler 完成）

## 11. 完整 e2e 手動測試驗證
**目標**  
確認整體流程符合 PRD，且無工程債。  
**測試項目**  
- Outlook Web sideload（優先）
- 手動 / 自動觸發  
- 上傳 / baseUrl 插入
- 7 天登入有效性  
- 錯誤分支  
- zh-TW 文字  
- ESLint clean（outlook/ 資料夾）
- 無 console error  
**結果**  
- 記錄 bug 與重現步驟  
- 問題以 follow-up prompt 修正  
**依賴**  
- Step 10（一致性盤點完成）