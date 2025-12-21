# APX.AI Secure Transmission Outlook Add-in Development Roadmap (v1.0)

基於 PRD v1.0，總步驟 9，單檔/小步模式，每檔 <300 行。依賴 /shared 全重用，ESLint strict mode 全程。

1. [ ] **Manifest**  
   - 產 manifest.xml (v1.1 XML，Ribbon button、Taskpane 350px)。  
   - 驗證：Sideload Outlook Web，檢查 Ribbon button 顯示。  
   預估：manifest 200 行。依賴：無。

2. [ ] **Storage Adapter**  
   - 產 outlook/storage-adapter.js (indexedDB/roamingSettings 封裝 storage-core，暴露 apxStorage 方法，JSDoc 詳細)。  
   - 驗證：Mock Office.js，手動測試 save/load/remove 與 7 天 expiry。  
   預估：150 行。依賴：步1 (Office 環境)。

3. [ ] **Taskpane HTML & ESLint Setup**  
   - 產 taskpane.html (像素級 View：Server URL、Login、私鑰驗證、Main、Loading、Error，Bootstrap CDN、toggle 樣式)。  
   - 產 package.json + .eslintrc.json (ESLint strict mode，devDependencies eslint，參考 Gmail)。  
   - 驗證：開 taskpane.html 瀏覽器，檢查 View 布局一致。跑 npx eslint .  
   預估：taskpane 250 行，package/eslintrc 150 行。依賴：步1 (Taskpane 定義)。

4. [ ] **View Switcher**  
   - 產 outlook/view-switcher.js (Office.initialize 初始化、View 切換、storage 檢查導航、收件人讀取 to.getAsync，事件綁定)。  
   - 驗證：Sideload，測試登入流程導航無誤。  
   預估：200 行。依賴：步2-3 (storage + HTML)。

5. [ ] **Upload Handler**  
   - 產 outlook/upload-handler.js (登入/私鑰驗證、上傳 apiService.uploadFile、uploadStatus 更新，用 constants.getMessage)。  
   - 驗證：模擬上傳，檢查 status 文字與 Gmail 一致。  
   預估：180 行。依賴：步4 (View 切換)。

6. [ ] **Link Inserter**  
   - 產 outlook/link-inserter.js (上傳成功後 prependAsync 插入 baseUrl 連結格式、成功訊息後關閉 Taskpane)。  
   - 驗證：上傳後檢查郵件本文只 baseUrl 連結。  
   預估：120 行。依賴：步5 (上傳完成)。

7. [ ] **Ribbon & Attachment Handlers**  
   - 產 outlook/ribbon-handler.js (Ribbon button 手動開 Taskpane)。  
   - 產 outlook/attachment-handler.js (監聽 attachmentsChanged，偵測 >25MB 單檔、removeAttachmentAsync、自動開 Taskpane)。  
   - 驗證：點 Ribbon 開 Taskpane；加大檔自動觸發 + 移除原附件。  
   預估：各 150 行。依賴：步6 (Taskpane 完整)。

8. [ ] **Error Handler & Global Integration**  
   - 產 outlook/error-handler.js (全域錯誤處理、導回 View、清除過期 storage、Error View 用 constants.getMessage)。  
   - 小修多檔 (taskpane.js 等：Office Theme 適配、import shared 全覆蓋、無 magic)。  
   - 驗證：模擬錯誤，檢查導回 + log 正確。  
   預估：error 150 行，各檔 +30 行。依賴：步7 (所有 handler)。

9. [ ] **完整 e2e 手動測試驗證**  
   - Sideload Web/Desktop (PRD 第6節步驟)，測試全流程 (手動/自動、上傳/插入、7 天登入、錯誤分支、zhTW 文字、ESLint clean、無 console error)。  
   - 驗證：記 bug，重現步驟；若錯，用 follow-up prompt 修正。  
   預估：無 (測試文件)。依賴：步8 (全功能)。