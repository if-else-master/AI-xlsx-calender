# AI-xlsx-calender

請先申請 Gemini API <br>


## 📝 步驟 1：建立 Google Cloud 專案

### 1.1 前往 Google Cloud Console
1. 開啟瀏覽器，前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 使用您的 Google 帳號登入
3. 點擊 **API和服務**

### 1.2 建立新專案
1. 點擊頁面頂部的 **專案選擇器**（通常顯示當前專案名稱）
2. 在彈出視窗中，點擊 **新增專案**
3. 填寫專案資訊：
   - **專案名稱**：例如 `Excel-Calendar-Parser`
   - **組織**：選擇適當的組織（個人用戶通常為 `無組織`）
   - **位置**：選擇適當的位置
4. 點擊 **建立**

### 1.3 切換至新專案
1. 等待專案建立完成（通常需要幾秒鐘）
2. 確保頁面頂部顯示的是您剛建立的專案名稱
3. 如果不是，請使用專案選擇器切換至新專案

---

## 🔌 步驟 2：啟用 Google Calendar API

### 2.1 前往 API 管理頁面
1. 在 Google Cloud Console 中，點擊左側選單 ☰
2. 選擇 **APIs & Services** > **Library**

### 2.2 搜尋並啟用 Calendar API
1. 在搜尋欄中輸入 `Google Calendar API`
2. 點擊搜尋結果中的 **Google Calendar API**
3. 在 API 詳細頁面中，點擊 **啟用** 按鈕
4. 等待 API 啟用完成

> **✅ 確認步驟**：啟用成功後，您會看到 "API 已啟用" 的訊息

---

## 🛡️ 步驟 3：設定 OAuth 同意畫面

### 3.1 前往 OAuth 同意畫面設定
1. 在左側選單中選擇 **APIs & Services** > **OAuth consent screen**

### 3.2 選擇用戶類型
- **Internal**：僅供組織內部使用（需要 Google Workspace 帳號）
- **External**：可供任何 Google 帳號使用（推薦）

> **💡 建議**：個人用戶選擇 **External**

### 3.3 填寫應用程式資訊
填寫以下必填欄位：

#### 基本資訊
- **應用程式名稱**：`Excel Calendar AI Parser`
- **使用者支援電子郵件**：您的 Gmail 地址

#### 開發人員聯絡資訊
- **電子郵件地址**：您的 Gmail 地址

---

## 🔑 步驟 4：建立 OAuth 2.0 客戶端 ID

### 4.1 前往憑證頁面
1. 在左側選單中選擇 **APIs & Services** > **Credentials**

### 4.2 建立新憑證
1. 點擊頁面頂部的 **+ CREATE CREDENTIALS**
2. 選擇 **OAuth client ID**

### 4.3 設定客戶端類型
1. **應用程式類型**：選擇 **Desktop application**
2. **名稱**：輸入描述性名稱，例如 `Excel Calendar Parser Desktop`

### 4.4 建立客戶端
1. 點擊 **CREATE**
2. 會顯示一個對話框，包含您的客戶端 ID 和客戶端密鑰
3. 點擊 **OK**（稍後我們會下載完整的憑證檔案）

---

## 📥 步驟 5：下載憑證檔案

### 5.1 尋找您的客戶端
1. 在 **Credentials** 頁面的 **OAuth 2.0 Client IDs** 區段
2. 找到您剛建立的客戶端

### 5.2 下載 JSON 檔案
1. 點擊客戶端名稱右側的 **下載** 圖示 ⬇️
2. 將檔案下載到您的專案目錄
3. **重新命名**檔案為 `credentials.json`

### 5.3 確認檔案內容
憑證檔案應該包含以下結構：
```json
{
  "installed": {
    "client_id": "your-client-id.apps.googleusercontent.com",
    "project_id": "your-project-id",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_secret": "your-client-secret",
    "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob", "http://localhost"]
  }
}
```
---
## 聯繫方式
**Email**：[rayc57429@gmail.com]

*最後更新：2025年8月31日*

