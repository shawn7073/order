# 揪團點餐機器人：設定說明

整套免費服務組成，大約 40 分鐘可以設定完成。

| 元件 | 用途 | 費用 |
|---|---|---|
| LINE 官方帳號（Messaging API） | 機器人本體，群組開團、私訊建檔菜單 | 免費方案即可 |
| LINE Login channel + LIFF | 在 LINE 裡開啟點餐頁，自動帶入每個人的名字 | 免費 |
| Google 試算表 + Apps Script | 後端與資料庫（店家、訂單、付款狀態） | 免費 |
| GitHub Pages | 放點餐頁 `index.html` | 免費 |
| Gemini API | 辨識菜單照片 | 有免費額度 |

檔案：
- `Code.gs`：貼到 Apps Script 的後端程式
- `index.html`：點餐頁，填好兩個設定值後放到 GitHub Pages

---

## 步驟 1：Google 試算表與 Apps Script（後端）

1. 到 Google 雲端硬碟新增一份試算表，取名「揪團點餐」。
2. 選單「擴充功能 → Apps Script」，把預設內容全部刪掉，貼上 `Code.gs` 全部內容，存檔。
3. 上方函式下拉選 `setup`，按「執行」，依畫面授權（會出現「Google 尚未驗證」，點「進階 → 前往」即可，這是你自己的程式）。
   - 執行完試算表會多出 stores、orders、entries、drafts 四個分頁。
   - 左側「專案設定 → 指令碼屬性」會出現 `WEBHOOK_KEY`，等等會用到。
4. 右上「部署 → 新增部署作業」→ 類型選「網頁應用程式」：
   - 執行身分：**我**
   - 誰可以存取：**所有人**
5. 部署後複製「網頁應用程式網址」（結尾是 `/exec`），以下稱 **API 網址**。

> 之後若修改 `Code.gs`，要到「部署 → 管理部署作業 → 編輯 → 版本選『新版本』→ 部署」才會生效，網址不變。

## 步驟 2：申請 LINE 機器人（Messaging API）

1. 到 [LINE Official Account Manager](https://manager.line.biz/) 用你的 LINE 帳號登入，建立一個官方帳號（名稱例如「家裡點餐小幫手」）。
2. 進入帳號後，右上「設定」：
   - 「帳號設定」→ 功能切換 →「加入群組或多人聊天室」改為**接受邀請**。
   - 「回應設定」→ 關閉「自動回應訊息」與「加入好友的歡迎訊息」，開啟 **Webhook**。
   - 「Messaging API」→ 按「啟用 Messaging API」，建立一個 Provider（名稱隨意）。
3. 到 [LINE Developers Console](https://developers.line.biz/console/)，點剛建立的 Provider → 你的官方帳號 channel → 「Messaging API」分頁：
   - 最下方「Channel access token」按 Issue，複製起來。
   - 「Webhook URL」填：`API 網址?k=WEBHOOK_KEY`（例如 `https://script.google.com/macros/s/xxxx/exec?k=a1b2c3d4e5`），按 Update，再按 Verify 應顯示成功。
   - 開啟「Use webhook」。

## 步驟 3：點餐頁網址（GitHub Pages）

1. 註冊或登入 [GitHub](https://github.com/)，新增一個 Public repository，例如 `order`。
2. 先上傳 `index.html`（還沒填設定值也沒關係）。
3. repository 的「Settings → Pages」→ Source 選 `main` 分支、`/ (root)`，存檔。
4. 約一分鐘後會得到網址，例如 `https://你的帳號.github.io/order/`，以下稱 **點餐頁網址**。

## 步驟 4：LIFF（讓點餐頁在 LINE 裡開）

1. 回到 LINE Developers Console，同一個 Provider 下「Create a new channel」→ 選 **LINE Login**，地區選台灣，App type 勾 Web app。
2. 建好後在「Basic settings」複製 **Channel ID**，以下稱 **Login Channel ID**。
3. 「LIFF」分頁 → Add：
   - Size：Full
   - Endpoint URL：點餐頁網址
   - Scopes：勾 `openid`、`profile`
   - Bot link feature：On (Normal)，選剛剛的官方帳號
4. 建好後複製 **LIFF ID**（格式像 `2001234567-AbCdEfGh`）。
5. 重要：channel 狀態預設是「Developing」，只有你能用。到 channel 頁面上方按 **Publish**，你太太和朋友才能開點餐頁。

## 步驟 5：Gemini API key（辨識菜單）

1. 到 [Google AI Studio](https://aistudio.google.com/) 登入 → 「Get API key」→ 建立 API key，複製起來。
2. 若之後想改用 Claude 辨識，可在指令碼屬性加 `AI_PROVIDER=claude` 與 `ANTHROPIC_API_KEY`。

## 步驟 6：把設定值填進去

**Apps Script「專案設定 → 指令碼屬性」新增：**

| 屬性 | 值 |
|---|---|
| `CHANNEL_ACCESS_TOKEN` | 步驟 2 的 Channel access token |
| `LIFF_ID` | 步驟 4 的 LIFF ID |
| `LOGIN_CHANNEL_ID` | 步驟 4 的 Login Channel ID |
| `GEMINI_API_KEY` | 步驟 5 的 API key |
| `PAY_INFO`（選填） | 付款說明，例如「LINE Pay 或現金給開團的人」 |

（`WEBHOOK_KEY` 已自動產生，`ADMIN_USER_IDS` 會由機器人自動維護，都不用動。）

**`index.html` 最上方 CONFIG 改成：**

```js
const CONFIG = {
  LIFF_ID: '步驟 4 的 LIFF ID',
  API_URL: '步驟 1 的 API 網址'
};
```

改好重新上傳到 GitHub，覆蓋原檔。

## 步驟 7：開始使用

**設定管理員（建檔菜單的人）**
1. 你先加機器人好友（官方帳號頁面有 QR code），私訊「設定管理員」，你就成為第一位管理員。
2. 私訊「邀請管理員」，機器人給你 6 位數邀請碼（30 分鐘有效）。
3. 你太太加機器人好友後，私訊「管理員 123456」，她也成為管理員。

**建檔菜單（管理員私訊機器人）**
- 直接傳菜單照片 → 機器人回傳辨識結果，包含店名、電話、地址、營業時間、其他聯絡方式與所有品項價格。
- 菜單有好幾頁就連續傳，會自動合併成同一家。
- 有錯直接打字改，例如「珍奶 L 改 65」「電話是 02-2345-6789」「刪掉季節限定那一類」。
- 回「確認」建檔；同名店家會更新菜單，聯絡方式只會補上新的。
- 其他：「預覽」「取消」「店家」「刪除店家 店名」。

**群組點餐**
1. 把機器人拉進朋友群組。
2. 輸入「開團」→ 點選店家 → 機器人貼出點餐連結。開團的人就是收款人。
3. 每個人點連結點餐、送出；「總表・付款」分頁大家都看得到誰點了什麼、該付多少、付了沒，每 6 秒自動更新。
4. 付款後本人按「我付了 LINE Pay／現金」→ 狀態變「待確認」→ 開團的人按「確認收款」→ 變「已收款」。
5. 群組指令：
   - 「誰沒付」：列出未付款、待確認、已收款名單
   - 「運費 60」「折扣 30」：平均分攤給每個人
   - 「收單」：停止點餐；「重新開放」：恢復點餐
   - 「結單」：這團結束，才能再開下一團
   - 「點餐」：重貼連結；「說明」：指令列表

## 常見問題

- **Webhook Verify 失敗**：確認網址結尾有 `?k=` 且和指令碼屬性的 `WEBHOOK_KEY` 一樣；部署的存取權要選「所有人」。
- **朋友打開點餐頁顯示錯誤或要求權限**：確認 LINE Login channel 已 Publish。
- **點餐頁顯示「登入已過期」**：關掉頁面，從群組連結重新開啟。
- **辨識結果不準**：照片拍正、避免反光；或分區拍近一點再連續傳。
- **想直接改資料**：打開試算表就能改，例如 stores 分頁的電話；菜單存在 `menuJson` 欄位，建議還是用私訊修改比較不容易改壞。
- **費用**：機器人用「回覆」訊息不計入官方帳號免費則數；只有辨識太久（超過回覆時限）時才會改用推播回傳結果。
