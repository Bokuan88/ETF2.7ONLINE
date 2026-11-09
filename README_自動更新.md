# ETF 持股每日自動更新（00403A／00981A／00988A／00991A）

## 安裝（一次就好）
1. 把本包裡的 `scripts/` 與 `.github/` 兩個資料夾，原樣放到你的 GitHub repo 根目錄（和 index.html、holdings/ 同一層），commit & push。
2. GitHub repo → Settings → Actions → General → Workflow permissions，選 **Read and write permissions**，Save。
3. GitHub repo → Actions → 「每日更新 ETF 持股」→ Run workflow，先手動跑一次，看 log 是否出現 ✅。

之後每週一至週五台北 18:30（另有一次隔日 08:30 補抓）自動執行：下載 → 解析 → 只新增還沒有的日期 → 寫入 `holdings/` → 自動 commit。網頁重新整理就會讀到新資料。

## 帳戶名稱對應
預設會找「名稱含該 ETF 代號」的既有帳戶（例如 `00981A統一增長`）；找不到才用代號新建（例如 `00403A`）。
若你的帳戶名稱不含代號，請編輯 `scripts/update_holdings.mjs` 最上方的 `ETFS`，填入 `accountName`（要和 holdings 裡的帳戶名稱完全一致）。

## 本機手動執行
    npm i xlsx
    node scripts/update_holdings.mjs
