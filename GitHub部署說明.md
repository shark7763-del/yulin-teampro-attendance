# TeamPro 體育班出勤戰情室 — GitHub Pages 部署說明

本版本採用和「雄麒道館點名 Pro」相同架構：

- 前端 App：GitHub Pages
- 雲端資料：Google Sheet
- 後端 API：Google Apps Script Web App `/exec`
- 手機安裝：iPhone / Android 可加入主畫面

## 1. Apps Script 後端

1. 打開 Google 試算表。
2. 擴充功能 → Apps Script。
3. 將本資料夾的 `Code.gs` 全部貼到 `程式碼.gs`。
4. 執行一次 `upgradeSchema()`。
5. 若是全新試算表，先執行 `setup()`，再執行 `importRoster()`。
6. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署。

目前前端預設連線到：

```text
https://script.google.com/macros/s/AKfycbw57gZN1eOqmfCfsRl5efB8eCheZxx3hO7C-w16DL82KWVxy1zkoplEAevwSudE1riA/exec
```

如果未來換 Apps Script 部署網址，請修改 `index.html` 內的 `DEFAULT_API_URL`。

## 2. GitHub Pages 前端

GitHub repo 推送後，到 GitHub：

1. Settings → Pages
2. Build and deployment → Source 選 `Deploy from a branch`
3. Branch 選 `main`
4. Folder 選 `/ (root)`
5. Save

完成後網址會類似：

```text
https://shark7763-del.github.io/yulin-teampro-attendance/
```

## 3. 手機安裝

### Android / Chrome

1. 用 Chrome 開啟 GitHub Pages 網址。
2. 選單 → 加入主畫面 / 安裝應用程式。

### iPhone / Safari

1. 用 Safari 開啟 GitHub Pages 網址。
2. 分享 → 加入主畫面。

## 4. 注意事項

- GitHub Pages 只負責顯示 App。
- 所有學生、點名、統計資料仍寫入 Google Sheet。
- 改 `Code.gs` 後必須重新部署 Apps Script 新版本。
- 改 `index.html` / `manifest.json` / `service-worker.js` 後，推送 GitHub 即可更新前端。
