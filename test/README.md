# 測試

`Code.gs` 的行為測試（用假的 Google 服務在本機跑，驗證邏輯正確）。

## 執行

```bash
node mwd-loyalty-card/test/gs_test.js
```

全部通過會印出 `RESULT: 38 passed, 0 failed`（結束碼 0）；有任何失敗結束碼為 1。

## 涵蓋範圍

- **每天限一次**：同號碼同日第二次蓋章被擋；`lastStampDate` 以純文字儲存、不會再被試算表轉成日期物件。
- **舊資料相容**：就算 `lastStampDate` 是日期物件也能正確判斷（`normDateCell_`）。
- `previewPhone`（唯讀不改資料）、`flagTestPhone`（Points 刪、Log 只標記、不動到別人）、`fixInflatedPoints`（改點數＋Log 留紀錄）、`backupNow_`、`ensureLogNoteCol_`。

## 限制

這是用**模擬**的 `SpreadsheetApp` / `LockService` 等（重現了「字串日期被自動轉成 Date」的坑）。它驗證的是**邏輯**，不含真實 Apps Script 環境的行為（真實鎖競爭、`ss.copy()` 的 Drive 權限/配額、跨午夜時區）。改完 `Code.gs` 後，除了跑這個測試，仍建議到 Apps Script 編輯器實跑一次「🛠 維護 → 一鍵備份」做煙霧測試。
