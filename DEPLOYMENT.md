# 部署資訊（已上線）

部署日期：2026-06-25

## 網址

- **客人集點頁（給客人掃的 QR 連到這裡）**
  `https://script.google.com/macros/s/AKfycbx3fsiH5m2r1vWo6WV1CM8WwXwRf5f4OjlQ9R5iint6vcRZgUOqUHPcONFEGVtgwCb8/exec`

- **店員管理頁（看／列印今日通關碼）** — 同網址加 `?page=admin`
  `https://script.google.com/macros/s/AKfycbx3fsiH5m2r1vWo6WV1CM8WwXwRf5f4OjlQ9R5iint6vcRZgUOqUHPcONFEGVtgwCb8/exec?page=admin`
  建議把這個網址加到店員手機書籤。

## 預設密碼（請盡快到試算表 Config 工作表改掉）

- `adminPin`：2468（管理頁登入用）
- `staffCode`：8888（客人集滿時，店員兌換用）
- `codeSalt`：自動產生的隨機字串，**請勿外流、勿更動**（改了整月的碼會全變）

## 目前設定（Code.gs 的 CFG）

- 集滿 10 點換「炸物拚盤＋Snoopy 鑰匙圈（送完為止）」
- 最低消費：平日（週二–週五）$180、假日（週六、日）$300
- 今日通關碼：系統依日期自動產生（CODE_MODE = auto）
- 定位：記錄模式（REQUIRE_LOCATION = false，先記座標、不擋人）
- 每支手機每天限蓋 1 點

## 改設定後要重新部署

改了 `Code.gs` 的 CFG 後：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署。網址不會變。

## 相關資產

- 集點 QR：`MWDImages/MWD集點卡_QRcode.png`
- 今日通關碼卡（含 QR）：`MWDImages/今日通關碼卡.png`／`.pdf`
- 資料試算表「未命名的試算表」存放 Points／Config／Log 三個工作表（在你的 Google 雲端硬碟）

## 版本紀錄

- **v4（2026-07-20）**
  - fix：修正「每支手機每天限一次」失效——`lastStampDate` 字串被試算表自動轉成日期物件，
    `String(Date)` 永遠不等於 `"yyyy-MM-dd"`，導致同一支手機同天可重複蓋點。
    改用 `normDateCell_` 統一型別後比對，並以文字格式寫入日期。
  - feat：新增「🛠 維護」選單與維護工具 `previewPhone` / `flagTestPhone` /
    `fixInflatedPoints` / `backupNow_`（Log 只標記不刪、只在 Points 改狀態、先預覽再確認）。
  - test：新增 `test/gs_test.js` 行為測試（38 項通過）。
  - 資料維護 SOP 見 `docs/部署教學.md`。
- **v3（2026-06-25）**：假日最低消費 280→300、獎勵改為「炸物拚盤＋Snoopy 鑰匙圈」。
- **v2**：`normPhone_` 去除前導 0，修正試算表把 `09xx` 存成數字導致查無點數／每日限制失效。
- **v1（2026-06-25）**：首次上線。
