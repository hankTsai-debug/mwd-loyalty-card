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

- 集滿 10 點換指定餐點或飲料一份
- 最低消費：平日（週二–週五）$180、假日（週六、日）$280
- 今日通關碼：系統依日期自動產生（CODE_MODE = auto）
- 定位：記錄模式（REQUIRE_LOCATION = false，先記座標、不擋人）
- 每支手機每天限蓋 1 點

## 改設定後要重新部署

改了 `Code.gs` 的 CFG 後：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署。網址不會變。

## 相關資產

- 集點 QR：`MWDImages/MWD集點卡_QRcode.png`
- 今日通關碼卡（含 QR）：`MWDImages/今日通關碼卡.png`／`.pdf`
- 資料試算表「未命名的試算表」存放 Points／Config／Log 三個工作表（在你的 Google 雲端硬碟）
