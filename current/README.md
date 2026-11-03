# ThreeASC：Fail Address + FBC by BLK

[下載最新版 Excel（內含全部 VBA）](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.16.0/ASC_ThreeASC_FBC.xlsm)

新版只有三個分頁：`Config`、`Fail Address`、`FBC by BLK`。

## 使用方式

1. 使用 Windows 桌面版 Excel 開啟檔案並允許巨集，不需要安裝 Python 或額外匯入 VBA。
2. Config 的 A2:A901 與 D2:D901 為兩組 TNO 清單，兩組合併處理，TNO 不可重複。保留 22 個預設 TNO。B/C/E/F 僅供備註，Func.、電壓可空白，暫不參與判斷。
3. Config!J13 為「僅匯入 Fail Address」；Config!J16 為「僅匯入 FBC by BLK」。兩個按鈕各自選取三份 ASC。完整 LOT、WNO 必須相同；STEP 數字由小到大對應 WT1/WT2/WT3。單獨重跑 FBC 不會解析或寫入 Fail Address，單獨匯入 Address 也不解析 FBC。
4. 執行完成後另存輸出。BLK 錯誤或無資料訊息位於 Config!T2。

## 輸出

`Fail Address` 保留既有六欄地址與 WT/TNO 的 1 標記邏輯。A 欄 TNO 清單先於 D 欄清單，兩組均參與輸出。

`FBC by BLK` 集中三個 WT，不依 Func. 分頁或分組，不做 Vmin 或 Cat Fail 判斷。

| 欄位 | 內容 |
|---|---|
| A | WT1 / WT2 / WT3 |
| B | TNO |
| C / D | X / Y |
| E:IZ | BLK 0～255 |

WT1、WT2、WT3 依序追加至同頁，以藍／橘／綠底色區分。各 WT 內按 X/Y 排序，同 Chip 依設定 TNO 順序輸出。只輸出實際存在的有效 FBC 記錄；找不到記錄不會補成 0。

## Raw Data 讀取規則

- 只解析含 `FBC BY BLK=` 的記錄；同一標頭行必須包含該筆 `Tno=`、`X=`、`Y=`，`DUT=` 不參與索引。等號後依序讀取恰好 256 個非負整數，數值允許跨行，支援空白、Tab、逗號、分號分隔；全形與不換行等特殊空白先正規化。一般說明行的 TNO/X/Y 字樣不再觸發標頭檢查，也不沿用前筆 Chip 座標。
- 第 1 個數字直接對應 BLK0，最後一個數字對應 BLK255；沒有從 Fail Address 重新計算 Count。
- TNO=0 或未選定 TNO 的 FBC 會略過，不再因此報 `TNO must be positive`。
- 選定 TNO 的缺失座標、無效數字、數量不足／超過 256 仍會報錯，並將該 WT 的 FBC 視為無效，避免輸出部分資料。其他 WT 與 Fail Address 保持獨立處理。
- 同一份 ASC 的同一 TNO/X/Y 重複出現時保留第一筆，不加總、不覆蓋。每份檔案使用獨立索引。

## v0.16.0：獨立按鈕與記錄式 Parser

原有 Address 寫入迴圈與解析 helper 保持一致。FBC 按鈕只更新 FBC by BLK 與相關狀態，保留既有 Address 資料；輸出檔名附加 _ONLY。兩個按鈕共用防重入保護、LOT/WNO/STEP 檢查與讀檔工具。

## 驗證與限制

已測試原有 Fail Address 1/3/9/900 個 TNO 的對應、全部 256 個 Raw Index、三 WT 隔離、重複／零值／異常資料、TNO0、分塊讀取，以及 3,078 列 FBC 的實際批次輸出迴圈。XLSM 內嵌 VBA 與來源一致，Office 2016 OpenXML 驗證通過。

尚未在 Excel 2016 實機執行，也未取得使用者實際 ASC 進行驗證。iPad 可查看輸出或遠端操作 Windows Excel，iPad Excel 不執行 VBA。

來源與測試放在 `current/src`、`current/tools`；`current/reference` 僅保存地址回歸測試所需的舊版基準，不是另一份可執行工具。`docs/coordinates.json` 與 wafer 示意圖保留原座標資訊。舊程式仍可從 Git 歷史取得，主分支只保留新版工具。
