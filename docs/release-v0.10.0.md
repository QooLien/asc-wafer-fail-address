# v0.10.0 — WT Fail 三分頁分類版

下載 **ASC_Fail_Address_Classified.xlsm**，VBA 已內嵌、不需 Python。另提供完整 **ImportThreeASC_Classified.bas** 原始碼。

- 沿用三份 ASC 的完整 LOT / WNO 一致檢查與 STEP 數字排序。
- 只針對 Config 指定 TNO 判定，其他 TNO 與 monitor item 不參與。
- 依完整 X/Y/BLK/WL/BL/IO 地址分類，輸出 WT3 Fail、WT2 Fail、WT1 Fail 三分頁。
- WT3 Fail：WT3 有 1，WT1/2 全無 1；WT2 Fail：WT2 有 1，WT1/3 全無 1；WT1 Fail：WT1 有 1，WT2/3 全無 1。
- 各頁保留三組 WT/TNO 分布、配色與篩選，方便核對其他 WT 欄是否空白。
- 同一地址跨多個 WT 出現指定 TNO 的 1 時，停止並顯示衝突數及前五個地址；檢查在清除舊輸出前執行。
- Config 顯示各分類地址數與來源檔案對應。結果另存為例如 `FAX768#07_WT_Classified.xlsm`，同名加流水號。
- 前一版 Fast、三份合併版及其 VBA 檔案保持不變。

已通過 1/3/9/900 個 TNO 的獨立逐格分類/欄位測試、monitor 排除、跨 WT 衝突、空分頁及無 Fail 情境、內嵌 VBA 原始碼核對，以及 Office2016 格式驗證。

尚未使用實際三份 ASC 在原生 Excel 2016 進行端到端驗證。Pass / No test 依使用者確認的流程及指定 TNO 記錄解讀。
