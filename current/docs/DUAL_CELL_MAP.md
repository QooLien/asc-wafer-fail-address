# Fail BLK Address Map

新增單一分頁；每個 Pattern 為一個區段，每區依已匯入 WT 水平排列。每個 Chip 佔左右兩格。

- 左格 Fail：直接取 BLK Fail Count（原 Scan/March Map 相同判斷）。
- 右格 Loss：Missing BLK Address 中，按 WT / Pattern / X / Y 精確分組，Missing Count > 0 的不同 BLK 數。來源核對已採 WT / TNO / X / Y / BLK 精確匹配。
- No Address 與 Partial Address 都算短缺 BLK；Complete 不算；同一 BLK 不重複加總。不加總 Missing Count 的 Fail Bit 筆數。
- 例如：5 個 Fail BLK，BLK236 缺少 Address，顯示 5 | 1；即使該 BLK 缺少 100 筆仍算 1。
- Cat. 統一取 WT1 第一份 Raw Data，套用 Fail Address G 欄及所有 WT Map；WT1 缺少分類時不以 WT2/WT3 補值。
- C1/C2 兩格空白、不填色；其他類別 No Data 延用灰底；數字零隱藏、正數黃色至紅色色階。
- Calibri 24，X/Y 軸 20；隱藏分頁格線。
- 主匯入及更新 Map 自動重新核對 Address 並建立新 Map。核對失敗清除新 Map，避免顯示過期數字。

## 驗證

Excel 原生執行 1/2/3 ASC，所有原有 TNO Address flags 逐格驗證通過，3→1 重跑通過。新 Map 的 24 組 WT/Pattern/Chip 左右數值與獨立統計一致。額外驗證 5|1、5|2（部分缺失）、BLK 去重、Complete 排除、C1/C2 空白、24/20 字級、核對失敗後清除過期數字。已視覺檢查 Excel 匯出預覽；正式 XLSM ZIP 完整且不含測試入口。
