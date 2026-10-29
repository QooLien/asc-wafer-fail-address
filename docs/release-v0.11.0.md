# v0.11.0 — ThreeASC 結果第二階段拆分 VBA

本次提供 **SplitThreeASCResults.bas**，不是新的匯入 Excel 範本。

1. 先用 v0.9.0 的 ThreeASC 範本完成三份 ASC 合併並保存結果。
2. 在該結果活頁簿按 Alt+F11，透過「檔案 → 匯入檔案」匯入附件 VBA。
3. 按 Alt+F8 執行 `SplitWTFailAddresses`。

程式讀取既有 Fail Address，另建報告，包含 Summary、WT1 Fail、WT2 Fail、WT3 Fail、WT Conflict。原始合併頁與匯入程式保持不變。跨 WT 的 Fail 會進入 Conflict，不中止正常地址分類；全無 Fail 的異常列也進入 Conflict 並註明原因。

每列原始地址及所有 WT/TNO 儲存格按原欄位複製，不重建索引、不重新解析 ASC。Summary 核對四個分類分頁地址數之和與原始地址數一致。報告在原資料夾另存為 `_Split.xlsx`，同名增加流水號。

已通過 1/3/9/900 個 TNO、反向 WT 欄位排列、全部八種 WT 組合、逐格原值保留、Conflict 原因、地址無遺失/重複及異常資料拒絕測試。尚待在原生 Excel 2016 以實際資料執行。完整說明見 docs/stage-two-split.md。
