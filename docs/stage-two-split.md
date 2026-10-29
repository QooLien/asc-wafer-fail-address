# ThreeASC 完成後的第二階段拆分

程式：`src/SplitThreeASCResults.bas`；入口：`SplitWTFailAddresses`。適用 Excel 2016，不需 Python。這是獨立模組，既有 ThreeASC 匯入程式不變。

## 使用

1. 使用 v0.9.0 `ASC_Fail_Address_ThreeASC.xlsm` 跑完三份 ASC，先確認其 Fail Address 合併結果。
2. 在這份已完成的結果活頁簿按 Alt+F11，選「檔案 → 匯入檔案」，匯入 `SplitThreeASCResults.bas`。模組名不同，不會取代原匯入模組。
3. 返回 Excel，按 Alt+F8 執行 `SplitWTFailAddresses`。
4. 程式讀取此活頁簿的 Fail Address，建立另一份報告，包含 Summary、WT1 Fail、WT2 Fail、WT3 Fail、WT Conflict。

不修改原始 Fail Address、Config 或既有輸出頁。報告存於原活頁簿資料夾，檔名加 `_Split.xlsx`；同名加流水號。來源未存檔或使用 URL 路徑時，報告留在開啟狀態讓使用者手動另存。來源路徑過長、資料不合法或存檔失敗會提示並關閉未完成報告，原合併頁保留。

## 分類

讀取 ThreeASC 輸出的標題（例如 WT1 換行 100）尋找 WT，不使用固定第幾欄的假設。三組必須包含相同的非空 TNO 集合，不允許重複 WT/TNO 標題。以輸出標題代表此次原本選定的 TNO，不重新讀取可能已被改動的 Config。

- 只有 WT1 有 1 → WT1 Fail。
- 只有 WT2 有 1 → WT2 Fail。
- 只有 WT3 有 1 → WT3 Fail。因此此頁 WT1/2 全部指定 TNO 都沒有 1。
- 同一完整地址在兩個或三個 WT 有 1 → WT Conflict，不中止其他正常地址分類。
- 全部 WT 無 1 → WT Conflict，原因另標 No selected TNO fail；原 ThreeASC 正常輸出理論上不會有這種列，但仍保留以維持地址總數。

原始每列 A 至最後一個 TNO 欄直接逐格複製，欄位位置不變，不重建 TNO 索引，也不重新解析 ASC。Conflict 頁額外增加一欄原因。每個地址只分配一個分頁。

Summary 檢查：原始地址數 = WT1 Fail + WT2 Fail + WT3 Fail + WT Conflict。遇到重複地址、缺欄、非空白/1 的 TNO 值等輸入錯誤會停止，不將資料錯誤偽裝為流程衝突。

## 驗證範圍

`tools/test_split_results.py` 擷取實際標題驗證、分類與列複製程式，語法轉換後以 VBScript 測試。涵蓋 1/3/9/900 個 TNO、反向 WT 欄位順序、全部八種 WT Fail 組合、Conflict 原因、三分頁與 Conflict 的逐格原值保留、地址無遺失/重複、重複標題與非法 Fail 值拒絕。VBScript 的測試 fixture 不包含 Excel Error Variant，IsError 檢查以固定 false 轉接；原 VBA 仍保留原生 IsError 檢查。

尚未在原生 Excel 2016 以實際資料執行；測試驗證邏輯與資料搬移，不等同已驗證 Excel UI、存檔或實際來源檔。此版以 v0.11.0 發布程式碼，未更動先前活頁簿。
