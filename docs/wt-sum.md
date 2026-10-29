# WT 獨立 Sum / Pass-Fail 報告（v0.12.0）

程式：`src/SplitWTBySum.bas`；執行入口 `SplitWTByTNOSum`。適用 Windows Excel 2016，不需 Python。

先用 ThreeASC v0.9.0 完成合併，在該結果活頁簿按 Alt+F11 → 檔案 → 匯入檔案，匯入本模組；再按 Alt+F8 執行 SplitWTByTNOSum。這是另一個模組，不會覆寫前版。不要執行前版 SplitWTFailAddresses，以免得到舊的互斥分類結果。

## 判定

直接讀取 Fail Address，依標題辨識 WT/TNO。這些欄位代表 ThreeASC 執行時 Config 選取的 TNO，不再讀取可能已被修改的 Config。

每個完整 X/Y/BLK/WL/BL/IO 地址，三個 WT 各自加總指定 TNO 的 1。Sum=0 輸出 Pass，Sum>0 輸出 Fail。沒有跨 WT 互斥或優先權：同一地址在兩個 WT Fail，就同時出現在那兩個 Fail 分頁。不再產生 WT Conflict。

此處 Pass 按使用者定義表示該範圍 Sum=0，不另外識別 No test。來源仍只接受空白或 1；原始資料欄填 0、其他數值或錯誤值時會提示，避免把異常資料當成正確結果。

## 輸出

原 Fail Address 不修改，另存一份 `_WT_Sum.xlsx` 報告；同名加流水號：

- WT Status：所有地址，六欄地址 + WT1 Sum / Status + WT2 Sum / Status + WT3 Sum / Status。開啟篩選，可獨立篩選每個 Sum 的 0 或 >0；右側統計各 WT Pass 地址數、Fail 地址數、Fail TNO 標記總數。
- WT1 Fail / WT2 Fail / WT3 Fail：只保留該 WT Sum>0 的地址，六欄地址 + 該 WT 的 TNO 明細 + Sum + Status。保留來源對應的 TNO 標題及原始值，以淡藍、淡橘、淡綠區分。

每個 WT 的 Pass 地址數 + Fail 地址數 = WT Status 地址數。不同 Fail 頁允許同一地址重複，所以三頁列數總和可能大於原始地址數，這是預期行為。來源地址重複、標題不合法等仍停止並保留原始資料。

## 驗證

`tools/test_wt_sum.py` 擷取實際 Sum、欄位映射及每 WT 明細輸出程式，以 VBScript 測試 1/3/9/900 個 TNO、反向 WT 欄位、全部八種 WT 組合、多個 Fail TNO 的真實總和、跨 WT 重複列、逐個 TNO 值與地址一致、文字/空白及非法值處理。

測試使用獨立 fixture 預期值。尚未在原生 Excel 2016 使用實際資料端到端執行，也未量測實際速度。原 ThreeASC 與舊第二階段程式保持不變。
