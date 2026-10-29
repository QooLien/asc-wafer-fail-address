# v0.12.0 — WT 獨立 Sum / Pass-Fail

下載 **SplitWTBySum.bas**，匯入已完成 ThreeASC 合併的結果活頁簿，執行 `SplitWTByTNOSum`。不需 Python，原 ThreeASC 匯入程式與 Fail Address 不修改。

- 每個完整地址分別加總 WT1、WT2、WT3 的指定 TNO：Sum=0 為 Pass，Sum>0 為 Fail。
- 新報告包含 WT Status、WT1 Fail、WT2 Fail、WT3 Fail。
- WT Status 列出各地址三個 Sum 與狀態，可篩選 0 或 >0；右側統計各 WT 的 Pass/Fail 地址數與 Fail TNO 標記總數。
- 各 Fail 分頁只顯示該 WT Sum>0 的地址，保留該 WT 的 TNO 明細、Sum 與 Status。
- 三個 WT 獨立判定，同一地址可出現在多個 Fail 分頁，不再互斥分類，也不產生 Conflict。
- 另存 `_WT_Sum.xlsx` 報告，同名增加流水號。Pass 按 Sum=0 定義，不另外識別 No test。

已通過 1/3/9/900 個 TNO、反向 WT 欄位、全部八種 WT 組合、多 TNO 加總、重疊 Fail 分頁及明細逐格測試。尚待在原生 Excel 2016 使用實際資料端到端驗證。

請執行新入口 `SplitWTByTNOSum`，不要使用舊版的 `SplitWTFailAddresses`，兩者分類目的不同。
