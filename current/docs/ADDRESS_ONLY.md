# 僅匯入 Address

Config 新增「僅匯入 Address」按鈕，位於既有按鈕下方，Calibri 11 並置中；原有按鈕位置保留。

- 選取 1～3 份 ASC，依原有 LOT/WNO/STEP 驗證與 WT 排序。
- 使用原有 Address 解析及 bit flags 核心，依 Config TNO 輸出 A:F Address、G Cat.、H 起各 WT/TNO。
- Cat. 只取 WT1 第一份資料，不以 WT2/WT3 補值；採輕量 CATLST 逐行解析，不建立 FBC 索引。
- 不解析 FBC、不輸出 FBC、不重算比較表、Map 或 Missing BLK Address 明細。既有核對與雙格 Map 標示待更新，按「更新 Map」可重新計算。
- 更新 Address 數值 TNO 索引及 Address 來源，不變更 FBC 來源。Address 的 WT 數量獨立於舊 FBC，支援舊 FBC 為三份但 Address 僅匯入一份。
- 另存 *_Address.xlsm，完成視窗顯示執行時間。

## 驗證

Excel 原生測試 1/2/3 ASC：獨立 Address 輸出與完整分析逐值一致；FBC、比較表、原 Map、Map 計數、Chip Locations、核對明細及 FBC manifest 不變。3→1 Address 重跑與數值 TNO 索引正確。另以無效 FBC 與 WT1 C3／WT2 C1／WT3 C2 資料確認獨立流程跳過 FBC 且 G 欄取 WT1 C3。正式工作簿已建立可用按鈕並視覺檢查，ZIP 完整且未包含測試入口。
