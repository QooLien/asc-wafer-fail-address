# v0.5.0 — Calibri 置中排版與匯入快捷按鈕

- 所有分頁文字統一為 Calibri，水平及垂直置中。
- Fail Addresses 的標題及資料列固定 26 pt，匯入後持續維持相同列高。
- Config 新增「選擇 ASC 並匯入」原生 Forms 按鈕，已連結 ImportASCFailAddresses。
- 統一深藍標題、淺色說明區與淡黃 TNO 輸入欄，整理 Config 的狀態及操作區。
- 新增 Log 時維持 Calibri 與置中；Log 長內容仍換行並調整列高，便於診斷。

已驗證：儲存檔排版、按鈕文字/巨集連結、20 種解析案例及分段邊界、內嵌 VBA 一致性；Microsoft Open XML SDK Office2016 格式驗證 0 錯誤。

尚未在桌面版 Excel 2016 點擊按鈕或執行實際 ASC 匯入測試。VBA 已內嵌，操作按鈕前須依公司政策允許巨集執行。
