# v0.11.1 — 第二階段分割加速

下載 **SplitThreeASCResults.bas**。請先移除舊的 SplitThreeASCResults 模組，再匯入新檔，執行 `SplitWTFailAddresses`；既有 ThreeASC 匯入流程不變。

- 資料陣列改用 ByRef 傳遞。
- 空白格略過字串轉換，Excel 數值 Double 的 1 直接判斷。
- 快取第一次分類結果，Conflict 列不再重複掃描。
- 每八批更新狀態與處理介面事件，減少介面更新成本。

保留 WT1/2/3 Fail、WT Conflict 分類、原始欄位與儲存格內容、地址驗證及數量核對。

已通過 1/3/9/900 個 TNO 的逐格測試，並與 v0.11.0 的程式比對分類及地址鍵值；另涵蓋 Double、整數、文字、空白與非法值處理。尚未在原生 Excel 使用實際資料量測速度，不提供加速倍數承諾。
