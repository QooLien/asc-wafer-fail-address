# v0.13.0 — Import Three ASC + BLK Vmin

下載 **ASC_ThreeASC_BLK_Vmin.xlsm**，兩個正式 VBA 模組已內嵌。完整交付包包含程式、說明、三份合成 ASC 與預期結果。

此版本不含 Split、WT Sum、WT Fail 分類或 Conflict 功能。保留原 Import Three ASC 的完整 LOT/WNO 檢查、STEP 數字排序與 Fail Address 輸出，增加 WT1_BLK_Distribution、WT2_BLK_Distribution、WT3_BLK_Distribution。

- Config A/B：共同 TNO / Test Voltage，依輸入順序判定。
- Config N2：VminFBCThreshold；空白/0 為 Legacy，正整數為累積門檻。
- Config N6:O...：可選 DUT X/Y；留白則各 WT 按有效 FBC 資料發現 DUT 並以 X/Y 數值排序。
- 各 WT 原始 256 個 FBC 維持 BLK0→E、BLK255→IZ，最後輸出各 BLK 的 Vmin。
- 共用解碼後的資料串流，不額外重讀三份完整大檔；原地址解析及資料輸出迴圈不變。
- 缺少或無效 FBC 時，對應 BLK 頁顯示 No Data 或錯誤，保留地址匯入結果，不捏造 Vmin。

已通過原地址 1/3/9/900 TNO 回歸、256 BLK 逐格映射、Legacy/累積門檻、共用串流與錯誤隔離、內嵌原始碼及 Office2016 檔案格式驗證。尚待實際機台 ASC 在原生 Excel 2016 編譯及端到端執行。

專案程式位於 `threeasc-blk-vmin/`，歷史版本保持不變。
