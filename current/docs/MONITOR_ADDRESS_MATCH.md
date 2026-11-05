# Function Fail Address 與 Write/Read Low Monitor 匹配

## 分析規則

1. 僅分析原 Map / BLK Fail Count 納入的 Fail BLK，沿用 FBC、Hard Fail、前序 Pattern/WT 的既有判斷。
2. 同一 WT / Chip / BLK，收集該 Pattern 區間內所有 Function TNO 的 Fail Address。區間為 Config 列順序中，前一 Stop Sign / Hard Fail 之後，到當前 Stop Sign（含）。同一完整 Address 跨 Function TNO 去重。
3. 使用完整 X/Y/BLK/WL/BL/IO，在同一 WT 的 Fail Address 數值 TNO 欄位比對。不同 WT、Chip、BLK、WL、BL、IO 不可混用。
4. Config A:C 中 Func. = W-Low pBL / R-Low pBL 為 Write / Read Monitor；忽略大小寫及空格、連字號、底線差異。兩類 Monitor 都必須有設定。
5. 各類所有設定電壓中，任一 Monitor TNO 的同一 Address 旗標為 1，即視為該類 Monitor Fail，不做電壓一對一限制。
6. 分類互斥：Write Fail、Read Fail、Write&Read Fail、No monitor item Fail。四類數量合計等於實際可分析的 Pattern Address Count。
7. No monitor item Fail 僅表示沒有記錄到對應 Monitor Address 的 1；未輸出的 Function Address 不分類，不能視為 Monitor 無 Fail。

## 輸出與操作

- 匯入並分析、更新 Map / Address 核對時自動產生。
- Monitor Fail Analysis：逐筆完整 Address、Pattern、Stop TNO、命中的 Function TNO 清單、Write/Read Monitor TNO 清單、兩類旗標、分類及原 Fail Address 列號。可按 WT / Pattern / X / Y / BLK 篩選。
- BLK Analysis：每個 Map Fail BLK 一列，含 Chip Fail BLK Count、FBC、實際 Pattern Address Count、四類 Address 數、終點缺失數、Coverage Basis 及 Analysis Status。
- 沒有 Function Address 的 BLK 顯示 No Function Address，四類數量留空，不填入 No monitor item Fail。
- 只匯入 Address / FBC 後，既有 Monitor 分析標示待更新，不自動混用舊來源。來源或數值 TNO 索引不一致時拒絕分類，並清除不完整輸出。
- Calibri 11、置中、標題由下至上、有篩選。明細 TNO 清單換行，資料列高一致；Chip Locations 及 Log 維持最後兩頁。

## 驗證

原生 Excel 1/2/3 ASC 全流程與既有 Address 旗標逐格回歸、3→1 重跑通過。另以獨立測試逐筆驗證 1004 筆分類及全部 BLK 彙總，涵蓋四類、Pattern 前站有而終點無、跨 TNO 去重、不限定 Monitor 電壓、不同 WT/Chip/BLK/BL/IO/Pattern 排除、缺少 Function Address 不分類。故意破壞 Monitor TNO 索引可正確拒絕並清除過期分類。已視覺檢查兩個輸出頁籤；正式 XLSM 完整且沒有測試入口。

最新版已將區間/終點Address覆蓋診斷合併到BLK Analysis Q:U；四格Map輸出分類Address數量，參見COMPACT_COUNTS.md。
