# WT 三分頁分類版（v0.10.0）

範本 `templates/ASC_Fail_Address_Classified.xlsm`；完整 VBA `src/ImportThreeASC_Classified.bas`。與前一版獨立，原 Fast 與三份合併版均不更動。

在 Windows Excel 2016 或更新版本開啟範本、允許巨集，在 Config!A2:A901 指定 TNO，按 C11 的「選擇 3 份 ASC 並分類」。不需安裝 Python。

沿用三份完整 LOT / WNO 一致檢查、STEP 數字排序、讀檔、地址有效範圍及選定 TNO 解析。只有 Config 指定 TNO 參與判定；其餘 TNO 和 monitor item 不參與。分類單位為完整 X/Y/BLK/WL/BL/IO 地址。

分頁順序為 Config、WT3 Fail、WT2 Fail、WT1 Fail：

- WT3 Fail：WT3 至少一個指定 TNO 有 1，WT1/2 所有指定 TNO 都沒有 1。
- WT2 Fail：WT2 至少一個指定 TNO 有 1，WT1/3 都沒有 1。
- WT1 Fail：WT1 至少一個指定 TNO 有 1，WT2/3 都沒有 1。

每頁保留六欄地址及三組 WT/TNO 分布，資料仍從 G2 起。標題與值為 1 的資料格以 WT1 淡藍、WT2 淡橘、WT3 淡綠區分，保留篩選、Calibri 置中與固定列高。Config D7:D9 顯示各類地址數，G2:L4 保留來源對應。

同一地址若跨兩個或三個 WT 有指定 TNO Fail，即使 TNO 不相同也視為流程衝突；在清除舊輸出之前停止，顯示衝突數及前五個地址，不自動選擇一類。全無指定 Fail 的地址不輸出。Pass / No test 是依使用者確認的測試流程及所選 TNO 記錄解讀，不另外從原始資料的完成旗標判斷。

輸出另存至 STEP 最小的 ASC 所在資料夾，例如 `FAX768#07_WT_Classified.xlsm`，同名增加流水號。

## 驗證

`tools/test_classified_import.py` 使用實際分類函式、解析函式、標題及三分頁批次輸出程式，經語法轉換後在 VBScript 執行。獨立預期值逐格驗證地址、WT、TNO，涵蓋 1/3/9/900 個 TNO、monitor 排除、位元組與組別交界、全部 WT 出現組合、不同 TNO 跨 WT 衝突、完整/部分批次、空分頁及全無 Fail。

解析及 phase mapping 函式與前版逐字一致；XLSM 內嵌巨集獨立抽取核對，並經 Microsoft Open XML SDK Office2016 格式驗證。未在原生 Excel 2016 使用實際 ASC 端到端執行；測試不等於實際資料已確認。

手動匯入 `.bas` 時請使用獨立 `.xlsm`，建立 Config、WT3 Fail、WT2 Fail、WT1 Fail 分頁，匯入模組並執行 ImportASCFailAddresses；避免匯入原版的同一 VBA 專案造成模組名稱衝突。
