# 三份 ASC 合併版（v0.9.0）

Excel：`templates/ASC_Fail_Address_ThreeASC.xlsm`，VBA 已內嵌。
完整 VBA：`src/ImportThreeASC.bas`，可匯入獨立活頁簿的 VBA 專案。

## 使用

1. 在 Windows Excel 2016 或更新版本開啟新範本並啟用巨集，不需要安裝 Python。
2. Config!A2:A901 每列輸入一個 TNO，規則與 Fast 版一致。
3. 按「選擇 3 份 ASC 並合併」，一次選取三個檔案。
4. 檢查 Config!G2:L4 的 WT、原始路徑、STEP、來源版本及各檔統計。
5. 結果另存到 STEP 最小的 ASC 所在資料夾，例如 `FAX768#07_WT1-WT3.xlsm`，同名增加流水號。

## 合併及欄位

完整 LOT（不截成六碼）須完全一致，WNO 正規化整數後比較，因此 7、07、#07 相同。
STEP 必須恰有一段連續數字，按該數字由小到大分配 WT1、WT2、WT3；例如 WTT2、WTT30、WTT100。缺少數字、多段數字、重複數值或超過 Long 範圍時停止。三份各自的 Program 版本沿用第一個 VER_ 後兩個英文字母規則，可不同，來源對應表保留各自資訊。

Fail Address 保留單列標題及第 2 列起資料。A:F 仍為 X/Y/BLK/WL/BL/IO；G 欄起依 WT1、WT2、WT3 分成三組，每組依 Config 的 TNO 輸入順序排列。標題顯示 WT 與 TNO（換行、由下向上）。WT1 淡藍、WT2 淡橘、WT3 淡綠；只有值為 1 的資料格以對應條件格式填色。

三份檔案有效選定地址取聯集，同一六欄地址只輸出一列。各 WT 的 Fail 旗標分開，不做跨 WT 的 OR 合併。空白只代表該檔案沒有有效 Fail 記錄，不代表已測試或 Pass。每份均須找到至少一個指定 TNO 區段；若某份選定記錄全部無效則停止，輸出不取代舊結果。沒有 Fail 記錄但有選定區段可正常輸出空白。

## 與前版一致性

舊 `ImportFailAddresses.bas` 與 `ASC_Fail_Address_Fast.xlsm` 不變。新版本完整沿用前版從 DetectEncoding 起的所有函式，包含二進位分段讀取、標頭擷取、地址解析、有效範圍、地址索引、位元旗標、非法記錄處理及檔名安全處理。批次輸出資料迴圈也不變。

新功能只在主流程增加三檔預檢、STEP 排序、每組 TNO 索引偏移及三組欄位/格式。索引為 `(WT序號-1) × TNO數 + 原TNO索引`；標題與資料都使用相同索引。跨檔時重設解析緩衝區與目前 TNO，保留共同地址字典及旗標。

每檔上限 2 GB；地址聯集最多 1,048,575 列；最多 900 個 TNO × 3 組，仍在 Excel 2016 欄數上限內。記憶體用量隨地址數與 TNO 數增加；未進行原生 Excel 大檔速度/記憶體測試。

## 驗證

`tools/test_three_import.py` 擷取實際新 VBA 的分組索引、標題、解析及批次輸出，經語法轉換後在 VBScript 測試。預期值由獨立 fixture 的地址、WT、TNO 建立，逐格確認 1 與空白，涵蓋 1/3/9/900 個打亂順序的 TNO、分段邊界、組別交界、位元組交界、共同與單次獨有地址、容量擴增及完整/部分批次。另測試數字排序、檔案與 metadata 同步、完整 LOT 不同但前六碼相同、WNO 不同、WNO 補零、無效與重複 STEP。

XLSM 經獨立 VBA 抽取比對及 Microsoft Open XML SDK Office2016 格式驗證；尚未在原生 Excel 2016 使用實際三份 ASC 執行。這些測試支持索引與資料邏輯一致，但不替代實際檔案的端到端確認。

若手動使用 `.bas`，請建立獨立 `.xlsm`，準備 Config、Fail Address 兩張工作表，透過 VBA 編輯器「匯入檔案」匯入此模組，執行 `ImportASCFailAddresses`。模組名沿用 ImportFailAddresses，請勿匯入前版 Fast 的同一專案，以免名稱衝突；建議直接使用已內嵌範本。
