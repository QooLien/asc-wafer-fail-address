# v0.9.0 — 三份 ASC 合併版

下載 **ASC_Fail_Address_ThreeASC.xlsm**，VBA 已內嵌，不需 Python。另提供完整 **ImportThreeASC.bas** 原始碼。

- 一次選取三份 ASC，驗證完整 LOT、正規化後的 WNO 一致。
- 依 STEP 的數字由小到大分配 WT1、WT2、WT3；缺少數字、多段數字或數值重複時停止。
- 以 X/Y/BLK/WL/BL/IO 合併地址聯集，同一地址一列，三組 TNO 旗標分開保留。
- WT1 淡藍、WT2 淡橘、WT3 淡綠；Fail 值 1 以對應顏色顯示。標題包含 WT 與 TNO，開啟篩選。
- Config!G2:L4 保留各 WT 來源、STEP、版本及記錄統計。
- 另存至 STEP 最小的 ASC 所在資料夾，例如 `FAX768#07_WT1-WT3.xlsm`；同名增加流水號。
- 前一版 Fast 活頁簿與 VBA 原始碼不變。新版本的讀取/解析輔助函式與批次資料輸出迴圈沿用原版。

已通過：1/3/9/900 個 TNO 的獨立地址/WT/TNO 逐格測試、分段與組別/位元組邊界、排序與檔案 metadata 對應、LOT/WNO 不一致和 STEP 異常檢查、內嵌 VBA 一致性及 Microsoft Open XML SDK Office2016 格式驗證。

尚未在原生 Excel 2016 使用實際三份 ASC 進行端到端驗證。請先確認 Config 的 WT 來源對應，並抽查實際輸出。舊版單檔匯入功能請繼續使用 v0.8.0 的 Fast 範本。
