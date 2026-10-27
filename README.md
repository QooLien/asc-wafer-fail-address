# ASC Wafer Fail Address

使用 Excel VBA 從 `.asc` Raw data 讀取指定測試項目（TNO）的 SRAM Fail Address，並保留 128-chip wafer map 與四象限計數範本。

## 快速開始

1. [下載最新版 ASC_Fail_Address_Fast.xlsm](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.6.0/ASC_Fail_Address_Fast.xlsm)，使用 Windows 電腦版 Excel 開啟。此版已停用 Log，並顯示運行時間。
2. 依公司政策允許該檔案的巨集執行；**不需要匯入 `.bas`，也不需要執行 `.bat`**。
3. 在 `Config!A2:A901` 每列輸入一個 100–999 的整數 TNO，預設為 100。
4. 在 Config 分頁按 **「選擇 ASC 並匯入」**，選擇一個 `.asc` 檔案。也保留 Alt+F8 執行 `ImportASCFailAddresses` 的方式。

`src/ImportFailAddresses.bas` 保留供版本管理與檢閱。日常請使用 **ASC_Fail_Address_Fast.xlsm**；目錄內原有的 `ASC_Fail_Address.xlsm` 與 `.xlsx` 僅保留作為舊版參考。

VBA 必須在 Windows 電腦版 Excel 執行。iPad 可遠端操作該電腦或查看輸出；iPad 版 Excel 不執行 VBA。

**Excel 2016 請下載 v0.4.1 或更新版本。** v0.4.0 及較早的 `.xlsm` 有封裝內容類型與 OPC 命名空間問題，可能顯示「檔案格式或副檔名無效」。v0.4.1 已修正，並通過 Microsoft Open XML SDK 的 Office2016 格式驗證（0 錯誤）；尚未在 Excel 2016 實機開啟或執行測試。

## ASC 格式與輸出

v0.5.0 統一所有分頁的儲存格文字為 **Calibri、水平置中及垂直置中**。Fail Addresses 標題及資料列均固定 **26 pt** 列高，匯入新資料後維持一致。Config 的匯入按鈕使用 Excel 原生 Forms 控制項，已直接連結內嵌 VBA，不需手動指定巨集。按鈕需在 Windows 桌面版 Excel 允許巨集後操作。

```text
tno= 100 dbk= 1 > [2, 100, 10, 14, 0, 25, 3, 4,]
```

方括號欄位依序為：第一欄、TNO、X、Y、BLK、WL、BL、IO。`dbk` 與第一欄不參與地址判定；TNO 必須與目前 `tno=` 區段一致。以逗號為欄位分界，讀取各欄逗號前的完整整數；接受最後的逗號、一般空格、Tab、換行、不換行空格、全形空白、同一行多組地址及跨行方括號。小數及數字中間的空格不會被強行拼接成地址。

| 欄位 | 有效範圍 |
|---|---|
| TNO | 100–999，由 Config 指定 |
| X、Y | 原始整數座標，本版不依 wafer 外形篩除 |
| BLK | 0–255 |
| WL | 0–1023 |
| BL | 0–15 |
| IO | 0–15 |

超出範圍、缺欄及非整數記錄略過。原需求中的 WLK 按 WL 解讀。

`Fail Addresses` 分頁 A1:F1 固定為 **X / Y / BLK / WL / BL / IO**。G1 起是指定的 TNO，按 Config 首次出現順序排列。同一組六欄地址只輸出一列；該 TNO 有有效 Fail 記錄填 **1**，其餘留白。重複 Fail 不累加，重複 TNO 不重複建欄。

空白表示輸入檔中未找到對應的有效 Fail 記錄，不能據此推論已測試或 Pass。每次匯入取代舊結果；取消或解析失敗時保留舊結果。正式寫入開始後若失敗，可能只完成部分輸出。錯誤對話框顯示錯誤代碼、原因、處理階段及已讀位元組數。

## 無 Log 與運行時間

v0.6.0 **完全停用 Log 輸出**，避免非地址方括號產生大量 WARN，導致 Import Log 超出 Excel 列數上限。不建立、累積或寫入逐筆紀錄；Import Log 分頁僅保留停用說明，不影響匯入。

完成視窗顯示 **Run time (hh:mm:ss)** 與總秒數。從選定 ASC 後開始計時，涵蓋讀檔、解析、合併、輸出與格式設定，不計入選檔等待時間，可跨午夜。

## 128-chip wafer 參考

| Y | X 範圍 | 每列 chips |
|---|---|---|
| 10–11 | 10–13 | 4 |
| 12–13 | 8–15 | 8 |
| 14–21 | 7–16 | 10 |
| 22–23 | 8–15 | 8 |
| 24–25 | 10–13 | 4 |

共 128 chips；X 向右增加，Y 向下增加。完整座標見 `docs/coordinates.json`。

| 左上：第二象限 | 右上：第一象限 |
|---|---|
| Read | Write |
| **左下：第三象限 ReadWrite** | **右下：第四象限 No detected Fail address** |

四格皆記錄 address count。空白表示未記錄，0 表示實際零值。四象限表位於 `templates/wafer128_SRAM_quadrants.xlsx`；目前尚未定義各 TNO 對應哪一分類，ASC 匯入不會自動回填 wafer map。

![單顆 chip 四象限](docs/chip_quadrants.png)

[查看完整 wafer map](docs/wafer128_quadrants.png)

## 大型 ASC 與記憶體

v0.6.0 每次以 **256 KB** 分段讀檔，不建立整份 ASC 字串，並快取共用的 token 正規表示式。保留跨段的 TNO、方括號地址及 UTF-8 字元。

去重使用地址字典及 Byte 陣列，每個地址與 TNO 的 Fail 狀態只佔一個 bit（另有容量預留及位元組補齊）。解析過程不逐筆操作 Excel，也不建立暫存工作表。完成後批次輸出，每批最多 131,072 個 Variant 欄位、最多 4,096 列，依 TNO 數自動調整。執行期間暫停畫面更新、事件及自動計算，結束或可攔截錯誤時恢復。

唯一地址索引、旗標及 Excel 結果仍需記憶體。單檔小於 2 GB；唯一地址最多 1,048,575 列。未閉合方括號累積超過 16,384 字元會略過。詳細最佳化與驗證範圍見 [最佳化說明](LOCAL_OPTIMIZATION.md)。

## 狀態與限制

- v0.6.0 已測試 20 種地址案例、1 / 7 / 262,144 字元分段邊界、900 個 TNO 的位元旗標、容量擴增、重複地址與滿批/不足一批的輸出，核對未 Fail 欄位保持空白。
- 約 259 MB 合成檔解析通過，保留的未解析緩衝區最高 128 字元。此檔主要為填充內容及少量資料案例，不代表大量唯一地址的 Excel 寫入效能測試。
- VBA 容器已經獨立的 `olevba` 讀回，核對內嵌原始碼與 `.bas` 完全一致；XML/ZIP、巨集關聯及文件代碼名稱已檢查。
- v0.4.1 另加入實際生效的巨集內容類型及 OPC 命名空間檢查，並通過 Microsoft Open XML SDK 2.20.0 的 Office2016 規則驗證（0 錯誤）。前版僅檢查 XML/ZIP 與 VBA 內容，未能發現影響開啟的封裝錯誤。
- v0.5.0 已檢查儲存檔中的字型、水平/垂直置中、Fail Addresses 固定列高及 Forms 按鈕的巨集連結，並通過 Office2016 格式驗證（0 錯誤）；按鈕實際點擊及巨集執行仍待 Excel 實機驗證。儲存格圖片預覽不包含原生 Forms 按鈕。
- 開發環境未安裝 Excel；上述測試不代表已完成 Excel 原生編譯、檔案選擇及真實 ASC 匯入驗證。
- 自動辨識帶 BOM 的 UTF-8 / UTF-16；無 BOM 時使用 Windows 系統文字編碼。位元組解碼與 Excel 工作表寫入尚待桌面版 Excel 實機驗證。
- 尚未驗證 iPad 端下載、開啟及遠端執行流程。

## 開發驗證

Windows 上執行 `python tools/test_parser.py` 可使用 `cscript.exe` 測試同一份分段解析邏輯（測試程式僅轉換 VBA/VBScript 語法差異）。加上 `--large-kb 265646` 可建立約 259 MB 的暫存合成檔，透過 Windows FSO 分段讀取測試解析緩衝；此測試不會執行 VBA 二進位讀檔或 Excel 寫入。安裝開發用 `oletools` 後，執行 `python tools/verify_xlsm.py` 可獨立抽取並比對內嵌巨集。一般使用者不需要這些工具。

可使用 NuGet 官方 `DocumentFormat.OpenXml` 2.20.0 套件的 `lib/net46/DocumentFormat.OpenXml.dll`，執行 `tools/validate_office2016.ps1 -SdkAssembly <DLL路徑>`，檢查 Office2016 格式。這是格式驗證，不代表已在 Excel 原生應用程式執行巨集。

`python tools/verify_layout.py` 檢查儲存格排版、固定列高與按鈕連結。`add_import_button.py` 只使用 XlsxWriter 產生 Forms 控制項，不建立或替換原有儲存格資料。

VBA 二進位容器以 [ovba-writer](https://github.com/kay-ws/ovba-writer)（MIT）建置，使用其測試容器的專案參考資料；範例程式與事件已全部移除，僅保留必要文件屬性及本專案模組。建置工具與授權來源見 `tools/THIRD_PARTY_NOTICES.md`。
