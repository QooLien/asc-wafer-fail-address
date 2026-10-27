# ASC Wafer Fail Address

使用 Excel VBA 從 `.asc` Raw data 讀取指定測試項目（TNO）的 SRAM Fail Address，並保留 128-chip wafer map 與四象限計數範本。

## 快速開始

1. [下載內嵌 VBA 的 ASC_Fail_Address.xlsm](https://github.com/QooLien/asc-wafer-fail-address/raw/refs/heads/main/templates/ASC_Fail_Address.xlsm)，使用 Windows 電腦版 Excel 開啟。
2. 依公司政策允許該檔案的巨集執行；**不需要匯入 `.bas`，也不需要執行 `.bat`**。
3. 在 `Config!A2:A901` 每列輸入一個 100–999 的整數 TNO，預設為 100。
4. 按 Alt+F8 執行 `ImportASCFailAddresses`，選擇一個 `.asc` 檔案。

`src/ImportFailAddresses.bas` 保留供版本管理與檢閱。舊版 `.xlsx` 是沒有巨集的原始範本；日常使用請開啟新的 `.xlsm`。

VBA 必須在 Windows 電腦版 Excel 執行。iPad 可遠端操作該電腦或查看輸出；iPad 版 Excel 不執行 VBA。

## ASC 格式與輸出

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

空白表示輸入檔中未找到對應的有效 Fail 記錄，不能據此推論已測試或 Pass。每次匯入取代舊結果；取消檔案選擇、輸入驗證失敗、沒有指定的 TNO 區段及全部地址無效時保留舊結果。錯誤對話框保留原始錯誤代碼及原因；全部地址無效時附第一筆遭略過的記錄。

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

## 狀態與限制

- 第二版已內嵌修正版 VBA，依使用者提供的格式建立，尚未取得實際 ASC 樣本。
- 使用 Windows VBScript 引擎執行從 VBA 原始碼轉換的解析函式及區段處理：19 種地址案例通過，涵蓋空白字元、範圍上下限、錯誤欄位、多 TNO、重複地址與 1 標記。
- VBA 容器已經獨立的 `olevba` 讀回，核對內嵌原始碼與 `.bas` 完全一致；XML/ZIP、巨集關聯及文件代碼名稱已檢查。
- 開發環境未安裝 Excel；上述測試不代表已完成 Excel 原生編譯、檔案選擇及真實 ASC 匯入驗證。
- ASC 目前整檔讀入記憶體。自動辨識帶 BOM 的 UTF-8 / UTF-16；無 BOM 時使用 Windows 系統文字編碼。
- 尚未驗證 iPad 端下載、開啟及遠端執行流程。

## 開發驗證

Windows 上執行 `python tools/test_parser.py` 可使用 `cscript.exe` 測試同一份解析邏輯（測試程式僅轉換 VBA/VBScript 語法差異）。安裝開發用 `oletools` 後，執行 `python tools/verify_xlsm.py` 可獨立抽取並比對內嵌巨集。一般使用者不需要這些工具。

VBA 二進位容器以 [ovba-writer](https://github.com/kay-ws/ovba-writer)（MIT）建置，使用其測試容器的專案參考資料；範例程式與事件已全部移除，僅保留必要文件屬性及本專案模組。建置工具與授權來源見 `tools/THIRD_PARTY_NOTICES.md`。
