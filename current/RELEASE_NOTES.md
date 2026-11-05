# Wafer ASC 分析 v0.26.0

[下載正式 VBA Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout.xlsm) · [下載64顆Chip Demo](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout_Demo64.xlsm) · [下載完整套件](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.26.0/ASC_Wafer_Analysis_Map_Layout_Latest.zip)

- **Map版面**：縮放70%、欄寬6、Map區列高22，字體加粗；四宮格右下數量改RGB(128,128,128)灰字。導覽按鈕配合欄寬縮排，不重疊。
- **四宮格標題**：例如 `X768#11 BL=0 WT1 Scan Fail`；March／WT2／WT3依對應Pattern與WT命名。Lot先取前6碼，再去掉前2碼，Wafer補足兩位。
- **Fail BLK Count標題**：例如 `WT1 X768#11 1.20v~0.56v Scan Fail BLK`。實際查找該Pattern Function區間最高／最低電壓，不只取首尾測試點。缺電壓顯示`V?~V?`。
- **BL=0僅為標題文字**：統計仍包含BL0～15，所有Fail數字維持不變。
- **四宮格藍底**：Cat固定採WT1；Cat1～4且當前WT／Pattern終點完整256個BLK FBC全部為0時，整Chip填RGB(0,176,240)。缺資料不當0、Cat0／5以上不依此條件填藍；已有Fail數字時優先保留。
- 保留v0.25.0的BLK電壓雙擊、Observed W/R分析、Source Row、合併Map與一鍵PPT，以及Fail Address凍結與格線。
- 原生Excel驗證：384個Count與1536個四格數值不變；384組獨立Cat／終點FBC條件、Cat0～5邊界、非首尾站電壓極值、標題、數字優先與儲存重開均通過。

全部VBA已內嵌，執行不需Python；採Excel2016相容VBA。詳見 [Map版面與標題規則](MAP_LAYOUT.md)。

---

以下為歷史版本說明；最新版以上方v0.26.0為準。

# Wafer ASC 分析 v0.25.0

[下載正式 VBA Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis.xlsm) · [下載64顆Chip Demo](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis_Demo64.xlsm) · [下載完整套件（含VBA及模擬ASC）](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_BLK_Voltage_Analysis_Latest.zip)

- **Scan／March Map雙擊Chip → Chip Fail Details**：16×16 BLK圖按欄往下排列，第一欄BLK0～15，第二欄BLK16～31，右下BLK255。
- 格內顯示依Config順序、第一筆Function FBC>0的測試電壓；不是完整BLK Vmin，也不把TNO自行排序。
- 底色顯示已捕捉Address的Observed W/R傾向：Write worse／Read worse／Equal／Mixed／Insufficient。只有雙側有效的Monitor失效電壓才比較；缺紀錄不當Pass。星號表示部分捕捉樣本仍無法比較。
- **再雙擊BLK**：查看Scan／March電壓-FBC圖、各TNO原始數據、逐Address W/R比較及Source Row跳轉。沿用Hard Fail、前WT全部Function及Scan→March排除規則。
- **Wafer Maps合併導覽**：各WT／Pattern的Fail BLK、W&R四宮格及FailBLK Loss二宮格集中同頁；原始來源隱藏保留，原Config五個按鈕保留。
- **一鍵PPT**：寬37公分單頁，上排六張Fail BLK圖，下排六張分Pattern四宮格，各下排圖右下方另放縮小二宮格，不重疊，使用點陣圖。需安裝桌面版PowerPoint。[PPT範例](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.25.0/ASC_Wafer_Maps_Demo64.pptx)
- Fail Analysis整合Chip／BLK／Monitor Address明細；Import Info整合路徑、座標與執行訊息。Fail Address固定第1列及A:G，保留細格線與TNO註解。
- 原生Excel驗證：2,304筆獨立ASC BLK預期值、9組Chip情境、W/R混合及缺資料、真實雙擊事件、Source Row、Scan Pass後March Fail、儲存重開均通過。前期6,809個Address/TNO旗標回歸亦保留紀錄。

全部VBA已內嵌，執行不需Python或手動匯入程式碼；採Excel2016相容VBA，已以本機桌面版Excel驗證。iPad可檢視已輸出檔案；VBA互動需桌面Excel。公司實際ASC／Excel2016環境仍請以正式資料確認。

操作細節：[BLK電壓分析](current/BLK_VOLTAGE_ANALYSIS.md) · [合併Map與PPT](current/MAP_DASHBOARD_PPT.md)。WL／BL／IO三維圖仍屬規劃，未加入本版。

---

以下為歷史版本說明；最新版以上方v0.25.0為準。

# v0.18.0 - Stop Sign endpoint Fail Maps

- Config G1 Stop Sign is red/white; G Pattern Name selects the same-row Function TNO in D/E/F as the cumulative endpoint. Monitor A/B/C never participates in Map Judge.
- Count only endpoint BLKs with FBC>0, never re-add earlier TNOs. Green0 Pass, red1..256 Fail BLK Count, gray No Data, yellow Data Error.
- Raw paired X/Y inventory includes DUT/DC/ISB/no-FBC records and is shared across WT1/2/3. Map geometry is data-driven, with one three-WT sheet per Pattern.
- Add Chip Locations and BLK Fail Count outputs; J19 rebuilds Maps without Raw reimport. J13 remains Address-only; J16 imports FBC and builds Maps.
- Fix native VBA EOF call: explicit Call FinishRecord prevents a line-start label from dropping the last record. Native tests cover both complete and incomplete final records.
- Passed native Excel16 tests for synthetic64/128 Chips, two Patterns, three WTs, endpoint501-only0/3/256 counts, missing/DC-only gray, parser-error yellow, managed Map cleanup and Address preservation. Also passed original Address regression and Office2016 format checks. Company ASC/environment remain untested.
