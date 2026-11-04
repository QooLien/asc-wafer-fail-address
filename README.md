# Wafer ASC 分析 v0.21.0

[下載最新版完整 VBA Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.21.0/ASC_Wafer_Analysis_Coverage_Map.xlsm)

- 新增 **Fail BLK Address Map**：每個 Chip 左格為原 Map 的 Fail BLK Count，右格為 Address 不足的不同 BLK 數；例如 5 個 Fail BLK，其中 1 個缺少 Address，顯示 **5｜1**。
- 右格包含 No Address 與 Partial Address，每個 BLK 只算一次，不加總缺少的 Fail Bit 筆數。依 WT／Pattern／Chip 分開統計，原地址核對保持 WT／TNO／Chip／BLK 精確對應。
- Fail Address G 欄新增 **Cat.**，TNO 從 H 欄開始；TNO=1000 的 CATLST 第一個 1 的位置即為 Cat.，例如 001000 → C3。G 欄採最後可取得 WT 的分類，各 WT Map 使用各自分類。
- **C1／C2 Chip Map 保持空白、不填色**。Map 字體 24，座標軸 20，Calibri。
- 主匯入自動建立新 Map；核對失敗時清除新 Map，避免顯示過期結果。
- Excel 原生驗證：1～3 份 ASC 的原有 Address 旗標逐格一致；新 Map 全部 24 組測試位置、5｜1、部分缺失、BLK 去重、C1／C2、字級及失敗清除均通過。

完整 VBA 及事件內嵌，Windows Excel 2016 可執行，不需 Python／BAT 或額外匯入程式碼。原有互動 Map 與 Fail Address 功能保留。

以下保留歷史版本說明；最新下載請以上方 v0.21.0 為準。

# Wafer ASC 互動分析 v0.20.0

[下載正式 Excel](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.20.0/ASC_Wafer_Analysis_Interactive.xlsm) · [下載互動示範檔](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.20.0/ASC_Wafer_Interactive_Demo.xlsm) · [下載完整試跑包](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.20.0/ASC_ThreeASC_HomeTry.zip)

所有 VBA 與雙擊事件均內嵌，Windows Excel 2016 可執行，不需要 Python 或另外匯入程式碼。

- Config H2「匯入並分析」一次完成 Fail Address、FBC 比較、Map 與地址數量核對；支援 1～3 份 ASC，只輸出已匯入 WT。
- 雙擊 Map 的 Fail Chip，查看所點 Pattern 範圍內的 Function／BLK，並跳到精確匹配 WT／TNO／Chip／BLK 的 WL／BL／IO。詳見 [互動操作說明](current/INTERACTIVE_MAP.md)。
- Hard Fail 與前序 Pattern／WT 已 Fail 的 BLK 排除；March 必須先通過 Scan。Map 不以 Address Count 判斷。
- 地址核對區分完整、部分不足、完全無地址；來源無法核對時另行標示。
- 黃到紅色階顯示 Fail BLK Count，零值隱藏；Calibri 11、置中，TNO 標題由下至上。Chip Locations 倒數第二、Log 最後。

原生 Excel 已驗證雙擊事件、內部連結、Pattern 範圍、998 筆 Scan Address、FBC100／Address50，以及來源資料保持一致。

以下保留前版功能說明；下載與新功能請以上方 v0.20.0 為準。

# ThreeASC：Fail Address、FBC 與 Stop Sign Fail Map

[下載完整 VBA Excel v0.19.0](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.19.0/ASC_ThreeASC_FBC.xlsm)

所有 VBA 都已內嵌；日常使用不需要 Python、BAT 或另外匯入程式碼。

## Config 與按鈕

| 欄位 | 用途 |
|---|---|
| A/B/C | Monitor TNO／Func.／Test Voltage；保留資料匯入，不參與 Map Judge |
| D/E/F | Function TNO／Func.／Test Voltage |
| G | Stop Sign：填入 Pattern Name，例如 Scan，代表同列 D 欄是此 Pattern 的最後測試點 |

G1 使用紅底白字。G 欄空白表示該 TNO 不建立 Map，但仍正常匯入 FBC。Pattern Name 不可重複；`Hard Fail` 為排除標記，可在多個 Function TNO 使用，不另產生 Map。

- **J13：僅匯入 Fail Address**。保留原有地址／TNO 對應，不解析 FBC。
- **J16：匯入 FBC 並建立 Map**。選取三份 ASC；不重算、不改寫 Address。
- **J22：尋找缺少詳細 Address 的 Function BLK**。依相同 WT／TNO／Chip／BLK 比對 FBC>0 與 Address 的1標記；操作見 [核對說明](current/AUDIT_USAGE.md)。
- **J19：只建立／更新 Map**。使用現有 FBC 與 Chip Locations，修改 G 欄名稱後可直接重建，不重讀 Raw。

若新增 D 欄 TNO，而它未包含在先前的 FBC 匯入中，需按 J16 重新取得該 TNO 資料。A/D 兩組選定 TNO 合併匯入且不可重複。三份 ASC 的完整 LOT/WNO 必須相同，STEP 數字由小到大對應 WT1/WT2/WT3。

## Map 判斷：Soft Fail BLK Count

例如 TNO500 的 Fail bit 已累積至501，而501同列 G 填 Scan：

1. 只取 TNO501 的 FBC by BLK，不再加總或合併500。
2. 先排除當前及之前 WT 的 Hard Fail TNO 中 FBC>0 的 BLK，該 BLK 後續不再統計。
3. WT2 只保留 WT1 **所有 Config D 欄 Function TNO 均 FBC=0** 的 BLK；WT3 要求 WT1、WT2 的所有 Function TNO 均為0。包含沒有 Stop Sign 的 Function TNO；Monitor 不參與。
4. 對剩餘 BLK，統計當前 Pattern 終點 FBC>0 的個數；每個 BLK 只計一次，FBC大小不影響計數。缺少必要記錄不視為0或Pass。
5. 同一個 Stop TNO 分別取 WT1/WT2/WT3 的記錄，產生 **Scan Fail Map**，頁內並列三個 WT。

| 顯示 | 意義 |
|---|---|
| 綠底、0 | 過濾後剩餘 Fail BLK Count 為0；不代表原始 Chip 沒有 Hard Fail |
| 紅底、1～256 | 通過排除條件後，指定終點 FBC>0 的 BLK 數量 |
| 灰底、留白 | 缺少指定終點或必要的 Hard Fail／前 WT Function 記錄，不視為 Pass |
| 黃底、留白 | 該 WT 發生解析／輸出錯誤，不與未測混淆 |

Monitor 不參與 Pass/Fail 判斷。Map 也不做 Vmin 或額外 Cat Fail 推論。

Map 格內只顯示 Count，座標軸只顯示數字；預設格線關閉，保留 Chip、標題及外框框線。FBC by BLK 的0顯示空白，底層數值仍為0。

## 座標及輸出

- 從三份 Raw 中帶有 TNO 或 DUT 的記錄蒐集實際 X/Y 配對並去重，包含 DC/ISB 沒有 FBC 的 DUT。不是把獨立 X/Y 做笛卡兒積。
- Map 不固定套用64或128配置，依當次 Raw 座標排列；X向右、Y向下。三份 Raw 都未出現過的座標無法推知。
- 缺少終點記錄顯示灰底，但不能單憑缺資料確認實際 DC/ISB 原因。
- **Chip Locations** 保存 Raw 座標與各 WT 解析狀態。
- **BLK Fail Count** 保存 Pattern／Stop TNO／Func.／Voltage／WT／X／Y／Fail BLK Count／Status，方便篩選核對。
- 每個 Pattern 自動建立一頁，名稱例如 Scan Fail Map。取消 Pattern 時只清理程式建立的舊 Map；保留其他工作表。

## 既有輸出與解析

Fail Address 保留六欄地址及 WT/TNO 的1標記；Row1 保留 WT，顯示 Func._電壓，例如 W-Low pBL_1.20v、W-Low pBL_1.035v。缺少可用名稱／電壓則保留 TNO。標題150pt，資料26pt。

FBC by BLK 欄位為 WT／TNO／X／Y／BLK0～255，按同 Chip、Config TNO 順序、WT1/2/3 相鄰排列。只解析同筆標頭具有 TNO/X/Y/FBC BY BLK= 的記錄，DUT不參與數值索引。數值支援跨行與一般／Tab／全形／不換行空白。TNO0與未選定TNO略過；選定資料必須恰好256個非負整數。同一 WT/TNO/X/Y 重複時保留第一筆。

**v0.18.0 也修正 EOF 收尾呼叫**：先前行首 FinishRecord: 在原生VBA被解析成標籤，可能漏掉最後一筆。現在使用明確 Call，已測試最後完整記錄及最後不足256筆的錯誤。

## 驗證

已於本機原生 **Excel16.0** 執行合成64／128顆案例：實際解析、排序、兩個Pattern、三WT、Count0/3/256、DC-only Chip灰底、缺終點灰底、解析錯誤黃底、舊Map清理及Address資料保留均通過。原有Address1/3/9/900-TNO回歸、標題/欄位對應及Office2016 OpenXML格式驗證也通過。

v0.19.0 另通過原生 Excel 的逐 BLK Hard Fail、多個排除 TNO、WT1/2 全 Function 過濾（含前面 TNO Fail 而最後 TNO=0）、缺少必要記錄、來源不變，以及詳細 Address 核對測試。規則見 [Soft Fail Map](current/SOFT_FAIL_MAP.md)。

仍未使用公司實際 ASC 驗證，也不代表已在公司的特定 Excel2016 組態測試。iPad可查看輸出或遠端操作Windows Excel；iPad版Excel不執行VBA。

程式：current/src；測試：current/tools；reference只保留地址回歸基準。docs/coordinates.json 為舊128座標及合成測試基準，產品Map使用Raw座標。
