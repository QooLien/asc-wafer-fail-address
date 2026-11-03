# ThreeASC：Fail Address、FBC 與 Stop Sign Fail Map

[下載完整 VBA Excel v0.18.0](https://github.com/QooLien/asc-wafer-fail-address/releases/download/v0.18.0/ASC_ThreeASC_FBC.xlsm)

所有 VBA 都已內嵌；日常使用不需要 Python、BAT 或另外匯入程式碼。

## Config 與按鈕

| 欄位 | 用途 |
|---|---|
| A/B/C | Monitor TNO／Func.／Test Voltage；保留資料匯入，不參與 Map Judge |
| D/E/F | Function TNO／Func.／Test Voltage |
| G | Stop Sign：填入 Pattern Name，例如 Scan，代表同列 D 欄是此 Pattern 的最後測試點 |

G1 使用紅底白字。G 欄空白表示該 TNO 不建立 Map，但仍正常匯入 FBC。Pattern Name 不可重複。

- **J13：僅匯入 Fail Address**。保留原有地址／TNO 對應，不解析 FBC。
- **J16：匯入 FBC 並建立 Map**。選取三份 ASC；不重算、不改寫 Address。
- **J19：只建立／更新 Map**。使用現有 FBC 與 Chip Locations，修改 G 欄名稱後可直接重建，不重讀 Raw。

若新增 D 欄 TNO，而它未包含在先前的 FBC 匯入中，需按 J16 重新取得該 TNO 資料。A/D 兩組選定 TNO 合併匯入且不可重複。三份 ASC 的完整 LOT/WNO 必須相同，STEP 數字由小到大對應 WT1/WT2/WT3。

## Map 判斷：只讀累積終點

例如 TNO500 的 Fail bit 已累積至501，而501同列 G 填 Scan：

1. 只取 TNO501 的 FBC by BLK，不再加總或合併500。
2. 對每個 Chip 的256個 BLK，統計 FBC>0 的個數；每個 BLK 只計一次，FBC大小不影響計數。
3. 同一個 Stop TNO 分別取 WT1/WT2/WT3 的記錄，產生 **Scan Fail Map**，頁內並列三個 WT。

| 顯示 | 意義 |
|---|---|
| 綠底、0 | 指定最後測試點的256個FBC全為0，Pass |
| 紅底、1～256 | 指定最後測試點中 FBC>0 的 BLK 數量 |
| 灰底、No Data | 此 Chip/WT 缺少指定 Stop TNO 的 FBC，不視為 Pass |
| 黃底、Data Error | 該 WT 發生解析／輸出錯誤，不與未測混淆 |

Monitor 不參與 Pass/Fail 判斷。Map 也不做 Vmin 或額外 Cat Fail 推論。

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

仍未使用公司實際 ASC 驗證，也不代表已在公司的特定 Excel2016 組態測試。iPad可查看輸出或遠端操作Windows Excel；iPad版Excel不執行VBA。

程式：current/src；測試：current/tools；reference只保留地址回歸基準。docs/coordinates.json 為舊128座標及合成測試基準，產品Map使用Raw座標。
