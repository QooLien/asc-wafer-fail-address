# 可直接貼到 Chat mode 的分析提示詞

我正在分析半導體 SRAM Wafer 的 ASC Raw Data。請先提出分析方法與值得驗證的假說，這次不要寫 VBA，也不要變更既有判斷規則。請區分資料證據、統計推論與無法從現有輸出判定的部分。

## 資料與層級

- 一次匯入同 LOT、同 WNO 的 1～3 份 ASC，依 STEP 數字由小到大定義 WT1、WT2、WT3；WT 代表不同測試溫度，目前沒有固定溫度數值。
- Chip 以 X/Y 識別；每顆 Chip 有 BLK 0～255。
- 完整 Address 為 X/Y/BLK/WL/BL/IO；WL 0～1023，BL 與 IO 各 0～15。同一 WL、不同 BL 或 IO 不可合併。
- FBC by BLK：每筆 WT/TNO/Chip 記錄有256個數值，依原始順序對應 BLK0～255。FBC>0 表示該 BLK 在該 TNO 有 Fail。Fail BLK Count 是符合篩選規則且 FBC>0 的不同 BLK 數，不是 FBC 的總和。
- Fail Address：各 WT/TNO 有 Fail 的完整 Address 輸出旗標1。Raw Data 有每 Chip 約1000個 Fail Bit 詳細 Address 的輸出上限，因此 FBC 正確但詳細 Address 可能不完整，不能把未找到 Address 視為實際 Pass。

## 設定與測試角色

Config A/B/C = Monitor TNO、Func.、Test Voltage。
Config D/E/F = Function TNO、Func.、Test Voltage。
Config G = Stop Sign；填 Pattern 名稱表示該列 Function TNO 為該 Pattern 的最後測試點。填 Hard Fail 表示硬缺陷排除點。測試及 Pattern 順序依 Config 列序，不自行改成 TNO 數值排序。

Monitor 包括 W-Low pBL 與 R-Low pBL，WT1/2/3 都可能有 Monitor：
- 預設 W-Low pBL 為 TNO350～360，R-Low pBL 為370～380。
- 各自電壓序列為1.20、1.035、0.99、0.90、0.81、0.765、0.72、0.68、0.64、0.60、0.56V。
- Monitor 用來比較 Function Fail Address 的 Write/Read 特性，不參與 Function Map 的 Pass/Fail 篩選。
- 目前 W&R 分類採所有設定 Monitor 電壓的任一 Fail，不限定與 Function 電壓相同。請把依電壓細分當成可另外討論的分析，不可說成現有分類規則。

Function 包括 Scan、March 等 Pattern。例：Hard Fail 終點502、Scan 終點530、March 終點717，實際值以 Config 為準。Function 的 FBC 會累積到後續測試點，因此使用 Stop Sign 終點的 FBC 做 Pattern Map 判斷，避免把同一累積缺陷重複歸因到多個 Pattern。

## Soft Fail Map 的順序篩選，單位是 Chip 的每一個 BLK

1. 先排除當前及所有較早 WT 的 Hard Fail TNO 有 FBC>0 的 BLK；這些 BLK 不納入後續 Soft Fail Map。
2. WT1 Scan：Hard Fail 通過且 Scan 終點 FBC>0 才計入。
3. WT1 March：還必須同一 BLK 在 Scan 終點 FBC=0，March 終點 FBC>0 才計入。同一 Chip 可有某 BLK 歸 Scan、另一個 BLK 歸 March，不是整顆 Chip 一次排除。
4. WT2 任一 Pattern：同一 BLK 在 WT1 所有設定 Function TNO 都必須 FBC=0；當前 WT 還需通過 Hard Fail 及更早 Pattern 終點。
5. WT3 任一 Pattern：同一 BLK 在 WT1、WT2 所有設定 Function TNO 都必須 FBC=0；再套用當前 WT 的 Pattern 順序。
6. 若未取得必要 FBC 記錄，保留資料不足狀態，不把缺失當0，不改用 Address Count 代替 FBC。

例如，同一 Chip 的 BLK5 在 Hard Fail 非零，永遠排除；BLK6 Scan 非零，只歸 Scan；BLK7 Scan為0且March非零，只歸March。WT1任一設定 Function TNO 非零的BLK，也不加入WT2 Map。

## Function Address 與 Monitor 的匹配

先取得 Map 篩選後的 Fail BLK，再收集同 WT/Chip/BLK、該 Pattern 區間內所有設定 Function TNO 的 Fail Address。區間為前一個 Stop Sign（含 Hard Fail）之後，到當前 Pattern 終點為止。同一完整 Address 在此範圍去重。

對每個 Function Address，用同 WT、同完整 Address 與所有設定的 W-Low pBL、R-Low pBL Monitor 旗標匹配，分類互斥：
- Read Fail：任一 Read Monitor Fail，沒有匹配到 Write Monitor Fail。
- Write Fail：任一 Write Monitor Fail，沒有匹配到 Read Monitor Fail。
- Write&Read Fail：兩者皆匹配。
- No monitor item Fail：兩者都沒有匹配到。

最後一類只代表沒有觀察到匹配的 Monitor Fail，不可無條件推論為所有 Monitor 真實 Pass。沒有輸出的 Function Address 不可塞進此類；FBC 的單位也不可直接當成已捕捉到的唯一完整 Address 數。

Scan Fail BLK 不會進入同 WT 的 March 分析，因此不可將同一 BLK 的累積缺陷當成兩個獨立 Pattern Fail。

## Address 覆蓋程度

只看被 Map 納入的 Fail BLK：同 WT/Chip/BLK，在同 Pattern 區間任一 Function TNO 有 Address，即可標示區間覆蓋 Complete。即使終點 TNO 沒有 Address，仍保留「僅較早測試點找到」的診斷資訊。

區間完全找不到才為 No Address。終點記錄不足與終點 FBC 的差額另外保留，不把「區間 Complete」誤解成「所有 Fail Bits 的詳細 Address 都已完整輸出」。不同 TNO 的 FBC 與 Address 數要在對應測試點內比較。

## 目前視覺輸出

- Scan Fail Map、March Fail Map：各 WT、各 Chip 的符合條件 Fail BLK Count；0留白，正數黃到紅色階。
- FailBLK Loss Map：每顆 Chip 左格為該 WT/Pattern Fail BLK Count；右格為區間完全沒有 Address 的不同 Fail BLK 數。右格不是遺失的 Fail Bit 數。
- W&R Map：各 WT 的所有合格 Function Pattern，依完整 Address 去重後輸出匹配數目。左上 Read only、右上 Write only、左下兩者都有、右下兩者皆無匹配。四格是數字，不是R/W等文字。
- W&R 及 FailBLK Loss Map：只有所有設定 Function TNO 的256個 BLK 都有有效 FBC且全為0、也無 Function Fail Address，才整 Chip 藍色；缺資料不能判藍。Monitor 單獨有 Fail 不會否定 Function 全 Pass。
- Cat 採 WT1 第一份 ASC 的 TNO1000/CATLST 第一個1的位置。C1/C2不能蓋過真實的正 FBC 或 Fail Address；有實際Fail仍優先顯示數字。
- Map 可雙擊進 Chip Fail Details，Source Row 可回到 Fail Address 原列。

## 請你討論並提供

1. 在以上篩選造成的樣本族群下，建議哪些 Chip/BLK/Address 層級指標，以及各自的分子、分母。
2. 如何比較 W/R 電壓敏感度、Scan/March 差異、WT溫度差異，同時避免累積FBC、前站淘汰與輸出上限造成的誤判。
3. 如何把「完整可觀測」、「僅早期點有Address」、「完全缺Address」分層，評估 W/R 匹配分析的偏差。
4. 值得追查的具體案例排序與明細欄位，並說明每項建議需要哪些資料。
5. 哪些結果可以由現有資料直接證明，哪些只能提出假說，哪些需要補充實際溫度、測試時序或未截斷Address才能驗證。
6. 給出少量具體例子推導，不先產生程式碼。任何想變更既有規則的建議，請明確標為提案。
