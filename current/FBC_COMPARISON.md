> 最新版依實際匯入1～3個WT輸出，不保留未匯入WT；H2已包含完整分析。參見 ONE_CLICK_IMPORT.md。

# BLK_FBC_Comparison：三 WT 的逐 BLK 原始 FBC 比較

依使用者工作表1版型更新：A=X、B=Y、C=BLK；每個 Chip 依原始 Index 排列 BLK0～255，各占256列。Chip 按X、Y排序。

Row1 分成六個區段：WT1 / Monitor → WT2 / Monitor → WT3 / Monitor → WT1 / Function → WT2 / Function → WT3 / Function。先排列三個 WT 的 Monitor，再排列三個 WT 的 Function，維持各組 Config 設定順序。目前每個 WT 為22個 Monitor＋37個 Function，共59欄；含左側3欄，總共180欄。

Row2 每個 TNO 各一欄，標題沿用 Func._Test Voltage，例如 W-Low pBL_1.20v、W-Low pBL_1.035v；不將同名或不同電壓的 TNO 合併。滑鼠停在標題可看註解中的 WT／Monitor或Function／數字TNO；隱藏 Import Index 也保存精確欄位對照，避免只靠人類可讀標題判斷TNO。

實際 FBC=0 顯示空白，底層保留數值0；缺少該 WT/TNO/Chip 的記錄則灰底留白，底層為空值。Calibri11、水平及垂直置中，關閉預設格線，凍結前3欄及前2列。

此表顯示原始FBC，包括 Hard Fail 及之前 WT 已失敗的 BLK。Soft Fail Map 另外使用已確認的 Hard Fail／前 WT 全 Function 排除條件，Monitor 不參與 Map judge。不要將比較表的所有非零 BLK 不經排除直接當成 Map Count。

舊 FBC by BLK 保留為 very hidden 內部資料表，讓原本 Map 與 Address 核對流程維持一致。新版 H5 匯入 FBC 完畢後會自動建立／更新比較表；H2仍只匯入Address，H8仍只重畫Map，H11進行Map計入BLK的Address數量核對。

正式空白範本已包含新比較表雙層標題。BLK_FBC_Comparison_Uniform_Demo.xlsm 是將使用者提供的結果檔轉換後的範例，供直接查看，不取代乾淨範本。原始提供的工作表1／結果檔保留不變；新範例把參考工作表1更新為正式比較表。

原生Excel16已逐項確認105728個FBC值全部一致，包含BLK0與BLK255、WT2及WT3每個Monitor欄位、灰底缺資料與零值區分、Func/Voltage精度及數字TNO對照。三份ASC完整匯入的Map及Address核對預期仍通過：Scan3/1/1、March1/1/1；Complete3、Partial3、No Address2。

Function標題使用與Fail Address相同的向上90度排列，Monitor也使用相同方向。輸出後自動調整欄寬與標題列高，資料列統一26pt；Config保持既有位置與列高，Wafer Map保留等尺寸Chip格。按鈕簡化為匯入Fail Address、匯入FBC、更新Map、核對BLK Address。

Uniform版：Monitor和Function的所有TNO標題皆由下至上90度排列，與Fail Address一致；版型已直接儲存，開啟即顯示。
