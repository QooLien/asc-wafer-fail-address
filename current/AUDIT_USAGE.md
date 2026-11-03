# Function FBC / 詳細 Address 的缺失 BLK 比對

本功能只讀既有結果，不重讀 ASC、不重新計算或改寫 Fail Address / FBC。

## 使用

1. Config!Q2 選 **All Function**（D欄全部 Function TNO）或 **Stop Sign**（只檢查 G欄非空白的 Function TNO）。Monitor 不參與。
2. 按 **Config!J22**。
3. 如果新版範本尚無資料，選取之前已輸出的 Excel 結果檔，檔內需同時包含 Config、Fail Address、FBC by BLK。工具會唯讀開啟來源，把報表寫在新版工具，不必再匯入 Address。
4. 查看 **Missing BLK Address**，依 WT、TNO、X/Y、BLK 篩選。

## 精確條件

對每筆 Function FBC，逐 BLK 檢查：

`FBC > 0` 且 `Fail Address 中相同 WT + TNO + X + Y + BLK 的1標記筆數 = 0`。

不同 TNO 或不同 WT 的 Address 不算相符。有任一筆相符 Address 的 BLK 不列入；本功能不推算缺少的 WL/BL/IO。

報表列出 WT、TNO、Func.、電壓、Pattern、X、Y、BLK、FBC，以及該 WT/TNO/Chip 已記錄的 Address 數。記錄數>=1000時標示 Yes，作為可能達輸出上限的佐證；未達1000時標示 Not confirmed。差異本身不證明截斷原因。

## 正確性保護

- 新匯入結果會保存數字 TNO/欄位索引與 LOT/WNO/STEP/版本，避免只靠顯示標題猜測。
- 舊結果可用唯一的 WT+TNO 或 Func./Voltage 標題辨識；若重複標題造成 TNO 不明確，工具拒絕比對。此時需用新版匯入一次 Address 建立索引，之後不必重複匯入。
- 已記錄的來源識別不一致、未完成的匯入、FBC解析錯誤、無效數字或重複 FBC 記錄，都會停止比對並保留前一份報表。
- 來源沒有舊版識別資料時，報表會提示來源未驗證，使用者需確認兩個輸出來自同一批 Raw。
- 檢查完整資料，不受目前畫面的 AutoFilter 影響。
- 批次讀取、256-BLK位元旗標索引與分塊輸出，避免逐格搜尋及大量 BLK 字串索引。

已使用原生 Excel16 合成測試驗證：1000筆Address、精確WT/TNO/Chip/BLK、BLK255、大FBC、Monitor排除、篩選隱藏列、單列與Stop範圍、數字/舊版標題索引、來源不符/未完成/解析錯誤、失敗保留報表及唯讀舊檔。尚未以公司實際資料驗證。
