# Pattern順序與Soft Fail BLK遞延

## 判斷順序

所有判斷都在相同 Chip / BLK 上獨立進行：

1. 先要求 Hard Fail TNO 的 FBC=0，排除大量Defect。Hard Fail由Config G的標記決定，通常是第一個電壓Function最後一個Pattern；不依數字TNO猜測。
2. 同一WT依Config由上至下的Pattern終點順序判斷。排除所有較早Pattern終點FBC>0的BLK，僅讓前面都Pass的BLK進入當前Pattern。Hard Fail標記不列為普通Pattern。
3. WT2仍要求WT1所有設定Function TNO均FBC=0；WT3要求WT1及WT2所有Function均為0。Monitor不參與。
4. 當前Pattern終點FBC>0，該BLK計入當前Map一次，最後加總為per-chip Fail BLK Count。FBC大小不作加權。

目前Config順序：Hard Fail502 → Scan530 → March717。WT1 Scan先通過Hard Fail；WT1 March必須再通過WT1 Scan。WT2 March必須通過WT1全部Function及WT2 Scan。依此類推。更多Pattern亦按Config位置串接，不按TNO數值或Pattern名稱排序。

缺少必要前序FBC記錄不視為0／Pass。Map的0代表沒有符合納入條件的新Fail BLK，不等於整顆Chip原始測試全部Pass。

## Address Count與Map完全獨立

Map只讀FBC。不讀取Address Count，不以1000筆額度或剩餘可輸出量決定Pass/Fail。
例如前面輸出950筆，TNO717 FBC=100而只有50筆Address：只要該BLK通過前序FBC條件，仍計入March Map一次；完整性核對另外顯示50/100、Partial Address、缺50。
核對分頁使用相同Map納入範圍，但Address數量僅用來分類Complete／Partial／No Address／Excess。

## 新增模擬Raw案例

對Chip(10,17)：

|WT|BLK|Scan終點|March終點|Map結果|
|---|---:|---|---:|---|
|WT1|1|Fail|5|只計Scan，不計March|
|WT2|4|Fail|7|只計WT2 Scan，不計WT2 March|
|WT3|5|Fail|9|只計WT3 Scan，不計WT3 March|

Raw保留這些March FBC；比較表原樣顯示，不修改或扣除FBC。Map與其Address核對範圍排除重複Fail，因此既有預期Scan3/1/1、March1/1/1及核對8列維持不變。

原生Excel另驗證：Scan→March→第三個Pattern（第三個TNO較小，證明遵循Config順序）、Hard Fail、跨WT所有Function、BLK0/255、Monitor排除、零Address仍能產生Map、缺少前序記錄、來源FBC不變及相同coverage mask。
1/2/3份單鍵匯入、全部Address的WT/TNO旗標、3份改跑1份、Config框線與最後分頁順序，均通過。原始Address解析器與1/3/9/900-TNO回歸亦通過。
