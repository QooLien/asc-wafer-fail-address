> 最新版面、灰字與Cat1～4藍底規則見 [Map版面規則](../MAP_LAYOUT.md)。下文保留較早版規格供對照。

# Chip Monitor vs Function Map 固定規格

最新版改為 Fail Bit Match **數量**，不是 R/W/RW/X 文字。

| 位置 | 分類數量 | 填滿 | 字體 |
|---|---|---|---|
| 左上 | Read Fail only | 無 | 黑 |
| 右上 | Write Fail only | RGB255/217/202 | 黑 |
| 左下 | Write&Read Fail | RGB255/102/0 | 黑 |
| 右下 | No monitor item Fail | 無 | RGB255/255/255 |

零值空白；右下數量為白字，儲存格仍有實際數字。每個已匯入WT獨立計數，只納入既有Map規則的Function Fail BLK與區間內可取得的完整Address，任一設定Monitor電壓命中即Fail。同一Address去重，不將未輸出Address當作No monitor item Fail。Scan在前，March只接收Scan Pass的BLK。

確認同WT所有Config Function記錄完整、全部256 BLK FBC=0、沒有Function Address且沒有實際分類時，四格填滿RGB0/176/240。缺資料、有Hard/Function Fail或實際分類均不可誤填藍；Monitor-only Fail不影響Function全Pass。

字體Calibri24、軸20，隱藏工作表格線，只畫Chip/標題/座標/外框。詳細分頁精簡與測試證據見 COMPACT_COUNTS.md。
