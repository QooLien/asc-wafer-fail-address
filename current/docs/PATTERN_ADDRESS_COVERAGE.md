# Pattern 區間 Address 判定

Missing BLK Address 的 Finding 放寬為 BLK 層級的 Address 是否存在。Map 的 Fail BLK Count 仍使用原有 FBC 與 BLK 篩選，不變更。

## 範圍及判定

- 同一 WT、同一 Chip X/Y、同一 BLK。
- 沿用互動 Map 的 Pattern Function 區間：Config 列順序中，前一個 Stop Sign（含 Hard Fail）之後，到目前 Stop Sign 終點（含）為止。
- 只有 Function TNO；Monitor、其他 Pattern、其他 WT／Chip／BLK 不可互相代替。
- 區間任何 TNO 有 Address → Finding=Complete。這代表 BLK 層級的區間有記錄，不代表終點所有 Fail Bit 已記錄完整。
- 區間完全沒有該 BLK 的 Address → Finding=No Address。
- 相同實體 Address 在多個 TNO 重複 Fail，Pattern Address Count 只計一次。

## 輸出欄位

A:I 延用 WT/TNO/Func./Voltage/Pattern/X/Y/BLK/FBC。

| 欄 | 名稱 | 意義 |
|---|---|---|
| J | Endpoint Address Count | Stop Sign 終點 TNO 的精確 Address 數 |
| K | Endpoint Missing Count | max(終點 FBC - 終點 Address Count, 0) |
| L | Finding | 新的 Complete / No Address |
| M | Pattern Address Count | 區間內去重後 Address 數 |
| N | Coverage Basis | Endpoint Address present / Earlier TNO only / No Address in Pattern |
| O | Endpoint Finding | 原始精確終點 Complete / Partial Address / No Address / Address Count > FBC |

Earlier TNO only 用藍底區別；可能受輸出上限影響，但不能僅憑此標記確定原因。原有互動 Chip 明細仍提供精確 TNO 的 Address 診斷。

## 雙格 Map

左格不變。右格只計數整個 Pattern 區間都沒有 Address 的不同 BLK。區間內有記錄但終點未記錄／未達 FBC，不再計入右格；其精確差額保留在 J/K/O。

## 驗證

Excel 原生 1/2/3 ASC 與原有 Address flags 逐格回歸及 3→1 重跑通過。額外測試：前站有、終點無 → Complete + Earlier TNO only + 終點 No Address；終點部分記錄 → Complete + 終點 Partial Address；相同 Address 跨 TNO 去重；其他 WT/Chip/Pattern、Monitor/Hard Fail 均不誤判為有效範圍；右格與 Finding=No Address 的不同 BLK 數一致。已視覺檢查新增欄位及 Earlier TNO only 區別色；正式 ZIP 及內嵌 VBA 完整性通過。
