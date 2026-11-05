# 待開發圖表規格（2026-11-05）

本文件記錄此專案的分析圖表規格。16×16 BLK與雙擊電壓分析已實作於VBA；WL/BL/IO三維圖仍僅規劃。

## 16×16 BLK Distribution
- 依2026-11-05後續更正採column-major：每欄由上往下16個BLK，再向右下一欄。
- 128 BLK示例為16列×8欄，第一欄BLK0～15，第二欄BLK16～31，最後一欄BLK112～127。
- 256 BLK為16列×16欄，第一列BLK0、16、32…240，右下BLK255。
- BLK = rowIndex + 16*columnIndex，兩個Index均從0開始。
- 雙擊BLK，保留WT/Chip/BLK條件，顯示Scan/March的Fail Voltage分布。
- 建議明細以Function設定中的實際電壓與FBC呈現，不把Monitor電壓冒充Function電壓。
- 仍沿用Hard Fail排除、Scan先於March、較早WT所有Function需Pass等既有條件。
- Scan Fail的同一BLK，其March應顯示排除/不納入，而不是假造March Pass或Fail曲線。
- 電壓未設定、必要FBC缺失、排除與實際FBC=0分別顯示。
- 實作採格內首次Function失效電壓，底色為已捕捉且可比較的W/R Monitor傾向；模擬PNG的FBC色階已由此規格取代。

## WL/BL/IO三維失效分布
- 選定WT/Chip/Pattern/BLK，三軸為WL0～1023、BL0～15、IO0～15。
- 每個點為一個已捕捉、依完整Address去重的Function Fail Address。
- 可以旋轉/縮放，點選後查看WL/BL/IO、Function TNO與Monitor分類；目前示意為靜態圖。
- 模擬展示WL區間集中、固定BL與固定IO等邏輯特徵，不表示真實物理結構或已確診根因。
- 未記錄Address不能當成Pass；輸出上限1000可能使點雲不完整，保留FBC與已捕捉Address數等覆蓋資訊。
- 三維圖顯示邏輯座標，不當作實體晶片位置；濃密區可另外搭配三個二維投影。

## 後續提案：電壓標籤與W/R底色
- 選定WT/Chip/Pattern，每BLK以該Pattern內依Config測試順序第一筆FBC>0所對應的Function電壓作代表，優先使用完整FBC by BLK，不依賴可能截斷的詳細Address。
- 此標籤表示第一個觀察到失效的測試電壓，不直接等同完整BLK的Vmin。
- 底色希望區分該BLK內Function Fail Bits的Write worse／Read worse傾向。
- 已實作：逐Address比較W/R兩側最高已觀察Fail電壓；較高者worse、相同Equal。混合類型為Mixed，單側或缺電壓為Insufficient；與原本任一電壓有Fail的四分類分開。
- 建議按完整Address比較W/R Monitor的失效電壓，再彙整BLK；混合、雙側相當、Monitor沒有匹配與資料不足需保留區別。
- 詳細Address輸出上限可能使樣本不完整，因此尚不能宣稱已判斷所有實際Fail Bits；不得以缺少記錄直接判Pass。
- 16×16 BLK、首次失效電壓與Observed W/R分析已實作；3D圖仍僅規劃。

實作版本與操作細節見 README_BLK_Voltage.md；不改ASC Parser及Map計數核心。
