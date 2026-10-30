# Import Three ASC + BLK Vmin Judge（本機整合版）

範本：`templates/ASC_ThreeASC_BLK_Vmin.xlsm`。VBA 已內嵌，Windows Excel 2016 可用；不需要 Python。GitHub 發布版本 v0.13.0。

## 這一版的範圍

保留原 Import Three ASC，新增每個 WT 的 BLK Distribution / Vmin Judge。活頁簿只有：

1. Config：共用設定與原始三檔對應資訊。
2. Fail Address：原本三份地址聯集與 WT1/WT2/WT3 TNO 旗標。
3. WT1_BLK_Distribution。
4. WT2_BLK_Distribution。
5. WT3_BLK_Distribution。

不包含 Split、WT Sum、WT1/2/3 Fail 分類或 Conflict 功能。歷史版本檔案未刪除；新範本完全不載入這些模組。

## 操作與設定

- Config!A2:A901：三位數 TNO（100–999），B2:B901：對應 Voltage。地址與 BLK 共用同一組 TNO。依輸入順序判定 Vmin，不按 TNO 數值排序。重複 TNO 或缺失 Voltage 會先提示錯誤。
- Config!N2：VminFBCThreshold。空白/0 = Legacy；正整數 N = 累積 FBC ≥ N。負數、小數、非數字拒絕。
- Config!N6:O...：可選 DUT X/Y 清單，保留第一次出現順序；留白時，各 WT 依該檔案的有效選定 FBC 資料自動發現 DUT，再按 X、Y 數值排序。
- 按 Config!G13 的「匯入 3 份 ASC + BLK Vmin」。必須正好選取三份，同一完整 LOT / WNO；依 STEP 唯一數字序列由小到大分配 WT1、WT2、WT3。
- Config!G2:L4 保留 WT 來源、STEP、版本及地址記錄統計。Config!Q2 顯示 BLK 缺失/錯誤訊息，不與可選 DUT 清單混用。
- 完成後另存為例如 `FAX768#07_WT1-WT3_BLK_Vmin.xlsm`，存於 STEP 最小的 ASC 所在資料夾；同名加流水號，不覆蓋既有檔案。

範本預設 103/0.60、101/0.65、104/0.70、102/0.75，刻意不依 TNO 排序，用於附帶測試。實際使用前請修改。

## ASC 輸入與兩項功能的關係

Fail Address 沿用原本的 `tno=...` 與 `[2,TNO,X,Y,BLK,WL,BL,IO,]` 記錄，地址有效範圍及位元索引不變。

BLK Vmin 另外辨識：

```text
Tno=103 X=8 Y=12
FBC BY BLK=0 1 0 ... 共256個非負整數
```

從與地址解析相同的解碼後區塊分流，不額外重讀三份完整大檔。每 WT 開始會重設 FBC 資料、TNO/X/Y、暫存及錯誤狀態，避免跨檔混用。

FBC 可跨行，分隔符支援空白、Tab、逗號、分號；不支援帶 index label、方括號、十六進位或科學記號。需有明確 Tno/X/Y 標頭；不猜測其他機台格式。每筆必須 256 個有效非負整數，同 TNO/X/Y 保留第一筆。FBC 與 Threshold 上限 9,007,199,254,740,991；檔案讀取上限 2 GB，行長上限 1 MB。帶 BOM 支援 UTF8/UTF16，無 BOM 使用 Windows ANSI。

**若沒有 FBC 或 FBC 格式不合法，不從 Fail Address 推算或捏造 FBC**：該 WT BLK 頁標示 No Data 或錯誤；原 Fail Address 解析繼續並保留結果。錯誤 FBC 不輸出部分、不確定的 Vmin。

## BLK 欄位與 Judge

各 WT BLK 頁：A=TNO/Vmin Label、B=Voltage、C=X、D=Y、E=BLK0，直到 IZ=BLK255。每個 DUT 依設定 TNO 輸出，再增加一列 A:D 合併的藍底白字 Vmin。公式始終為 Excel 欄號 = BLK Index + 5。

- Legacy：首次 FBC≠0，第一個 TNO 觸發為 Fail，否則輸出前一個 Voltage；全部有效記錄為0才是 Pass。
- Threshold：同 DUT/BLK 依序累積，首次累積≥N，同樣取前一 Voltage；第一筆觸發為 Fail，未達為 Pass。
- 不同 BLK、DUT、WT 的累積互相獨立。
- 缺失測試列顯示 Not Found；若缺失影響首次觸發判定，Vmin 顯示 Missing Data；整個 DUT 無有效資料為 No Data。已在較早記錄確定的首次觸發不被後續缺失推翻。

## 測試與限制

1. `test_three_regression.py`：原地址索引、分組標題、批次輸出，1/3/9/900 TNO 逐格回歸。
2. `test_blk_regression.py`：256 BLK 全映射、Legacy/Threshold1/5/100、Missing/No Data、重複鍵及跨檔隔離，與獨立預期值逐格比較。
3. `test_shared_feed.py`：實際共用串流、混合 FBC/地址記錄、三個 WT 不混用、BLK 錯誤隔離；另驗證提供的三份範例。
4. 範本獨立抽取 VBA 比對、驗證無 Split 模組及五分頁、檢查每頁 E:IZ 所有標題，並經 Office2016 格式驗證。

`examples/three_asc` 提供可直接選取的三份合成 ASC，完整 LOT/WNO 相同，STEP=2/30/100；預期結果見 expected.json。資料共9個唯一地址，各檔地址筆數與 FBC 合計一致。

尚未以實際機台 ASC 在原生 Excel 2016 編譯、執行、合併儲存格及另存；VBScript 測試不代表原生 Excel 端到端驗證。無法宣稱已符合未提供的機台 FBC 變體格式。

兩個正式模組為 `src/ImportThreeASC_BLKVmin.bas` 與 `src/BLKVminEngine.bas`，可在獨立活頁簿中匯入。建議直接使用範本，避免缺少分頁或同名模組。`reference` 僅供回歸比對，請勿匯入。
