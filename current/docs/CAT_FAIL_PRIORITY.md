# Cat1/2 的 Fail 優先規則

Cat. 仍固定採 WT1 第一份 Raw Data；僅調整 Map 淨空的優先順序。

1. 以同一 WT / X / Y 檢查：Fail Address 的任何已匯入數值 TNO 旗標為 1，或任一已匯入 FBC by BLK > 0，即視為實際 Fail 證據。
2. 有證據時不套用 Cat1/2 淨空，原 Map 與雙格 Map 保留依原有 Hard Fail、前序 Pattern、前序 WT 條件計算的結果。
3. 沒有上述兩種證據時，Cat1/2 依舊保持空白、不填色。
4. 只有 Address、對應 FBC 缺失時保留 No Data，不推估 Fail BLK Count，不由 Cat1/2 改成 Pass。
5. 零值隱藏及正數色階規則維持；Fail 證據不會取消原本 BLK 排除條件。

Address 證據使用數值 TNO 索引並驗證標題一致，採批次陣列讀取。完整匯入在建立 Map 前保存此次 Address 索引，避免讀取舊索引。

## 驗證

Excel 原生 1/2/3 ASC：Cat1/2 加上真實 Fail 保留計數；FBC > 0 而沒有詳細 Address 仍顯示 Fail 與 Loss BLK 數；只有 Address、沒有 FBC 時原 Map 和雙格 Map 保留 No Data；Cat2 沒有證據時空白。既有所有 Address flags 逐格回歸、原核對流程、3→1 重跑均通過。正式 Excel 開啟及五個按鈕檢查通過；已視覺檢查 Cat2 正值 Map。
