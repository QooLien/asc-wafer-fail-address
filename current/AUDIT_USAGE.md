# Fail Map BLK / 詳細 Address 數量核對

本功能只核對實際被 Fail Map 計入的 BLK。讀取既有結果，不重讀 ASC，不改寫 Fail Address 或 FBC。

## 使用

1. Config G 設定 Pattern 終點及 Hard Fail；Function TNO 設定必須與 Map 一致。
2. 按 Config J22。新版空白範本可選取舊結果 Excel，工具唯讀開啟來源並把報表寫在新版工具。
3. 查看 Missing BLK Address，以 Finding 篩選分類。Config Q2 固定顯示 Fail Map；舊檔的 All Function / Stop Sign 不再改變核對範圍。

## 核對範圍

與 Map 共用同一段 BLK 篩選程式，不依工作表目前的 AutoFilter 畫面判斷：

- 只檢查有 Pattern Name 的終點 Function TNO。Hard Fail 標記本身不是 Pattern。
- 排除當前與之前 WT 的 Hard Fail BLK。
- WT2 排除 WT1 任一設定 Function TNO FBC>0 的 BLK。
- WT3 排除 WT1 或 WT2 任一設定 Function TNO FBC>0 的 BLK。
- 必要記錄缺失而使 Map 灰底留白的 Chip/WT 不列入。未出現在 Chip Locations 的位置不列入。
- 僅對通過上述條件、終點 FBC>0 的 BLK 輸出一列；Monitor 不參與。

## 數量與分類

Address Count = Fail Address 中相同 WT / TNO / X / Y / BLK 的1標記列數。
不借用不同 TNO、WT、Chip 或 BLK 的 Address。原始匯入輸出的一列是一個唯一 X/Y/BLK/WL/BL/IO 地址。

| Finding | 規則 | 顏色 |
|---|---|---|
| Complete | Address Count = FBC | 綠 |
| Partial Address | 0 < Address Count < FBC | 橘 |
| No Address | Address Count = 0，且 FBC>0 | 紅 |
| Address Count > FBC | Address Count > FBC，需檢查資料一致性 | 黃 |

Missing Count = max(FBC - Address Count, 0)。例如 FBC=10、Address Count=4，分類 Partial Address，缺6筆。
不足可能來自輸出上限，但差異本身不能證明截斷原因；工具不推算未輸出的 WL/BL/IO。

保留分頁名稱 Missing BLK Address，但現在包含完整及不完整的所有 Map 計入 BLK。Row1 顯示四類數量，各類總和等於核對 BLK 總數。

## 保護與驗證

數字 TNO 欄位索引、舊標題歧義、LOT/WNO/STEP/版本不符、未完成匯入、FBC 解析錯誤、重複記錄等既有保護持續保留。失敗保留上一份報表，來源唯讀。
批次讀取 Address，以每 WT/TNO/Chip 的256個 BLK計數陣列建立索引，批次輸出報表。

原生 Excel16 合成資料已驗證：Map 4/1/1 個 BLK 與核對6列一致、四種分類及缺少筆數、BLK255、超過 Long 的 FBC、精確 WT/TNO/Chip/BLK、跨 WT 排除、Hard Fail、篩選隱藏列、唯讀舊檔、錯誤保留及來源不變。尚未取得公司實際 ASC 進行驗證。
