> 最新正式檔為 ASC_BLK_Voltage_Analysis.xlsm；本頁所述功能保留於v0.25.0。

# 合併 Map 與一鍵 PPT

正式範本：ASC_Wafer_Maps.xlsm
已匯入64顆Chip的Demo：ASC_Wafer_Maps_Demo64.xlsm

## 操作

1. 使用 Config 原本的匯入並分析按鈕，可選同一LOT/WNO的1～3份ASC。
2. 開啟 Wafer Maps。第3～5列按鈕分別導向各WT/Pattern的Fail BLK、W&R四宮格、FailBLK Loss二宮格。
3. 雙擊Chip/分格開啟Chip Fail Details；Source Row可回到Fail Address原始列。
4. 點Export PPT。需要已安裝桌面版PowerPoint，採晚期繫結，不需另外設定VBA參考，也不需要Python。
5. PPT自動另存於Excel同一資料夾；已有同名檔會加流水號，不覆蓋舊檔。
6. 修改Config或重新匯入Address/FBC後，使用Refresh Maps更新再匯出。

## PPT版面

依最後補充採單頁寬37公分；上排六張Fail BLK Map，下排六張W&R四宮格，順序為WT1 Scan、WT1 March、WT2 Scan、WT2 March、WT3 Scan、WT3 March。
每張下排圖右下方的獨立空間放縮小的FailBLK Loss二宮格，不覆蓋主圖。全部採點陣圖貼上。
只有WT1或WT1/2時只輸出已有WT的圖；主圖保持比例，整組置中。

W&R四宮格現在依各WT/Pattern分別彙整區間內Function Address，仍沿用Hard Fail排除、Scan→March及WT順序篩選，不使用Address數判斷BLK是否Fail。原始Map來源隱藏保留供VBA重建。

Fail Address維持第1列與A:G凍結、細格線與TNO註解。原Config五個按鈕與位置保留。
iPad可檢視已輸出的Excel/PPT；VBA執行、導覽按鈕與Chip雙擊使用桌面版Excel。
