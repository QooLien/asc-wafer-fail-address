# Map 互動查詢

開啟 ASC_Wafer_Interactive_Demo.xlsm，允許巨集後，在 Scan Fail Map 雙擊有數字的 Chip（例如 B5）。程式開啟 Chip Fail Details；正式匯入使用 ASC_Wafer_Analysis_Interactive.xlsm。

只顯示所點 Pattern 範圍內的 Function：前一個 Stop Sign 之後至目前 Pattern 終點，且只保留該 Map 已計入的 BLK。Monitor 不參與。範例 Scan 為 503～530；March 為 700～717，仍以 Config 順序為準。

明細列出 WT、Pattern、數字 TNO、Func.、Voltage、X/Y、BLK、FBC、Address Count、Missing Count 與 Finding。點 Show Address 可跳至對應 WL／BL／IO 與原始 Fail Address 列號；點 Back to Map 返回原位置。

Address 必須同時符合 WT／TNO／Chip／BLK，不會借用其他測試的資料。FBC 100、Address 50 顯示 Partial Address、Missing 50；完全無地址顯示 No Address。來源無法核對時標示 Address unavailable，不能當作零。Map 不以 Address Count 判斷。

同一 BLK 可在多個 Function 有 FBC，因此明細列數不等於 Map 的 Fail BLK Count。查詢不重讀 ASC，也不修改原始 FBC 與 Fail Address。

只有有 Fail 數字的 Chip 會觸發雙擊；標題、零值與缺資料位置不觸發。重新匯入或重畫 Map 後，舊明細會標示需更新；再次雙擊即可刷新。Chip Locations 與 Log 仍保留在最後兩頁。

已完成 Windows Excel 原生驗證：雙擊事件、Pattern 範圍、WT/TNO/Chip/BLK 精確匹配、998 筆 Scan Address、完整／不足／完全缺失、FBC100／Address50、內部跳轉與來源資料保持一致。
