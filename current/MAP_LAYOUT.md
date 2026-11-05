# Map版面與標題規則

本次交付：ASC_Wafer_Analysis_Map_Layout.xlsm；已匯入64Chip的Demo為ASC_Wafer_Analysis_Map_Layout_Demo64.xlsm。

- Wafer Map縮放70%、欄寬6、Map區列高22；Map數字、座標與標題加粗。導覽區保留操作按鈕高度，按鈕寬度配合新欄寬，不重疊。
- 四宮格右下No monitor item Fail數量改為RGB(128,128,128)灰字，仍是實際數字；不改任何Fail計數。
- 四宮格標題：Lot原始值先取前6碼，再去掉前2碼，接#兩位Wafer號、BL=0、WT、Pattern、Fail。
- 例：FAX7680T0/WNO11 → X768#11 BL=0 WT1 Scan Fail；March為X768#11 BL=0 WT1 March Fail；WT2/3依實際WT替換。
- BL=0只作標題文字，不篩選BL，統計仍包含BL0～15。
- Fail BLK Count標題：WT1 X768#11 最高電壓~最低電壓 Scan Fail BLK。
- 電壓實際查找該Pattern Function區間（前一Stop Sign/Hard Fail之後至當前終點）內Config F的有效最大/最小值，不把第一站/最後一站當成最高/最低，也不取Monitor電壓。通常顯示1.20v~0.56v；1.035v/0.765v保留必要小數。缺電壓時顯示V?~V?。

## 四宮格藍底

- Cat固定使用WT1。
- Cat1～4，且目前WT/Pattern的Stop Sign終點有完整256個BLK FBC、全部等於0，整Chip合併填RGB(0,176,240)。
- 不再要求其他Pattern/所有Function也全部Pass；此規則只針對四宮格，其他Map維持各自原規則。
- Cat0/未知或Cat5以上不依此條件填藍；缺FBC不當0。
- 若四格已有實際匹配Fail數量，保留數字，不能用藍底覆蓋。

驗證：384個Fail BLK Count與1536個四格數字不變；384組Cat/終點FBC條件直接由模擬ASC獨立核對（157組藍底）；Cat1/2/3/4及Cat0/5邊界；FAX768#11標題；非首尾站最高/最低電壓；缺電壓；數字優先及正式檔保存重開。
