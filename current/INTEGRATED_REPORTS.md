> 最新正式檔為 ASC_BLK_Voltage_Analysis.xlsm；本頁所述功能保留於v0.25.0。

# 最新整合版

使用 ASC_Wafer_Analysis.xlsm，所有 VBA 已內嵌，不需要安裝 Python。

1. Excel 桌面版開啟後啟用巨集，確認 Config 的 Monitor A:C、Function D:F、Stop Sign G。
2. 使用原本匯入按鈕，選擇同LOT/同Wafer的1～3份ASC；程式依STEP排列WT。單獨Address按鈕仍保留。
3. Fail Analysis 整合 Chip統計、BLK Monitor彙整、Address匹配明細，可使用頁首連結移動到各表。
4. Import Info 整合 Raw Data Route、執行訊息、Chip座標，置於最後一頁。
5. W&R Map 及 FailBLK Loss Map 使用新名稱；Scan/March Fail Map、BLK_FBC_Comparison、Missing BLK Address 保留。
6. 雙擊Map的Chip或分格，開啟對應Chip Fail Details；點選Source Row可跳至Fail Address。
7. 四分格顯示匹配數字：左上Read only、右上Write only、左下兩者皆有、右下兩者皆未匹配（白字依原設定）。整Chip全Function Pass才藍色合併。
8. FailBLK Loss Map 左格為Fail BLK數、右格為區間無Address的BLK數；不同數字保持分格。

底層原始統計頁以隱藏工作表保留，供原有VBA使用；可見報表已合併。原本五個按鈕與位置保留。iPad可檢視已產生的結果，執行VBA與雙擊互動需使用桌面版Excel。

Pattern_Analysis_Prompt.md 可直接貼至Chat mode，先討論分析方式。
