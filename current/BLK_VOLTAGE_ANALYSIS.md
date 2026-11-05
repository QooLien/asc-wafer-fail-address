# BLK Voltage Analysis 操作說明

正式範本：ASC_BLK_Voltage_Analysis.xlsm
已匯入Demo：ASC_BLK_Voltage_Analysis_Demo64.xlsm
全部VBA已內嵌，使用者不需Python或另行匯入程式碼。延續Excel2016相容的VBA及原有匯入/PPT功能。

1. 在Wafer Maps的Scan/March Fail BLK圖雙擊一顆Chip。
2. Chip Fail Details上方出現16列×16欄BLK圖：每欄由上往下，第一欄BLK0～15，第二欄BLK16～31，右下BLK255。
3. Fail格內顯示該WT/Chip/Pattern依Config順序第一筆Function FBC>0的電壓。不是將TNO依數值排序，也不是完整BLK Vmin。
4. 再雙擊BLK，頁面下方產生Scan/March的Function電壓/FBC圖、各TNO數據、已捕捉Address的W/R比較與Source Row。
5. 點Back to BLK grid回上方，點Back to Map回Wafer Maps；Source Row導向Fail Address。

## 底色與限制

- Write worse：可比較Address的最高已觀察Write Fail電壓高於Read。
- Read worse：相反。
- Equal：最高已觀察W/R Fail電壓相同。
- Mixed：同BLK有兩種以上可比較類型，包括Equal。
- Insufficient：沒有雙側有效的Fail電壓；單側紀錄、無匹配或缺電壓不能直接當Pass。
- 星號：部分已捕捉Address仍無法比較；所有顏色都只反映已捕捉樣本，不能宣稱所有實際Fail Bits皆已比較。
- 若只有一側有Fail紀錄，目前保守保留Insufficient，不以另一側缺紀錄推論Pass。

Hard Fail、前WT所有Function及當前較早Pattern的排除規則維持不變。Excluded與No Data不當成Pass；被排除的Pattern不畫成失效曲線。缺少較早TNO資料時，不宣稱已找出確定的首次失效電壓。
圖表採原生XY散點：相同電壓可能有多個TNO，不插值或假造單調曲線；旁邊數據表保留Config順序與實際Raw FBC。

原有Chip BLK摘要、Monitor匹配Address及其連結保留。新增分析在Chip Fail Details內，不建立額外分析分頁。WL/BL/IO三維圖仍屬後續規劃，本版沒有加入。

驗證使用桌面版Excel：2304個獨立ASC預期值、W/R/Mixed/缺資料邊界、真實Workbook雙擊事件、Source Row跳轉及正式檔重新開啟。
