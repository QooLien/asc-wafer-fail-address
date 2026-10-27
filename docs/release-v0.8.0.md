# v0.8.0 — 命名修正與精簡 Fast 介面

下載附件 **ASC_Fail_Address_Fast.xlsm**，VBA 已內嵌。

- LOT 取前六碼，直接接 # 及補足兩位的 WNO。例如 LOT=FAX7680T0、WNO=7、STEP=WTTAAAA、Program 含 VER_FB，另存為 `FAX768#07_AAAA_Ver_FB.xlsm`。
- 第一個 Program 欄位中找到第一個 VER_，只取緊接的兩個英文字母並轉大寫，略過前方字碼與後方內容。
- 移除 Instructions、Import Log 分頁；Config!C11:D16 清空且執行後不回填。
- 輸出頁改名為 Fail Address。TNO 標題由下往上旋轉 90 度，開啟篩選並縮窄欄寬。
- 地址標題使用深藍底白字；TNO 標題使用金黃底深藍字。標題高 60 pt，資料列維持 26 pt、Calibri 置中。
- 保留無 Log、分段讀取、批次輸出、選檔按鈕及完成視窗運行時間。

已通過命名測試、內嵌 VBA 一致性、分頁與空白範圍、直向標題與篩選，以及 Microsoft Open XML SDK Office2016 格式驗證。尚未在桌面 Excel 實際執行匯入與 SaveAs；請使用實際 ASC 確認結果。
