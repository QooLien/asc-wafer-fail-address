# v0.7.0 — ASC 標頭自動命名另存測試版

下載附件 **ASC_Fail_Address_Fast.xlsm**，VBA 已內嵌。

- 使用第一次出現的 LOT、WNO、STEP、Program 標頭自動命名。
- WNO 不足兩位補零；STEP 去掉開頭 WTT。
- 例如 LOT=L12345、WNO=3、STEP=WTTAAAA、Program=亂碼VER_FB，另存為 `L12345_#03_AAAA_Ver_FB.xlsm`。
- 結果存於 ASC 所在資料夾；同名檔增加流水號，完成視窗顯示路徑與運行時間。
- 保留無 Log、分段讀取及批次輸出。

此版供確認資料格式。目前版本擷取仍允許 VER_ 後的英數字及底線；「只取第一個 VER_ 後兩個英文字母」的最新修正尚未納入。

已通過標頭命名測試、內嵌 VBA 比對及 Office2016 檔案格式驗證。尚未在桌面 Excel 實際執行 SaveAs 或使用公司 ASC 驗證。
