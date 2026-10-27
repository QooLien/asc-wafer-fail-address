# ASC 標頭自動命名另存（本機更新）

此更新發布為 v0.7.0 測試版。使用 `templates/ASC_Fail_Address_Fast.xlsm`，VBA 已內嵌。

## 規則

讀取 ASC 中各自第一次出現的 `LOT=`、`WNO=`、`STEP=`、`Program=`，後續同名欄位不覆寫第一次的值。

- LOT：使用 LotID。
- WNO：非負整數，少於兩位補零，例如 3 → 03、10 → 10。
- STEP：去掉開頭的 WTT，例如 WTTAAAA → AAAA。
- Program：在第一個 Program 欄位中尋找第一個 `VER_` 版本，忽略它前面的亂碼。支援英數字與底線版本，輸出前綴統一為 `Ver_`。

範例：`LOT=L12345 WNO=3 STEP=WTTAAAA Program=亂碼VER_FB` → `L12345_#03_AAAA_Ver_FB.xlsm`。

結果存到來源 ASC 所在資料夾。完成視窗顯示完整另存路徑與運行時間，Config!D16 也保存輸出路徑。計時包含另存檔所花的時間。

同名檔已存在時增加 `_001`、`_002` 等流水號，不覆蓋現有檔案。原範本的磁碟檔案保留；Excel 中目前的活頁簿會改為新名稱。

缺少必要欄位、第一次欄位空白/無效、找不到版本或路徑過長時停止自動另存，不自行代入後續欄位。標頭的未完成保留區限制為 16,384 字元；檔名中 Windows/Excel 不允許的字元會轉為底線。若資料夾無寫入權限，Excel SaveAs 錯誤以彈跳視窗顯示；不輸出 Log。

## 驗證

`python tools/test_metadata.py` 使用原始 VBA 的標頭及命名函式經語法轉換後，在 Windows VBScript 引擎測試。已涵蓋第一筆優先、1/7/262,144 字元分段、同一行多欄位、Program 前方 Unicode 亂碼、版本變更、WTT 去除、WNO 補零、缺漏/無效欄位、非法檔名字元及同名檔流水號。

地址解析、900 個 TNO 旗標及批次輸出測試仍通過。內嵌 VBA 經獨立抽取比對，Office2016 格式驗證通過。尚未在桌面版 Excel 實際執行 SaveAs 或使用公司的 ASC 驗證。
