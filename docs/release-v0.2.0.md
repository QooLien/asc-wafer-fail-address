# v0.2.0 — 內嵌 VBA 與空白字元解析修正

- 直接下載並開啟 `ASC_Fail_Address.xlsm`，不需匯入 `.bas`，不需執行 `.bat`。
- 方括號內以逗號分欄，辨識各欄的完整整數，容許一般空格、Tab、換行、不換行空格及全形空白。
- 保留 TNO、X、Y、BLK、WL、BL、IO 對應及既有範圍檢查；同一地址合併，各 TNO 有 Fail 標記 1。
- 自動辨識帶 BOM 的 UTF-8 / UTF-16；無 BOM 時依 Windows 系統文字編碼讀取。
- 錯誤訊息保留錯誤代碼、原因；全部地址無效時列出第一筆遭略過的內容。

使用方式：Windows Excel 開啟檔案，依公司政策允許巨集，在 `Config!A2:A901` 輸入 TNO，按 Alt+F8 執行 `ImportASCFailAddresses`，選擇 ASC。

驗證：19 種地址案例、多 TNO 區段、去重及標記通過；內嵌 VBA 經獨立抽取並與原始碼比對一致。尚未取得實際 ASC，且建置環境無桌面版 Excel，尚未完成 Excel 原生執行驗證。
