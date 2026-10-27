# v0.4.1 — 修正 Excel 2016 拒絕開啟的 XLSM 封裝

前版副檔名雖為 `.xlsm`，封裝仍透過 XML 預設內容類型宣告成 `.xlsx` 活頁簿；封裝工具只修改既有 Override，沒有處理缺少 workbook Override 的情況。另外，XML 序列化將 OPC 的 Types / Relationships 根元素改成帶前綴的形式，Microsoft 的 OPC 讀取器會拒絕開啟。

本版修正：

- 明確新增 `/xl/workbook.xml` 的巨集活頁簿內容類型。
- 維持 OPC 內容類型與關聯的預設命名空間，避免根元素序列化成 `ns0:Types` / `ns0:Relationships`。
- 檢查實際生效的內容類型、巨集關聯及 VBA 二進位檔，不再只檢查 ZIP 與 XML 是否完整。
- 新增 Microsoft Open XML SDK 的 Office 2016 格式驗證指令。

驗證結果：修正前的檔案可重現內容類型檢查失敗及 Microsoft OPC 讀取錯誤；修正後被識別為 MacroEnabledWorkbook，Office2016 驗證 0 錯誤。VBA 經獨立抽取與原始碼核對一致。分段讀檔及 Log 程式未改動。

此為官方 SDK 格式驗證結果，尚未在桌面版 Excel 2016 實機開啟或執行巨集。請下載此版重新測試，並保留原本的 ASC 資料。
