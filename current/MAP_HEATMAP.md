# Fail Map顯示：零值留白、黃→紅原生色階

- 數字0保留在儲存格中，使用數字格式隱藏顯示；零值條件格式使用白底白字，視覺留白。Excel字體沒有真正透明色，此方式達成相同顯示目的。
- 正值使用 Excel 2016 原生2色色階（FormatConditions.AddColorScale），黃色FFEB84到紅色F8696B，數字越大越紅。
- 同一Pattern的各WT共用範圍：1為黃色、當次所有WT的最大Fail BLK Count為紅色；最大值不足2時用2作上限，以確保全零／全1情境有效。
- 色階只套用Chip數值格，不套用座標、標題或外框。缺資料灰底、解析錯誤黃色留白保留；零值規則用ISNUMBER排除真正空值，避免把灰底缺資料塗成白色。
- 只改視覺，不修改FBC、Map納入條件、Fail BLK Count或Address核對邏輯。

已修正VBA變數名稱：使用colorScaleRule，避免scale保留字造成編譯錯誤。
原生Excel16驗證通過：0仍為數字且顯示空白、白字白底，1黃／3紅、2中間色可隨數值動態插值、兩個原生ColorScaleCriteria、座標不著色、灰底缺資料保留、全零有效，以及全部Map數值前後一致。
