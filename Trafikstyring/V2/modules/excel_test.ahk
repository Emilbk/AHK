excel := ComObject("Excel.Application")
excel.Visible := 0
excel_fil := "C:\Users\Emil\Documents\Ny mappe\AHK\Trafikstyring\V2\modules\Genudbud FG8 - FlexGaranti.xlsx"
workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
workbook_sheet := workbook.Sheets(1)


data := map("vognløb", [], "telefon", [])
EndRange := workbook_sheet.usedrange.rows.count
EndCol := workbook_sheet.columns.count
Firma := excel.Range("1:" EndCol)  ; Get a Range.
firmatest := firma.columns(1).cells
MsgBox(firmatest[1, 1].value)
for CurrentCell in firma.columns(1).cells  ; For each item (cell/range) in 'MyRange'...
{
    if (A_Index >= EndRange)
        break
    data["vognløb"].Push(CurrentCell.value)

}
for CurrentCell in firma.columns(4).cells  ; For each item (cell/range) in 'MyRange'...
{
    if (A_Index >= EndRange)
        break
    data["telefon"].Push(CurrentCell.value)

}



; MsgBox CurrentCell.value


return