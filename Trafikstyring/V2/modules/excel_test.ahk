excel := ComObject("Excel.Application")
excel.Visible := 0
excel_fil := "C:\Users\Emil\Documents\Ny mappe\AHK\Trafikstyring\V2\modules\Genudbud FG8 - FlexGaranti.xlsx"
workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
workbook_sheet := workbook.Sheets(1)


columns := array()
EndRange := workbook_sheet.usedrange.rows.count
EndCol := workbook_sheet.usedrange.columns.count
firma := workbook_sheet.usedrange.cells
firma1 := firma.columns(2).cells
firma2 := firma.columns(3).cells
firma3 := firma.columns(4).cells
; Firma := excel.Range("1:" EndCol)  ; Get a Range.
firmatest := firma.columns(1).cells
; MsgBox(firmatest[1, 1].value)
; find column
loop EndCol
    if (firma[1, a_index].value = "Firma" or firma[1, a_index].value = "Navn")
        columns.push(map("Index", a_index, "Kollone", firma[1, a_index].value))
data := Map()
for columns in columns
    data.set(columns["Kollone"], [])
for currentcolumn in columns
    for CurrentCell in firma.columns(currentcolumn["Index"]).cells  ; For each item (cell/range) in 'MyRange'...
    {
        if (A_Index >= EndRange)
            break
        data["Navn"].Push(CurrentCell.value)

    }

for CurrentCell in firma2.cells  ; For each item (cell/range) in 'MyRange'...
{
    if (A_Index >= EndRange)
        break
    data["telefon"].Push(CurrentCell.value)

}


; MsgBox CurrentCell.value


return