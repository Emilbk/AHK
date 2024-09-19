#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; var

; MsgBox test[1, 2].value

excelIndlæsArr(p_excel_fil)
{

undtagneTrKolStart := 11
undtagneTrKolSlut := 20

excel := ComObject("Excel.Application")
excel.Visible := 0
excel_fil := p_excel_fil
workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
workbook_sheet := workbook.Sheets(1)


EndRow := workbook_sheet.usedrange.rows.count
EndCol := workbook_sheet.usedrange.columns.count
usedrangeArr := workbook_sheet.usedrange.value

data := Array()

loop EndRow
{
    row_index := A_Index
    data.Push(Array())
    loop EndCol
    {
        col_index := A_Index
        currentCell := usedrangeArr[row_index, col_index]
        data[row_index].Push(currentCell)
    }

}
excel.quit()
MsgBox "Data indlæst!" 
return data
}
