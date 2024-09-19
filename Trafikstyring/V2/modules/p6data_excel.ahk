#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; var

undtagneTrKolStart := 11
undtagneTrKolSlut := 20

excel := ComObject("Excel.Application")
excel.Visible := 0
excel_fil := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
workbook_sheet := workbook.Sheets(1)


EndRow := workbook_sheet.usedrange.rows.count
EndCol := workbook_sheet.usedrange.columns.count
usedrangeArr := workbook_sheet.usedrange.value

; MsgBox test[1, 2].value

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
 

MsgBox "test"
