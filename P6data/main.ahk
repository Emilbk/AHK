#Requires AutoHotkey v2.0
Persistent
; omskriv til bedre organisering
#Include dataExcel.ahk
#Include dataGUI.ahk

excel_fil := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
; DataGUI.excelData := excelIndlæsArr(excel_fil)


DataGUI.Show("AutoSize")

vælgExcelFil()
{
    ; DataGUI.opt("+Disabled")
    ; WinActivate(DataGUINavn)
    valgtExcelFilLong := FileSelect()
    if !valgtExcelFilLong
        return
    SplitPath(valgtExcelFilLong, &valgtExcelFil)
    indlæstExcelFilTekst := "Indlæst excel-fil: " . valgtExcelFil
    overskriftExcelfil.Text := indlæstExcelFilTekst
    DataGUI.excelData := excelIndlæsArr(valgtExcelFilLong)

    ; listview
    dataListview.Delete()
    columnNumber := dataListview.GetCount("Col")
    if columnNumber != 0
        loop columnNumber
            dataListview.DeleteCol(1)
    for i, e in DataGUI.excelData[1]
    {
        dataListview.InsertCol(i, , e)
    }

    for i, e in DataGUI.excelData
        if i > 1
        {
            dataListview.Insert(1, , e*)
            ;dataListview.Insert(i, , DataGUI.excelData[i])
        }
    dataListview.ModifyCol()
    return
}