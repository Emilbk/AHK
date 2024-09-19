#Requires AutoHotkey v2.0
Persistent
; omskriv til bedre organisering
#Include dataExcel.ahk
#Include dataGUI.ahk

excel_fil := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
    DataGUI.excelData := excelIndlæsArr(excel_fil)




DataGUI.Show("AutoSize")

vælgExcelFil()
{
    valgtExcelFilLong := FileSelect()
    SplitPath(valgtExcelFilLong, &valgtExcelFil )
    indlæstExcelFilTekst := "Indlæst excel-fil: " . valgtExcelFil
    overskriftExcelfil.Text := indlæstExcelFilTekst
    DataGUI.excelData := excelIndlæsArr(valgtExcelFilLong)
    return 
}

