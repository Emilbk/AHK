#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; var

; MsgBox test[1, 2].value

excelIndlæsArr(p_excel_fil)
{

undtagneTrKolStart := 13
undtagneTrKolSlut := 22

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
        if Type(currentCell) = "Float"
            currentCell := String(Floor(currentCell))
        data[row_index].Push(currentCell)
    }

}
excel.quit()

for index, kolonne in data[1]
{
    if (kolonne = "Budnummer")
    {
        DataGUI.kolonneBudnummer := index
        ; tilvælg knap
    }

    if (kolonne = "vognløbsnummer")
    {
        DataGUI.kolonnevognløbsnummer := index
        ; tilvælg knap
    }

    if (kolonne = "Kørselsaftale")
    {
        DataGUI.kolonneKørselsaftale := index
        ; tilvælg knap
    }

    if (kolonne = "Styresystem")
    {
        DataGUI.kolonneStyresystem := index
        ; tilvælg knap
    }

    if (kolonne = "Startzone")
    {
        DataGUI.kolonneStartzone := index
        ; tilvælg knap
    }

    if (kolonne = "Slutzone")
    {
        DataGUI.kolonneSlutzone := index
        ; tilvælg knap
    }

    if (kolonne = "Hjemzone")
    {
        DataGUI.kolonneHjemzone := index
        ; tilvælg knap
    }

    if (kolonne = "MobilnrChf")
    {
        DataGUI.kolonneMobilnrChf := index
        ; tilvælg knap
    }

    if (kolonne = "Vognløbskategori")
    {
        DataGUI.kolonneVognløbskategori := index
        VognløbskategoriCheckbox.Enabled := 1
    }

    if (kolonne = "Planskema")
    {
        DataGUI.kolonnePlanskema := index
        PlanskemaCheckBox.Enabled := 1
    }

    if (kolonne = "Økonomiskema")
    {
        DataGUI.kolonneØkonomiskema := index
        økonomiskemaCheckbox.Enabled := 1
    }

    if (kolonne = "Statistikgruppe")
    {
        DataGUI.kolonneStatistikgruppe := index
        ; tilvælg knap
    }

    ; hvordan?
    if (kolonne = "Undtagne transporttyper")
    {
        DataGUI.kolonneUndtagneTransporttyper := index
        ; tilvælg knap
    }
}

MsgBox "Data indlæst!" 
; DataGUI.opt("-Disabled")
; WinActivate(DataguiNavn)
return data
}


