#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; var

; MsgBox test[1, 2].value
class excelObj extends Class
{

    excel_fil_long := ""
    excel_fil_tekst := ""
    excel_data := []

    kolonne_nummer := Map(
        "Budnummer", 0,
        "Vognløbsnummer", 0,
        "Kørselsaftale", 0,
        "Styresystem", 0,
        "Startzone", 0,
        "Slutzone", 0,
        "Hjemzone", 0,
        "MobilnrChf", 0,
        "Vognløbskategori", 0,
        "Planskema", 0,
        "Statistikgruppe", 0,
        "Undtagne transporttyper", []
    )

    vælgfil()
    {
        valgtExcelFilLong := FileSelect()
        if !valgtExcelFilLong
            return
        SplitPath(valgtExcelFilLong, &valgtExcelFil)
        this.excel_fil_long := valgtExcelFilLong

        this.excel_fil_tekst := "Indlæst excel-fil: " . valgtExcelFil

        return
    }

    indlæsfil()
    {

        ; undtagneTrKolStart := 13
        ; undtagneTrKolSlut := 22
        if !this.excel_fil_long
            throw Error("Ingen fil indlæst!")

        excel := ComObject("Excel.Application")
        excel.Visible := 0
        excel_fil := this.excel_fil_long
        workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
        workbook_sheet := workbook.Sheets(1)


        EndRow := workbook_sheet.usedrange.rows.count
        EndCol := workbook_sheet.usedrange.columns.count
        usedrangeArr := workbook_sheet.usedrange.value


        loop EndRow
        {
            row_index := A_Index
            this.excel_data.Push(Array())
            loop EndCol
            {
                col_index := A_Index
                currentCell := usedrangeArr[row_index, col_index]
                if Type(currentCell) = "Float"
                    currentCell := String(Floor(currentCell))
                this.excel_data[row_index].Push(currentCell)
            }

        }
        excel.quit()
        return
    }

    hentKolonneNummer()
    {
        for kolonneNummerExcel, kolonneNavnExcel in this.excel_data[1]
            for kolonneNavnIntern, kolonneNummerIntern in this.kolonne_nummer
                if kolonneNavnExcel = "Undtagne transporttyper"
                {
                    this.kolonne_nummer["Undtagne transporttyper"].push(kolonneNummerExcel)
                    break
                }
                else if kolonneNavnExcel = kolonneNavnIntern
                {
                    this.kolonne_nummer[kolonneNavnIntern] := kolonneNummerExcel
                    break
                }
        return
    }

    test()
    {
        if (this.excel_data.Length = 0)
            MsgBox "ingen data"
    }
}

test := excelObj()
test.vælgfil()
test.indlæsfil()
test.hentKolonneNummer()

MsgBox test.kolonne_nummer["Vognløbsnummer"]
MsgBox test.kolonne_nummer["Undtagne transporttyper"][3]
; excelIndlæsArr(p_excel_fil)
; {

;     undtagneTrKolStart := 13
;     undtagneTrKolSlut := 22

;     excel := ComObject("Excel.Application")
;     excel.Visible := 0
;     excel_fil := p_excel_fil
;     workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
;     workbook_sheet := workbook.Sheets(1)


;     EndRow := workbook_sheet.usedrange.rows.count
;     EndCol := workbook_sheet.usedrange.columns.count
;     usedrangeArr := workbook_sheet.usedrange.value

;     data := Array()

;     loop EndRow
;     {
;         row_index := A_Index
;         data.Push(Array())
;         loop EndCol
;         {
;             col_index := A_Index
;             currentCell := usedrangeArr[row_index, col_index]
;             if Type(currentCell) = "Float"
;                 currentCell := String(Floor(currentCell))
;             data[row_index].Push(currentCell)
;         }

;     }
;     excel.quit()

;     for index, kolonne in data[1]
;     {
;         if (kolonne = "Budnummer")
;         {
;             DataGUI.kolonneBudnummer := index
;             ; tilvælg knap
;         }

;         if (kolonne = "vognløbsnummer")
;         {
;             DataGUI.kolonnevognløbsnummer := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Kørselsaftale")
;         {
;             DataGUI.kolonneKørselsaftale := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Styresystem")
;         {
;             DataGUI.kolonneStyresystem := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Startzone")
;         {
;             DataGUI.kolonneStartzone := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Slutzone")
;         {
;             DataGUI.kolonneSlutzone := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Hjemzone")
;         {
;             DataGUI.kolonneHjemzone := index
;             ; tilvælg knap
;         }

;         if (kolonne = "MobilnrChf")
;         {
;             DataGUI.kolonneMobilnrChf := index
;             ; tilvælg knap
;         }

;         if (kolonne = "Vognløbskategori")
;         {
;             DataGUI.kolonneVognløbskategori := index
;             VognløbskategoriCheckbox.Enabled := 1
;         }

;         if (kolonne = "Planskema")
;         {
;             DataGUI.kolonnePlanskema := index
;             PlanskemaCheckBox.Enabled := 1
;         }

;         if (kolonne = "Økonomiskema")
;         {
;             DataGUI.kolonneØkonomiskema := index
;             økonomiskemaCheckbox.Enabled := 1
;         }

;         if (kolonne = "Statistikgruppe")
;         {
;             DataGUI.kolonneStatistikgruppe := index
;             ; tilvælg knap
;         }

;         ; hvordan?
;         if (kolonne = "Undtagne transporttyper")
;         {
;             DataGUI.kolonneUndtagneTransporttyper := index
;             ; tilvælg knap
;         }
;     }

;     MsgBox "Data indlæst!"
;     ; DataGUI.opt("-Disabled")
;     ; WinActivate(DataguiNavn)
; }
