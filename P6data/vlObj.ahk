#Requires AutoHotkey v2.0

class vlObj extends Class
{
    vl_data := Map(
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
        "Økonomiskema", 0,
        "Statistikgruppe", 0,
        "Undtagne transporttyper", []
    )


    ; fjern kobling til datagui
    IndhentData(p_data_array)
    {

        undtagneTransportTyperStart := DataGUI.xlObj.kolonne_nummer["Undtagne transporttyper"][1]
        undtagneTransportTyperSlut := undtagneTransportTyperStart + DataGUI.xlObj.kolonne_nummer["Undtagne transporttyper"].Length -1

        for kolonneNummerExcel, kolonneIndholdExcel in p_data_array
            for kolonneNavn, kolonneNummer in DataGUI.xlObj.kolonne_nummer
            {
                if kolonneNummerExcel >= undtagneTransportTyperStart and kolonneNummerExcel <= undtagneTransportTyperSlut
                {
                    this.vl_data["Undtagne transporttyper"].Push(kolonneIndholdExcel)
                break
                }
                if kolonneNummerExcel = kolonneNummer
                {
                    this.vl_data[kolonneNavn] := kolonneIndholdExcel
                    break
                }
            }

        return
    }
    ; DataGUI.totalExcelRække := p_data_array.Length - 1
    ; DataGUI.nuværendeExcelRække := p_række_nummer - 1
    ; DataGUI.excelRækkeTekst := "Excelrække " DataGUI.nuværendeExcelRække "/" DataGUI.totalExcelRække
    ; overskriftExcelRækker.Text := DataGUI.excelRækkeTekst

    ; PlanskemaEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonnePlanSkema]
    ; økonomiskemaEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonneØkonomiSkema]
    ; vognløbskategoriEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonneVognløbsKategori]

    ; overskriftVognløb.text := "Vognløb " this.vlVognløbsNummer ", " this.vlKørselsAftale "_" this.vlStyreSystem
}
; ???
; p6_indhent_data()
; {

; }
; p6IndlæsData()
; {
;     MsgBox this.vlBudnummer " - " this.vlVognløbsNummer
;     MsgBox this.vlVognløbsKategori

; }
