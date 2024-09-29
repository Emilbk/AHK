#Requires AutoHotkey v2.0
#SingleInstance Force
; Objects


nuværendeExcelRække := 0
totalExcelRække := 0
excelRækkeTekst := "Excelrække " nuværendeExcelRække "/" totalExcelRække
nuværendeVognløb := "0"
nuværendeKørselsaftale := "0"
nuværendeStyresystem := "0"
nuværendeKørStyr := nuværendeKørselsaftale "_" nuværendeStyresystem
vognløbTekst := "Vognløb " nuværendeVognløb " - Kørselsaftale " nuværendeKørStyr
indlæstExcelFil := "Ingen fil"
indlæstExcelFilTekst := "Indlæst excel-fil: " indlæstExcelFil

ikkeFuldført := ""
fuldført := "✔️"

kolonneBudnummer := ""

kolonneVognløbsNummer := ""
kolonneKørselsAftale := ""
kolonneStyreSystem := ""

kolonneMobilnrChf := ""
kolonneMobilnrVm := ""

kolonneØkonomiskema := ""
kolonnePlanSkema := ""
kolonneVognløbsKategori := ""
kolonneStatistikGruppe := ""

kolonneHjemzoneAdresse := ""
kolonneHjemzonePlanetZone := ""
kolonneStartzone := ""
kolonneSlutzone := ""

kolonneUndtagneTransportTyper := []


; GUImenu
DataMenu := MenuBar()

DataMenuFil := Menu()
DataMenuFil.Add("Exit", (*) => ExitApp())

DataMenuKategorier := Menu()
DataMenuKategorier.Add("Alle", (*) => ExitApp())
DataMenuKategorier.Add("Skemaer", (*) => ExitApp())
DataMenuKategorier.Add("Vognløbsnotat", (*) => ExitApp())

DatamenuData := Menu()
DatamenuData.Add("Indlæs Excel", (*) => indlæsExcelFil())
DatamenuData.Add("Vis indlæste vognløb", (*) => dataListviewGUI.Show("AutoSize"))

DataMenuHjælp := Menu()
DataMenuHjælp.Add("Hjælp", (*) => ExitApp())

; datoer
DatamenuDato := Menu()

; GUI
DataGUINavn := "P6-Data"
DataGUI := Gui(, DataGUINavn)
; objects
DataGUI.xlObj := ""
DataGUI.vlObjArray := []

; controls
datagui.overskrift := Map()
DataGUI.knap := Map()
dataGUI.checkbox := Map()
dataGUI.editbox := Map()
dataGUI.flueben := Map()

DataGUI.MenuBar := DataMenu
DataMenu.Add("Filer", DataMenuFil)
DataMenu.Add("Kategorier", DataMenuKategorier)
DataMenu.Add("Data", DatamenuData)
DataMenu.Add("Datoer", DatamenuDato)
DataMenu.Add("Om", DataMenuHjælp, "Right")

; GUIListview
dataListviewGUI := Gui(, "Indlæste vognløbsdata")
dataListviewGUI.listviewArray := Array()
dataListview := dataListviewGUI.Add("ListView", "Grid NoSort W1100 R30", dataListviewGUI.listviewArray)


; Pos-udgangspunkt
xUdgangspunkt := 10
yUdgangspunkt := 5

; Pos-Overskrift
overskriftX := xUdgangspunkt
overskriftY := yUdgangspunkt

; Pos-kategorier
planskemaX := xUdgangspunkt
planskemaY := yUdgangspunkt + 75
ØkonomiskemaX := planskemaX
ØkonomiskemaY := planskemaY + 25


; GUIstatus
; kategoriFuldført := 0
; katogoriTotal := 7
; DataStatus := DataGUI.Add("StatusBar", , "Fuldførte kategorier ud valgte kategorier: " kategoriFuldført "/" katogoriTotal)

DataGUI.SetFont("Bold")
; TODO lav fornuftig autoresize ved tekstændring overskrift
DataGUI.overskrift.Excelfil := DataGUI.Add("Text", "Y" overskriftY " W400", indlæstExcelFilTekst)
DataGUI.overskrift.ExcelRækker := DataGUI.Add("Text", "Y" overskriftY + 20 " X" overskriftX, excelRækkeTekst)
DataGUI.overskrift.Vognløb := DataGUI.Add("Text", "Y" overskriftY + 35 " X" overskriftX, vognløbTekst)
DataGUI.SetFont("Norm")

DataGUI.Add("Text", "X" planskemaX " Y" planskemaY - 20, "Skemaer")
dataGUI.checkbox.planskema := DataGUI.Add("Checkbox", "Disabled Section" " X" planskemaX " Y" planskemaY, "Planskema")
DataGUI.Add("Text", " X" planskemaX + 110 " Y" planskemaY - 20, "Forventet")
Datagui.editbox.planskemaForventet := DataGUI.Add("Text", "X" planskemaX + 110 " Y" planskemaY, "AB232")
DataGUI.Add("Text", " X" planskemaX + 160 " Y" planskemaY - 20, "Indlæst")
Datagui.editbox.planskemaIndlæst := DataGUI.Add("Text", "X" planskemaX + 160 " Y" planskemaY, "")
DataGUI.flueben.planskema := DataGUI.Add("Text", " X" planskemaX + 200 " Y" planskemaY, fuldført)

dataGUI.checkbox.Økonomiskema := DataGUI.Add("Checkbox", "Disabled Section" " X" ØkonomiskemaX " Y" ØkonomiskemaY, "Økonomiskema")
Datagui.editbox.ØkonomiskemaForventet := DataGUI.Add("Text", "X" ØkonomiskemaX + 110 " Y" ØkonomiskemaY, "AB232")
Datagui.editbox.ØkonomiskemaIndlæst := DataGUI.Add("Text", "X" ØkonomiskemaX + 160 " Y" ØkonomiskemaY, "")
DataGUI.flueben.Økonomiskema := DataGUI.Add("Text", " X" ØkonomiskemaX + 200 " Y" ØkonomiskemaY, ikkeFuldført)

vognløbskategoriX := xUdgangspunkt
vognløbskategoriY := yUdgangspunkt + 150

DataGUI.Add("Text", "X" vognløbskategoriX " Y" vognløbskategoriY - 20, "Vognløbskategori")
dataGui.checkbox.vognløbskategori := DataGUI.Add("Checkbox", "Disabled Section" " X" vognløbskategoriX " Y" vognløbskategoriY, "Vognløbskategori")
DataGUI.Add("Text", " X" vognløbskategoriX + 110 " Y" vognløbskategoriY - 20, "Forventet")
Datagui.editbox.vognløbskategoriForventet := DataGUI.Add("Text", "X" vognløbskategoriX + 110 " Y" vognløbskategoriY, "FG8")
DataGUI.Add("Text", " X" vognløbskategoriX + 160 " Y" vognløbskategoriY - 20, "Indlæst")
Datagui.editbox.vognløbskategoriIndlæst := DataGUI.Add("Text", "X" vognløbskategoriX + 160 " Y" vognløbskategoriY, "")
DataGUI.flueben.vognløbskategori := DataGUI.Add("Text", " X" vognløbskategoriX + 200 " Y" vognløbskategoriY, fuldført)

; Omskriv?
vognløbsnotatX := xUdgangspunkt + 300
vognløbsnotatY := yUdgangspunkt + 75
vognløbsnotatEditForventetTekst := "GV 8-16, Type 8 adasdlkjadlkjsaldkjasldladasdlj"
vognløbsnotatEditIndlæstTekst := "GV 8-16, Type 8 sdlfsldflkjglrejg reljg dflgkjfd glkdjg lkjd g"


; DataGUI.Add("Text", "X" vognløbsnotatX " Y" vognløbsnotatY -20, "Vognløbsnotat")
Datagui.editbox.vognløbsnotatForventet := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY - 20, "Vognløbsnotat")
dataGUI.checkbox.vognløbsnotat := DataGUI.Add("Checkbox", "Disabled Section" " X" vognløbsnotatX " Y" vognløbsnotatY, "Vognløbsnotat")
Datagui.editbox.vognløbsnotatIndlæst := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY + 25, vognløbsnotatEditIndlæstTekst)
DataGUI.flueben.vognløbsnotat := DataGUI.Add("Text", " X" vognløbsnotatX + 100 " Y" vognløbsnotatY, fuldført)

knapX := xUdgangspunkt + 200
knapY := yUdgangspunkt + 400
DataGUI.knap.sætIgang := DataGUI.Add("Button", "X" knapX " Y" knapY, "Sæt igang")
DataGUI.knap.sætIgang.OnEvent("Click", (*) => testfunk())
; DataGUI.Add("Text", "XP" , "Skema")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; PlanskemaCheckbox := DataGUI.Add("Checkbox", "XS Section", "Planskema")
; DataGUI.Add("Text", , "Skema")
; PlanskemaEditboxTidligere := DataGUI.Add("Edit", EditboxPos , "AB232")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; ØkonomiskemaCheckBox := DataGUI.Add("Checkbox", "YS+10", "Økonomiskema")
; PlanskemaCheckBox := DataGUI.Add("Edit", "X" PlanskemaEditW "" , "AB232")


; DataGUI.Show("AutoSize")
; funk
DataGUIopdater(p_vl_obj)
{

}

indlæsExcelFil()
{
    xl := excelObj()
    xl.indlæsfilFunk()
    DataGUI.xlObj := xl
    ; vl := vlObj()
    ; vl.vlIndhentData(DataGUI.xl.excel_data, 2)
    indlæsListeVl()
    DataGUI.overskrift.Excelfil.text := DataGUI.xlObj.excel_fil_tekst
    DataGUI.overskrift.ExcelRækker.text := "Excelrække: sadasd" (DataGUI.xlObj.excel_data.Length - 1)
    DataGUI.vlObjArray := []
    for index, element in Datagui.xlObj.excel_data
    {
        vl_count := 0
        if (index > 1)
        {
            vl_count += 1
            vl := vlObj()
            vl.IndhentData(element)
            DataGUI.vlObjArray.push(vl)
        }
    }
    tjekGyldigeExcelKolonner()
    opdaterGUI(DataGUI.vlObjArray[1])
    MsgBox "Data indlæst!"
}

indlæsListeVl()
{
    ; listview
    dataListview.Delete()
    columnNumber := dataListview.GetCount("Col")
    if columnNumber != 0
        loop columnNumber
            dataListview.DeleteCol(1)
    for i, e in DataGUI.xlObj.excel_data[1]
    {
        dataListview.InsertCol(i, , e)
    }

    for i, e in DataGUI.xlObj.excel_data
        if i > 1
        {
            ; +1 for at indsætte fra bunden
            dataListview.Insert(DataGUI.xlObj.excel_data.Length + 1, , e*)
            ;dataListview.Insert(i, , DataGUI.excelData[i])
        }
    dataListview.ModifyCol()
    ; lav bedre løsning
    return
}

tjekGyldigeExcelKolonner()
{
    for kolonneNavn, kolonneIndhold in DataGUI.xlObj.kolonne_nummer
    {
        if (DataGUI.xlObj.kolonne_nummer[kolonneNavn])
            try
            {
                if (DataGUI.checkbox.%kolonneNavn%) ; resten af checkbox skal sættes op på xl/vl-kategorier
                    DataGUI.checkbox.%kolonneNavn%.Enabled := 1
            }
    }

}

opdaterGUI(p_vl_obj)
{

    DataGUI.overskrift.Vognløb.text := p_vl_obj.vl_data["Vognløbsnummer"] ", " p_vl_obj.vl_data["Kørselsaftale"] "_" p_vl_obj.vl_data["Styresystem"]

    DataGUI.editbox.ØkonomiskemaForventet.text := p_vl_obj.vl_data["Økonomiskema"]
    DataGUI.editbox.PlanskemaForventet.text := p_vl_obj.vl_data["Planskema"]
    DataGUI.editbox.vognløbskategoriForventet.text := p_vl_obj.vl_data["Vognløbskategori"]
}

testfunk()
{
    for element in DataGUI.vlObjArray
    {
        opdaterGUI(element)
        MsgBox element.vl_data["Vognløbsnummer"]
    }

    return
}