#Requires AutoHotkey v2.0
#SingleInstance Force

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

kolonneØkonomiSkema := ""
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
DatamenuData.Add("Indlæs Excel", (*) => vælgExcelFil())
DatamenuData.Add("Liste indlæste vognløb", (*) => dataListviewGUI.Show("AutoSize"))

DataMenuHjælp := Menu()
DataMenuHjælp.Add("Hjælp", (*) => ExitApp())

; datoer
DatamenuDato := Menu()

; GUI
DataGUINavn := "P6-Data"
DataGUI := Gui(, DataGUINavn)
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
økonomiskemaX := planskemaX
økonomiskemaY := planskemaY + 25


; GUIstatus
; kategoriFuldført := 0
; katogoriTotal := 7
; DataStatus := DataGUI.Add("StatusBar", , "Fuldførte kategorier ud valgte kategorier: " kategoriFuldført "/" katogoriTotal)

DataGUI.SetFont("Bold")
; TODO lav fornuftig autoresize ved tekstændring overskrift
overskriftExcelfil := DataGUI.Add("Text", "Y" overskriftY " W400", indlæstExcelFilTekst)
overskriftExcelRækker := DataGUI.Add("Text", "Y" overskriftY + 20 " X" overskriftX, excelRækkeTekst)
overskriftVognløb := DataGUI.Add("Text", "Y" overskriftY + 35 " X" overskriftX, vognløbTekst)
DataGUI.SetFont("Norm")

DataGUI.Add("Text", "X" planskemaX " Y" planskemaY - 20, "Skemaer")
PlanskemaCheckbox := DataGUI.Add("Checkbox", "Disabled Section" " X" planskemaX " Y" planskemaY, "Planskema")
DataGUI.Add("Text", " X" planskemaX + 110 " Y" planskemaY - 20, "Forventet")
PlanskemaEditboxForventet := DataGUI.Add("Text", "X" planskemaX + 110 " Y" planskemaY, "AB232")
DataGUI.Add("Text", " X" planskemaX + 160 " Y" planskemaY - 20, "Indlæst")
PlanskemaEditboxIndlæst := DataGUI.Add("Text", "X" planskemaX + 160 " Y" planskemaY, "")
planskemaFuldført := DataGUI.Add("Text", " X" planskemaX + 200 " Y" planskemaY, fuldført)

økonomiskemaCheckbox := DataGUI.Add("Checkbox", "Disabled Section" " X" økonomiskemaX " Y" økonomiskemaY, "Økonomiskema")
økonomiskemaEditboxForventet := DataGUI.Add("Text", "X" økonomiskemaX + 110 " Y" økonomiskemaY, "AB232")
økonomiskemaEditboxIndlæst := DataGUI.Add("Text", "X" økonomiskemaX + 160 " Y" økonomiskemaY, "")
økonomiskemaFuldført := DataGUI.Add("Text", " X" økonomiskemaX + 200 " Y" økonomiskemaY, ikkeFuldført)

vognløbskategoriX := xUdgangspunkt
vognløbskategoriY := yUdgangspunkt + 150

DataGUI.Add("Text", "X" vognløbskategoriX " Y" vognløbskategoriY - 20, "Vognløbskategori")
VognløbskategoriCheckbox := DataGUI.Add("Checkbox", "Disabled Section" " X" vognløbskategoriX " Y" vognløbskategoriY, "Vognløbskategori")
DataGUI.Add("Text", " X" vognløbskategoriX + 110 " Y" vognløbskategoriY - 20, "Forventet")
vognløbskategoriEditboxForventet := DataGUI.Add("Text", "X" vognløbskategoriX + 110 " Y" vognløbskategoriY, "FG8")
DataGUI.Add("Text", " X" vognløbskategoriX + 160 " Y" vognløbskategoriY - 20, "Indlæst")
vognløbskategoriEditboxIndlæst := DataGUI.Add("Text", "X" vognløbskategoriX + 160 " Y" vognløbskategoriY, "")
vognløbskategoriFuldført := DataGUI.Add("Text", " X" vognløbskategoriX + 200 " Y" vognløbskategoriY, fuldført)

; Omskriv?
vognløbsnotatX := xUdgangspunkt + 300
vognløbsnotatY := yUdgangspunkt + 75
vognløbsnotatEditForventetTekst := "GV 8-16, Type 8 adasdlkjadlkjsaldkjasldladasdlj"
vognløbsnotatEditIndlæstTekst := "GV 8-16, Type 8 sdlfsldflkjglrejg reljg dflgkjfd glkdjg lkjd g"


; DataGUI.Add("Text", "X" vognløbsnotatX " Y" vognløbsnotatY -20, "Vognløbsnotat")
vognløbsnotatEditboxForventet := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY - 20, "Vognløbsnotat")
vognløbsnotatCheckbox := DataGUI.Add("Checkbox", "Disabled Section" " X" vognløbsnotatX " Y" vognløbsnotatY, "Vognløbsnotat")
vognløbsnotatEditboxIndlæst := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY + 25, vognløbsnotatEditIndlæstTekst)
vognløbsnotatFuldført := DataGUI.Add("Text", " X" vognløbsnotatX + 100 " Y" vognløbsnotatY, fuldført)

knapX := xUdgangspunkt + 200
knapY := yUdgangspunkt + 400
knap := DataGUI.Add("Button", "X" knapX " Y" knapY, "Sæt igang")
knap.OnEvent("Click", (*) => testfunk())
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