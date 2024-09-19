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
; GUImenu
DataMenu := MenuBar()

DataMenuFil :=  Menu()
DataMenuFil.Add("Exit", (*) => ExitApp())

DataMenuKategorier :=  Menu()
DataMenuKategorier.Add("Alle", (*) => ExitApp())
DataMenuKategorier.Add("Skemaer", (*) => ExitApp())
DataMenuKategorier.Add("Vognløbsnotat", (*) => ExitApp())

DatamenuData := Menu()
DatamenuData.Add("Indlæs Excel", (*) => vælgExcelFil()) 
DatamenuData.Add("Liste indlæste vognløb", (*) => ExitApp())

DataMenuHjælp := Menu()
DataMenuHjælp.Add("Hjælp", (*) => ExitApp())

; GUI
DataGUI := Gui(, "P6-Data")
DataGUI.MenuBar := DataMenu
DataMenu.Add("Filer", DataMenuFil)
DataMenu.Add("Kategorier", DataMenuKategorier)
DataMenu.Add("Data", DatamenuData)
DataMenu.Add("Om", DataMenuHjælp, "Right")

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
kategoriFuldført := 0
katogoriTotal := 7
DataStatus := DataGUI.Add("StatusBar", , "Fuldførte kategorier ud valgte kategorier: " kategoriFuldført "/" katogoriTotal)

DataGUI.SetFont("Bold")
overskriftExcelfil := DataGUI.Add("Text", "Y" overskriftY , indlæstExcelFilTekst)
overskriftExcelRækker := DataGUI.Add("Text", "Y" overskriftY +20 " X" overskriftX , excelRækkeTekst)
overskriftVognløb := DataGUI.Add("Text", "Y" overskriftY +35 " X" overskriftX, vognløbTekst)
DataGUI.SetFont("Norm")

DataGUI.Add("Text", "X" planskemaX " Y" planskemaY -20, "Skemaer")
PlanskemaCheckbox := DataGUI.Add("Checkbox", "Section" " X" planskemaX " Y" planskemaY, "Planskema")
DataGUI.Add("Text", " X" planskemaX +110 " Y" planskemaY -20, "Forventet")
PlanskemaEditboxTidligere := DataGUI.Add("Text", "X" planskemaX + 110 " Y" planskemaY, "AB232")
DataGUI.Add("Text", " X" planskemaX +160 " Y" planskemaY -20, "Indlæst")
PlanskemaEditboxTidligere := DataGUI.Add("Text", "X" planskemaX + 160 " Y" planskemaY, "AB232")
planskemaFuldført := DataGUI.Add("Text", " X" planskemaX +200 " Y" planskemaY, fuldført)

økonomiskemaCheckbox := DataGUI.Add("Checkbox", "Section" " X" økonomiskemaX " Y" økonomiskemaY, "Økonomiskema")
økonomiskemaEditboxTidligere := DataGUI.Add("Text", "X" økonomiskemaX + 110 " Y" økonomiskemaY , "AB232")
økonomiskemaEditboxTidligere := DataGUI.Add("Text", "X" økonomiskemaX + 160 " Y" økonomiskemaY , "AB232")
økonomiskemaFuldført := DataGUI.Add("Text", " X" økonomiskemaX +200 " Y" økonomiskemaY, ikkeFuldført)

vognløbskategoriX := xUdgangspunkt
vognløbskategoriY := yUdgangspunkt + 150

DataGUI.Add("Text", "X" vognløbskategoriX " Y" vognløbskategoriY -20, "Vognløbskategori")
VognløbskategoriCheckbox := DataGUI.Add("Checkbox", "Section" " X" vognløbskategoriX " Y" vognløbskategoriY, "Vognløbskategori")
DataGUI.Add("Text", " X" vognløbskategoriX +110 " Y" vognløbskategoriY -20, "Forventet")
vognløbskategoriEditboxTidligere := DataGUI.Add("Text", "X" vognløbskategoriX + 110 " Y" vognløbskategoriY , "FG8")
DataGUI.Add("Text", " X" vognløbskategoriX +160 " Y" vognløbskategoriY -20, "Indlæst")
vognløbskategoriEditboxTidligere := DataGUI.Add("Text", "X" vognløbskategoriX + 160 " Y" vognløbskategoriY , "FG9")
vognløbskategoriFuldført := DataGUI.Add("Text", " X" vognløbskategoriX +200 " Y" vognløbskategoriY , fuldført)

; Omskriv?
vognløbsnotatX := xUdgangspunkt + 300
vognløbsnotatY := yUdgangspunkt + 75
vognløbsnotatEditForventetTekst := "GV 8-16, Type 8 adasdlkjadlkjsaldkjasldladasdlj"
vognløbsnotatEditIndlæstTekst := "GV 8-16, Type 8 sdlfsldflkjglrejg reljg dflgkjfd glkdjg lkjd g"


; DataGUI.Add("Text", "X" vognløbsnotatX " Y" vognløbsnotatY -20, "Vognløbsnotat")
vognløbsnotatEditboxTidligere := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY -20, "Vognløbsnotat")
vognløbsnotatCheckbox := DataGUI.Add("Checkbox", "Section" " X" vognløbsnotatX " Y" vognløbsnotatY, "Vognløbsnotat")
vognløbsnotatEditboxTidligere := DataGUI.Add("Text", "W200" " X" vognløbsnotatX " Y" vognløbsnotatY +25, vognløbsnotatEditIndlæstTekst)
vognløbsnotatFuldført := DataGUI.Add("Text", " X" vognløbsnotatX +100 " Y" vognløbsnotatY , fuldført)

knapX := xUdgangspunkt + 200
knapY := yUdgangspunkt + 400
knap := DataGUI.Add("Button", "X" knapX " Y" knapY , "Sæt igang")
; DataGUI.Add("Text", "XP" , "Skema")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; PlanskemaCheckbox := DataGUI.Add("Checkbox", "XS Section", "Planskema")
; DataGUI.Add("Text", , "Skema")
; PlanskemaEditboxTidligere := DataGUI.Add("Edit", EditboxPos , "AB232")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; ØkonomiskemaCheckBox := DataGUI.Add("Checkbox", "YS+10", "Økonomiskema")
; PlanskemaCheckBox := DataGUI.Add("Edit", "X" PlanskemaEditW "" , "AB232")










DataGUI.Show("AutoSize")



; funk

vælgExcelFil()
{
    indlæstExcelFil := FileSelect()
    indlæstExcelFilTekst := "Indlæst excel-fil: " . indlæstExcelFil
    overskriftExcelfil.Text := "tekst"
    return
}