#Requires AutoHotkey v2.0
#SingleInstance Force

nuværendeExcelRække := 1
totalExcelRække := 1
excelRækkeTekst := "Excelrække " nuværendeExcelRække "/" totalExcelRække
nuværendeVognløb := 31200
nuværendeKørselsaftale := "3100_47"
vognløbTekst := "Vognløb " nuværendeVognløb " - " nuværendeKørselsaftale

ikkeFuldført := ""
fuldført := "✔️"
; GUImenu
DataMenu := MenuBar()

DataMenuKategorier :=  Menu()
DataMenuKategorier.Add("Alle", (*) => ExitApp())
DataMenuKategorier.Add("Skemaer", (*) => ExitApp())
DataMenuKategorier.Add("Vognløbsnotat", (*) => ExitApp())

DataMenuFil :=  Menu()
DataMenuFil.Add("Exit", (*) => ExitApp())

DataMenuHjælp := Menu()
DataMenuHjælp.Add("Hjælp", (*) => ExitApp())
; GUI
DataGUI := Gui(, "P6-Data")
DataGUI.MenuBar := DataMenu
DataMenu.Add("Filer", DataMenuFil)
DataMenu.Add("Kategorier", DataMenuKategorier)
DataMenu.Add("Om", DataMenuHjælp, "Right")
yUdgangspunkt := 20
xUdgangspunkt := 10

DataGUI.Add("Text", "Y" yUdgangspunkt , excelRækkeTekst)

planskemaX := xUdgangspunkt
planskemaY := yUdgangspunkt + 55
økonomiskemaX := planskemaX
økonomiskemaY := planskemaY + 25


DataGUI.Add("Text", "Y" yUdgangspunkt " X" xUdgangspunkt + 100, vognløbTekst)
DataGUI.Add("Text", "X" planskemaX " Y" planskemaY -20, "Skemaer")
PlanskemaCheckbox := DataGUI.Add("Checkbox", "Section" " X" planskemaX " Y" planskemaY, "Planskema")
DataGUI.Add("Text", " X" planskemaX +100 " Y" planskemaY -20, "Forventet")
PlanskemaEditboxTidligere := DataGUI.Add("Edit", "X" planskemaX + 100 " Y" planskemaY -5, "AB232")
DataGUI.Add("Text", " X" planskemaX +150 " Y" planskemaY -20, "Indlæst")
PlanskemaEditboxTidligere := DataGUI.Add("Edit", "X" planskemaX + 150 " Y" planskemaY -5, "AB232")
planskemaFuldført := DataGUI.Add("Text", " X" planskemaX +200 " Y" planskemaY -5, fuldført)

økonomiskemaCheckbox := DataGUI.Add("Checkbox", "Section" " X" økonomiskemaX " Y" økonomiskemaY, "Økonomiskema")
økonomiskemaEditboxTidligere := DataGUI.Add("Edit", "X" økonomiskemaX + 100 " Y" økonomiskemaY -5, "AB232")
økonomiskemaEditboxTidligere := DataGUI.Add("Edit", "X" økonomiskemaX + 150 " Y" økonomiskemaY -5, "AB232")
økonomiskemaFuldført := DataGUI.Add("Text", " X" økonomiskemaX +200 " Y" økonomiskemaY -5, ikkeFuldført)

vognløbskategoriX := xUdgangspunkt
vognløbskategoriY := yUdgangspunkt + 150

DataGUI.Add("Text", "X" vognløbskategoriX " Y" vognløbskategoriY -20, "Vognløbskategori")
VognløbskategoriCheckbox := DataGUI.Add("Checkbox", "Section" " X" vognløbskategoriX " Y" vognløbskategoriY, "Vognløbskategori")
DataGUI.Add("Text", " X" vognløbskategoriX +100 " Y" vognløbskategoriY -20, "Forventet")
vognløbskategoriEditboxTidligere := DataGUI.Add("Edit", "X" vognløbskategoriX + 100 " Y" vognløbskategoriY -5, "FG8")
DataGUI.Add("Text", " X" vognløbskategoriX +150 " Y" vognløbskategoriY -20, "Indlæst")
vognløbskategoriEditboxTidligere := DataGUI.Add("Edit", "X" vognløbskategoriX + 150 " Y" vognløbskategoriY -5, "FG9")
vognløbskategoriFuldført := DataGUI.Add("Text", " X" vognløbskategoriX +200 " Y" vognløbskategoriY -5, fuldført)

; Omskriv?
vognløbsnotatX := xUdgangspunkt
vognløbsnotatY := yUdgangspunkt + 200
vognløbsnotatEditForventetTekst := "GV 8-16, Type 8 adasdlkjadlkjsaldkjasldladasdlj"
vognløbsnotatEditIndlæstTekst := "GV 8-16, Type 8 sdlfsldflkjglrejg reljg dflgkjfd glkdjg lkjd g"


; DataGUI.Add("Text", "X" vognløbsnotatX " Y" vognløbsnotatY -20, "Vognløbsnotat")
vognløbsnotatCheckbox := DataGUI.Add("Checkbox", "Section" " X" vognløbsnotatX " Y" vognløbsnotatY, "Vognløbsnotat")
DataGUI.Add("Text", " X" vognløbsnotatX " Y" vognløbsnotatY +20, "Forventet")
vognløbsnotatEditboxTidligere := DataGUI.Add("Edit", "X" vognløbsnotatX+ 100 " Y" vognløbsnotatY +20, vognløbsnotatEditForventetTekst) 
DataGUI.Add("Text", " X" vognløbsnotatX " Y" vognløbsnotatY +40, "Indlæst")
vognløbsnotatEditboxTidligere := DataGUI.Add("Edit", "X" vognløbsnotatX +100 " Y" vognløbsnotatY +40, vognløbsnotatEditIndlæstTekst)

; DataGUI.Add("Text", "XP" , "Skema")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; PlanskemaCheckbox := DataGUI.Add("Checkbox", "XS Section", "Planskema")
; DataGUI.Add("Text", , "Skema")
; PlanskemaEditboxTidligere := DataGUI.Add("Edit", EditboxPos , "AB232")
; PlanskemaEditboxNy := DataGUI.Add("Edit",EditboxPos , "AB232")
; ØkonomiskemaCheckBox := DataGUI.Add("Checkbox", "YS+10", "Økonomiskema")
; PlanskemaCheckBox := DataGUI.Add("Edit", "X" PlanskemaEditW "" , "AB232")










DataGUI.Show("AutoSize")