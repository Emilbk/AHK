#Requires AutoHotkey v2.0

#Include vlObj.ahk
#Include dataGUI.ahk
#Include excelObj.ahk

test_vl := vlObj()

test_vl.vl_data["Vognløbsnummer"] := "31400"
test_vl.vl_data["Kørselsaftale"] := "3400"
test_vl.vl_data["Styresystem"] := "1"
test_vl.vl_data["Planskema"] := "31300" 
test_vl.vl_data["Økonomiskema"] := "31200"
test_vl.vl_data["Startzone"] := "Årh804"
test_vl.vl_data["Slutzone"] := "Årh804"
test_vl.vl_data["Hjemzone"] := "Årh804"
; test_vl.vl_data["Vognløbsnotering"] := 0
test_vl.vl_data["Vognløbsnotering"] := "Ny notering til VL"
test_vl.vl_data["Vognløbskategori"] := "9999"
test_vl.vl_data["MobilnrChf"] := "70112210"
test_vl.vl_data["Statistikgruppe"] := "2GVEL"
; test_vl.vl_data["Undtagne transporttyper"] := 0
test_vl.vl_data["Undtagne transporttyper"] := ["LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER" ]
test_vl.vl_data["Dato"] := ["MA"]
test_vl.vl_data["Starttid"] := "08:00"
test_vl.vl_data["Sluttid"] := "17:00"


test_vl2 := vlObj()

test_vl2.vl_data["Vognløbsnummer"] := "31400"
test_vl2.vl_data["Kørselsaftale"] := "3400"
test_vl2.vl_data["Styresystem"] := "1"
test_vl2.vl_data["Planskema"] := "31400" 
test_vl2.vl_data["Økonomiskema"] := "31400"