#Requires AutoHotkey v2.0
#Include p6Navigering.ahk
#Include testvl.ahk

; SendInput("^c")
P6_aktiver()
P6_luk_vinduer()
P6_nav_vognløb()
p6_åben_vognløb(test_vl, "MA")
p6_åben_vognløb_kørselsaftale(test_vl)
p6_åben_vognløb_åbningstider(test_vl)
p6_åben_vognløb_resten(test_vl)
; p6_åben_vognløb(test_vl, "TI")
; p6_åben_vognløb(test_vl2, "TI")
; p6_åben_kørselsaftale(test_vl2)
; p6_indlæs_data_kørselsaftale_æ()
; p6_indlæs_data_kørselsaftale_planskema(test_vl)
; p6_indlæs_data_kørselsaftale_økonomiskema(test_vl)

;     kolonne_nummer := Map(

;         "Budnummer", 0,
;         "Vognløbsnummer", 0,
;         "Kørselsaftale", 0,
;         "Styresystem", 0,
;         "Startzone", 0,
;         "Slutzone", 0,
;         "Hjemzone", 0,
;         "MobilnrChf", 0,
;         "Vognløbskategori", 0,
;         "Planskema", 0,
;         "Statistikgruppe", 0,
;         "undtagneTransportTyper", []

;     )


; for index, navn in kolonne_nummer
;     MsgBox index