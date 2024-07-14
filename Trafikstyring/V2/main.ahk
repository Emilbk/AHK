#Requires AutoHotkey v2.0
FileEncoding "UTF-8"

#include "modules/opstart.ahk"
#include "modules/svigtGUI.ahk"
#include "modules/P6.ahk"
; #include "modules/vognløbsdata.ahk"
#include "modules/class_vl.ahk"
#include "modules/trafikstyring.ahk"

vl := vognløbObj()

vl.kørselsaftale := "3256_26" ; gv over midnat
; vl.kørselsaftale := "3253_22" ; gv med forskellige periode hverdag og weekend
; vl.kørselsaftale := "3100_47"

; vlobj.hent_data_vognløb_alt_obj()

vl.vognløb_status("202407051500")
MsgBox vl.status
vl.vognløb_status("202407061630")

MsgBox vl.status



return

