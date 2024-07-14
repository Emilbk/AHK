#Requires AutoHotkey v2.0
FileEncoding "UTF-8"

#include "modules/opstart.ahk"
#include "modules/svigtGUI.ahk"
#include "modules/P6.ahk"
; #include "modules/vognløbsdata.ahk"
#include "modules/class_vl.ahk"
#include "modules/trafikstyring.ahk"
; metadata

admin_mail := "ebk@midttrafik.dk"




vl := vognløbObj()
vl.kørselsaftale := "3100_47"
vl.unpack_garantidata("34352")
; vl.kørselsaftale := "3256_26" ; gv over midnat
; ; vl.kørselsaftale := "3253_22" ; gv med forskellige periode hverdag og weekend
; vl.kørselsaftale := "3100_47"

; ; vlobj.hent_data_vognløb_alt_obj()

; vl.vognløb_status("202406050559")
; MsgBox vl.status

send_fejl_meddelelse(Exception, *)
{
    ; skriv send fejlmeddelse til admin-mail
    msgbox Exception.Message "`n" Exception.what
    return
}

opret_svigt()
{
    svigt_vl := vognløbObj()

    ; hent vognløbsdata
    ; data := P6_hent_data_vognløb_alt()
    ; svigt_vl.vognløbsnummer := data[1]
    ; svigt_vl.kørselsaftale := data[3] "_" data[4]
    ; svigt_vl.vognløbsdato := data[2]

    svigt_vl.vognløbsnummer := "31200"
    svigt_vl.kørselsaftale := "3256_262" ; gv over midnat
    ; svigt_vl.kørselsaftale := "3100_47"
    svigt_vl.vognløbsdato := A_Now
    svigt_vl.vognløb_status()


    SvigtGUIresetfunk()

    SvigtGUI.Title := "Svigt vl. " svigt_vl.vognløbsnummer " d. " FormatTime(svigt_vl.vognløbsdato, "dd/MM/yy") " kl. " FormatTime(svigt_vl.vognløbsdato, "HH:mm")
    SvigtGUI_vognløbsnummer_edit.Value := svigt_vl.vognløbsnummer
    SvigtGUI_vognløb_status.Text := svigt_vl.status


    if svigt_vl.gv
        SvigtGUI_vl_type_radio_gv.Value := 1
    if svigt_vl.gv_variabel
        SvigtGUI_vl_type_radio_gv_variabel.Value := 1
    if svigt_vl.variabel
        SvigtGUI_vl_type_radio_variabel.Value := 1
    if svigt_vl.vogngruppe
        SvigtGUI_vl_type_radio_vogngruppe.Value := 1


    SvigtGUI_beskrivelse_edit.Focus()
    SvigtGUI.Show("w448 h457",)
    return
}

return

!e::
{
    opret_svigt()

    return
}