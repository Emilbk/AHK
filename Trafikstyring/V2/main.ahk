#Requires AutoHotkey v2.0
FileEncoding "UTF-8"
Persistent

#include "modules/opstart.ahk"
#include "modules/svigtGUI.ahk"
#include "modules/P6.ahk"
; #include "modules/vognløbsdata.ahk"
#include "modules/class_vl.ahk"
#include "modules/trafikstyring.ahk"
#include "test.ahk"
; metadata

admin_mail := "ebk@midttrafik.dk"




; vl := vognløbObj()
; vl.kørselsaftale := "3100_47"
; vl.unpack_garantidata("34352")
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

opret_svigt(p_vl_obj)
{
    tid_for_svigt := p_vl_obj.vognløbsdato_timestamp . SubStr(a_now, -6, 6)

    MsgBox(FormatTime(tid_for_svigt, "HH:mm"))
    p_vl_obj.vognløb_status()


    SvigtGUIresetfunk()

    SvigtGUI.Title := "Svigt vl. " p_vl_obj.vognløbsnummer " d. " FormatTime(p_vl_obj.vognløbsdato_timestamp, "dd-MM-yy") " kl. " FormatTime(tid_for_svigt, "HH:mm")
    SvigtGUI_vognløbsnummer_edit.Value := p_vl_obj.vognløbsnummer
    SvigtGUI_vognløb_status.Text := p_vl_obj.status


    if p_vl_obj.vl_type = "aktiv garanti"
        SvigtGUI_vl_type_radio_gv.Value := 1
    if p_vl_obj.vl_type = "variabel garanti"
        SvigtGUI_vl_type_radio_gv_variabel.Value := 1
    if p_vl_obj.vl_type = "variabel"
        SvigtGUI_vl_type_radio_variabel.Value := 1
    if p_vl_obj.vl_type = "vogngruppe"
        SvigtGUI_vl_type_radio_vogngruppe.Value := 1


    SvigtGUI_beskrivelse_edit.Focus()
    SvigtGUI.Show("w448 h457",)
    return
}

return


!e::
{
    ; vl_data := P6_hent_data_vognløb_alt()
    ; test_vl.hent_data_vognløb_alt_obj(vl_data)
    ; opret_svigt(test_vl)

    return
}