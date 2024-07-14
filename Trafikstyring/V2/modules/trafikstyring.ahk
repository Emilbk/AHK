
;; Slut autoexec
;; Modules import

; Test

+Esc::
{
    ExitApp

    return
}

F3::
{
    P6_nav_kørselsaftale()

    return
}
+F8::
{
    vl := String(P6_hent_data_vognløb_vognløbsnummer_i_planbillede()[1])
    ; MsgBox vl

    return
}

!e::
{
    keywait "alt"
    vognløb := vognløbObj()

    vognløb.hent_data_vognløb_alt_obj()
    SvigtGUIresetfunk("gv_variabel")
    SvigtGUI.tid_for_svigt := FormatTime(, "HH:mm")
    SvigtGUI.vognløbsdato := FormatTime(, "dd-MM-yyyy") ; ændres til indhentet data
    SvigtGUI.Title := "svigt " vognløb.vognløbsnummer " kl. " SvigtGUI.tid_for_svigt
    SvigtGUI_vm_kontakt_tid_edit.Value := FormatTime(, "HHmm")
    SvigtGUI.Show("w448 h357",)

    return
}

;; testing
; +t::
; {

;     obj := vognløbObj()
;     obj.hent_data_vognløb_alt_obj()
;     MsgBox obj.vognløbsdato
; }



