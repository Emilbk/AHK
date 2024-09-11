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
    ; vl := String(P6_hent_data_vognløb_vognløbsnummer_i_planbillede()[1])
    ; MsgBox vl

    return
}

; !e::
; {
;     keywait "alt"
;     svigt_vl := vognløbObj()
;     svigt_vl.vognløbsnummer := "31200"
;     svigt_vl.kørselsaftale := "31002_47"
;     ; svigt_vl.hent_data_vognløb_alt_obj()
;     svigt_vl.vognløb_status("202407141500")

;     SvigtGUIresetfunk("gv_variabel")

;     if svigt_vl.variabel
;         SvigtGUI_vl_type_radio_variabel.Value := 1
;     if svigt_vl.gv
;        SvigtGUI_vl_type_radio_gv.Value := 1
;     if svigt_vl.gv_variabel
;         SvigtGUI_vl_type_radio_gv_variabel.Value := 1
;     if svigt_vl.vg
;         SvigtGUI_vl_type_radio_vogngruppe.Value := 1


;     SvigtGUI.tid_for_svigt := FormatTime(, "HH:mm")
;     SvigtGUI.vognløbsdato := FormatTime(, "dd-MM-yyyy") ; ændres til indhentet data
;     SvigtGUI.Title := "svigt " svigt_vl.vognløbsnummer " kl. " SvigtGUI.tid_for_svigt
;     SvigtGUI_vognløb_status := svigt_vl.status
;     SvigtGUI_vm_kontakt_tid_edit.Value := FormatTime(, "HHmm")
;     SvigtGUI.Show("w448 h357",)

;     return
; }

;; testing
; +t::
; {

;     obj := vognløbObj()
;     obj.hent_data_vognløb_alt_obj()
;     MsgBox obj.vognløbsdato
; }
