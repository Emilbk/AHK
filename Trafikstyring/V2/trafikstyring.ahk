#Requires AutoHotkey v2.0
;; Modules import
#include "modules/svigtGUI.ahk"
#include "modules/P6.ahk"
#include "modules/vognløbsdata.ahk"
FileEncoding "UTF-8"
;; Slut autoexec
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

    SvigtGUIresetfunk("gv_variabel")
    SvigtGUI.tid_for_svigt := FormatTime(, "HH:mm")
    SvigtGUI.vognløbsdato := FormatTime(, "dd-MM-yyyy") ; ændres til indhentet data
    SvigtGUI.Title := "svigt 31200 kl. " SvigtGUI.tid_for_svigt
    SvigtGUI_vm_kontakt_tid_edit.Value := FormatTime(, "HHmm")
    SvigtGUI.Show("w448 h357",)

    return
}

;; testing
+t::
{

    values_test := test_reset()
    test(values_test)
}

