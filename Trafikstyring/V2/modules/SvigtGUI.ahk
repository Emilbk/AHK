#Requires AutoHotkey v2.0

;; GUI
;; SvigtGUi menu

SvigtGUI_menu := MenuBar()
SvigtGUI_menu_fil := Menu()
SvigtGUI_menu_vogngruppe := Menu()
SvigtGUI_menu_hjælp := Menu()
SvigtGUI_menu_hjælp.Add("&Hjælp", (*) => MsgBox("ikke implementeret"))
SvigtGUI_menu_vogngruppe.Add("&vogngruppesvigt", (*) => MsgBox("ikke implementeret"))

SvigtGUI_menu.Add("&Fil", SvigtGUI_menu_fil)
SvigtGUI_menu.Add("&Vogngruppe", SvigtGUI_menu_vogngruppe)
SvigtGUI_menu.Add("&Om", SvigtGUI_menu_hjælp, "Right")
;; SvigtGUI

SvigtGUI := Gui(, "Svigt")
SvigtGUI.MenuBar := SvigtGUI_menu

SvigtGUI_vognløbsnummer_tekst := SvigtGUI.Add("Text", "x16 y5 w120 h23", "Vognløbsnummer")
SvigtGUI_vognløbsnummer_edit := SvigtGUI.Add("Edit", "vvognløbsnummer_edit x16 y29 w120 h21", "Vognløb")
SvigtGUI_vognløb_status := SvigtGUI.Add("Text", "x16 y58 h50 w100", "Garantiperiode: blalbalba")
SvigtGUI_gvlukket_groupbox := SvigtGUI.Add("Groupbox", "x150 y5 w140 h130", "Hvis &GV lukket:")
SvigtGUI_gvluk_radio_åbningstid := SvigtGUI.Add("Radio", "vgvluk_radio_åbningstid x160 y30", "&Åbningstid udskudt")
SvigtGUI_gvluk_radio_lukket := SvigtGUI.Add("Radio", "vgvluk_radio_lukket x160 y70", "Lukket &midt på VL")
SvigtGUI_gvluk_radio_slettet := SvigtGUI.Add("Radio", "vgvluk_radio_slettet x160 y110", "VL s&lettet")

SvigtGUI_GV_åbningstid_ændret_edit := SvigtGUI.Add("Edit", "vgv_åbningstid_ændret_edit x180 y45 w79 h21", "Vl start kl.")
SvigtGUI_GV_hjemzonetid_edit := SvigtGUI.Add("Edit", "vgv_hjemzonetid x180 y85 w79 h21", "Hjemzone kl.")

SvigtGUI_vl_type_groupbox := SvigtGUI.Add("Groupbox", "x294 y5 w140 h130", "Type VL:")
SvigtGUI_vl_type_radio_gv := SvigtGUI.Add("Radio", "vvl_type_garanti x304 y29 h16", "&Garanti")
SvigtGUI_vl_type_radio_gv_variabel := SvigtGUI.Add("Radio", "vvl_type_gv_variabel x304 y45 w120 h32", "Garantivognløb i variabel tid")
SvigtGUI_vl_type_radio_variabel := SvigtGUI.Add("Radio", "vvl_type_variabel x304 y77 h23", "&Variabel")
SvigtGUI_vl_type_radio_vogngruppe := SvigtGUI.Add("Radio", "vvl_type_vogngruppe x304 y97 h32", "&Vogngruppe")
SvigtGUI_årsag_tekst := SvigtGUI.Add("Text", "x16 y125 w120 h23", "Årsag (valgfri):")
SvigtGUI_årsag_edit := SvigtGUI.Add("Edit", "vårsag x16 y140 w120 h21")
SvigtGUI_årsag_tekst.SetFont("bold")

SvigtGUI_vm_kontakt_groupbox := SvigtGUI.Add("Groupbox", "x150 y135 w283 h48", "Kontakt til vognmand")
SvigtGUI_vm_kontakt_radio_ja := SvigtGUI.Add("Radio", "vvm_kontakt_ja x160 y152 h23", "Kontaktet")
SvigtGUI_vm_kontakt_radio_nej := SvigtGUI.Add("Radio", "vvm_kontakt_nej x240 y152 h23", "Forgæves kontakt")
SvigtGUI_vm_kontakt_tid_edit := SvigtGUI.Add("Edit", "vvm_kontakt_tid x360 y152 w50", "Ca. kl.")
SvigtGUI_beskrivelse_tekst := SvigtGUI.Add("Text", "x16 y165 h23 w100", "&Beskrivelse:")
SvigtGUI_beskrivelse_edit := SvigtGUI.Add("Edit", "vbeskrivelse_edit x16 y185 w410 h106")
SvigtGUI_beskrivelse_tekst.SetFont("bold")
SvigtGUI_forrige_skærmprint_checkbox := SvigtGUI.Add("CheckBox", "vforrige_skærmprint_checkbox x16 y299", "Brug &forrige skærmprint")
SvigtGUI_vis_mail_button := SvigtGUI.Add("Button", "vsendmail_button x160 y314 w60 h23 +default", "&Vis")
SvigtGUI_send_mail_button := SvigtGUI.Add("Button", "vvismail_button x240 y314 w60 h23 +default", "&Send")

SvigtGUI_vis_mail_button.Onevent("Click", SvigtGUI_vis_mail_funk)
SvigtGUI_vl_type_radio_gv.Onevent("Click", SvigtGUI_vl_type_radio_gv_funk)
SvigtGUI_vl_type_radio_gv_variabel.Onevent("Click", SvigtGUI_vl_type_radio_gv_funk)
SvigtGUI_gvluk_radio_åbningstid.Onevent("Click", SvigtGUI_GV_åbningstid_funk)
SvigtGUI_gvluk_radio_lukket.Onevent("Click", SvigtGUI_GV_hjemzonetid_edit_funk)
SvigtGUI_vl_type_radio_variabel.Onevent("Click", SvigtGUI_vl_type_radio_variabel_funk)
SvigtGUI_gvluk_radio_slettet.Onevent("Click", SvigtGUI_gvluk_radio_slettet_funk)

; misc
SvigtGUI.tid_for_svigt := ""

;; navigerings funktioner

; omskrives til at behanlde indhentet data
SvigtGUI_vis_mail_funk(*)
{
    values := SvigtGUI.Submit()
    ; garanti_status := vognløb_indhent_data("31200", "07-07-2024")
    values.tid_for_svigt := FormatTime(,"HH:mm")
    values.vognløbsdato := FormatTime(,"dd-MM-yy")

    if (values.vl_type_garanti)
    {
        brødtekst := svigt_opret_tekst_brødtekst_gv(values)
        emnefelt := svigt_opret_tekst_emnefelt_gv(values)
    }
    if (values.vl_type_gv_variabel)
    {
        brødtekst := svigt_opret_tekst_brødtekst_gv_variabel(values)
        emnefelt := svigt_opret_tekst_emnefelt_gv_variabel(values)
    }
    if (values.vl_type_variabel)
    {
        brødtekst := svigt_opret_tekst_brødtekst_variabel(values)
        emnefelt := svigt_opret_tekst_emnefelt_variabel(values)
    }
    if (values.vl_type_vogngruppe)
    {
        brødtekst := svigt_opret_tekst_brødtekst_vogngruppe(values)
        emnefelt := svigt_opret_tekst_emnefelt_vogngruppe(values)
    }

    MsgBox "Emnefelt: " emnefelt "`n`n Brødtekst: " brødtekst
    return
}
SvigtGUI_gvluk_radio_slettet_funk(*)
{

    SvigtGUI_GV_åbningstid_ændret_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 1
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 1
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 1

    return
}
SvigtGUI_GV_åbningstid_funk(*)
{
    SvigtGUI_GV_åbningstid_ændret_edit.Enabled := 1
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_ændret_edit.Focus()

    return
}

SvigtGUI_GV_hjemzonetid_edit_funk(*)
{
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 1
    SvigtGUI_GV_hjemzonetid_edit.Focus()
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 1
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 1
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 1


    return
}

SvigtGUI_vl_type_radio_gv_variabel_funk(*)
{
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 1
    SvigtGUI_gvluk_radio_lukket.Enabled := 1
    SvigtGUI_gvluk_radio_slettet.Enabled := 1
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Value := 0
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 0
    SvigtGUI_vm_kontakt_radio_nej.Value := 0
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_tid_edit.Value := "Ca. kl."


    return
}


SvigtGUI_vl_type_radio_gv_funk(*)
{
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 1
    SvigtGUI_gvluk_radio_lukket.Enabled := 1
    SvigtGUI_gvluk_radio_slettet.Enabled := 1

    return
}


SvigtGUI_vl_type_radio_variabel_funk(*)
{
    SvigtGUI_gvluk_radio_slettet.Enabled := 0
    SvigtGUI_gvluk_radio_slettet.Value := 0
    SvigtGUI_gvluk_radio_lukket.Enabled := 0
    SvigtGUI_gvluk_radio_lukket.Value := 0
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 0
    SvigtGUI_gvluk_radio_åbningstid.Value := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_ændret_edit.Enabled := 0
    SvigtGUI_GV_åbningstid_ændret_edit.Value := "Vl start kl."
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Value := 0
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 0
    SvigtGUI_vm_kontakt_radio_nej.Value := 0
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_tid_edit.Value := "Ca. kl."


    return
}

SvigtGUIresetfunk(vl_type)
{
    SvigtGUI_gvluk_radio_slettet.Enabled := 0
    SvigtGUI_gvluk_radio_slettet.Value := 0
    SvigtGUI_gvluk_radio_lukket.Enabled := 0
    SvigtGUI_gvluk_radio_lukket.Value := 0
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 0
    SvigtGUI_gvluk_radio_åbningstid.Value := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_ændret_edit.Enabled := 0
    SvigtGUI_GV_åbningstid_ændret_edit.Value := "Vl start kl."
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Value := 0
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 0
    SvigtGUI_vm_kontakt_radio_nej.Value := 0
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_tid_edit.Value := "Ca. kl."


    SvigtGUI_vl_type_radio_vogngruppe.Enabled := 1

    SvigtGUI_forrige_skærmprint_checkbox.Enabled := 0
    if (DllCall("IsClipboardFormatAvailable", "uint", 2))
    {
        SvigtGUI_forrige_skærmprint_checkbox.Enabled := 1
    }

    if (vl_type := "variabel")
    {
        SvigtGUI_vl_type_radio_variabel.Value := 1
        SvigtGUI_vl_type_radio_variabel_funk()
        SvigtGUI_beskrivelse_edit.Focus()
    }
    if (vl_type := "gv")
    {
        SvigtGUI_vl_type_radio_gv.Value := 1
        SvigtGUI_vl_type_radio_gv_funk()
        SvigtGUI_beskrivelse_edit.Focus()
    }
    if (vl_type := "gv_variabel")
    {
        SvigtGUI_vl_type_radio_gv_variabel.Value := 1
        SvigtGUI_vl_type_radio_gv_variabel_funk()
        SvigtGUI_beskrivelse_edit.Focus()
    }


    return
}

;; Svigt oprettelse tekst

svigt_opret_tekst_emnefelt_gv(input)
{
    emnefelt := ""
    emnefelt_vognløbsnummer := input.vognløbsnummer_edit
    emnefelt_tid_for_svigt := input.tid_for_svigt
    emnefelt_vognløbsdato := input.vognløbsdato
    emnefelt_tid_for_svigt := input.tid_for_svigt

    if (input.gvluk_radio_lukket and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": lukket i hjemzone kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_lukket and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": " input.årsag " - lukket i hjemzone kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (input.gvluk_radio_slettet and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": vognløb slettet kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_slettet and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": " input.årsag " -  vognløb slettet kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (input.gvluk_radio_åbningstid and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": Åbningstid udskudt. Kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_åbningstid and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": " input.årsag " - Åbningstid udskudt. Kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (!input.gvluk_radio_lukket and !input.gvluk_radio_slettet and !input.gvluk_radio_åbningstid and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (!input.gvluk_radio_lukket and !input.gvluk_radio_slettet and !input.gvluk_radio_åbningstid and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": kl. " input.årsag " - " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    return emnefelt

}
svigt_opret_tekst_brødtekst_gv(input_svigt)
{
    vl_status := "09-17" ; ændres til indhentet data

    brødtekst := ""
    brødtekst_vm_kontakt := ""
    brødtekst_vl_status := ""
    brødtekst_garanti_tid := "garanti " vl_status ", "

    ; VM kontakt
    if (input_svigt.vm_kontakt_ja)
        brødtekst_vm_kontakt := ". VM kontaktet ca. " input_svigt.vm_kontakt_tid " — "
    if (input_svigt.vm_kontakt_nej)
        brødtekst_vm_kontakt := ". VM forsøgt kontaktet ca. " input_svigt.vm_kontakt_tid " — "
    if (!input_svigt.vm_kontakt_ja and !input_svigt.vm_kontakt_nej)
        brødtekst_vm_kontakt := " — "


    ; GV lukket
    if (input_svigt.gvluk_radio_lukket)
        brødtekst_vl_status := "lukket kl. " input_svigt.gv_hjemzonetid
    if (input_svigt.gvluk_radio_slettet)
        brødtekst_vl_status := "vl slettet"
    if (input_svigt.gvluk_radio_åbningstid)
    {
        brødtekst_vl_status := "åbningstid ændret til " input_svigt.gv_åbningstid_ændret_edit
        brødtekst_vm_kontakt := " — "
    }
    ; alm svigt
    if (!input_svigt.gvluk_radio_lukket and !input_svigt.gvluk_radio_åbningstid and !input_svigt.gvluk_radio_slettet)
    {
        brødtekst_garanti_tid := ""
        brødtekst_vl_status := ""
        brødtekst_vm_kontakt := ""
    }

    brødtekst := brødtekst_garanti_tid brødtekst_vl_status brødtekst_vm_kontakt input_svigt.beskrivelse_edit

    return brødtekst
}
svigt_opret_tekst_emnefelt_gv_variabel(input)
{
    emnefelt := ""
    emnefelt_vognløbsnummer := input.vognløbsnummer_edit
    emnefelt_tid_for_svigt := input.tid_for_svigt
    emnefelt_vognløbsdato := input.vognløbsdato
    emnefelt_tid_for_svigt := input.tid_for_svigt

    if (input.gvluk_radio_lukket and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): lukket i hjemzone kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_lukket and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): " input.årsag " - lukket i hjemzone kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (input.gvluk_radio_slettet and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): vognløb slettet kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_slettet and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): " input.årsag " -  vognløb slettet kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (input.gvluk_radio_åbningstid and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): Åbningstid udskudt. Kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (input.gvluk_radio_åbningstid and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): " input.årsag " - Åbningstid udskudt. Kl. " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    if (!input.gvluk_radio_lukket and !input.gvluk_radio_slettet and !input.gvluk_radio_åbningstid and !input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if (!input.gvluk_radio_lukket and !input.gvluk_radio_slettet and !input.gvluk_radio_åbningstid and input.årsag)
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer " (variabel tid): " input.årsag " - " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    return emnefelt

}

svigt_opret_tekst_brødtekst_gv_variabel(input_svigt)
{
    vl_status := "09-17" ; ændres til indhentet data

    brødtekst := ""
    brødtekst_vl_status := ""
    brødtekst_garanti_tid := "garanti " vl_status ", "
    brødtekst_opdeler := " — "


    ; VM kontakt
    if (input_svigt.vm_kontakt_ja)
        brødtekst_vm_kontakt := ". VM kontaktet ca. " input_svigt.vm_kontakt_tid
    if (input_svigt.vm_kontakt_nej)
        brødtekst_vm_kontakt := ". VM forsøgt kontaktet ca. " input_svigt.vm_kontakt_tid
    if (!input_svigt.vm_kontakt_ja and !input_svigt.vm_kontakt_nej)
        brødtekst_vm_kontakt := ""

    ; GV slettet
    if (input_svigt.gvluk_radio_slettet)
        brødtekst_vl_status := "gv slettet ifm. variabel kørsel"

    ; GV lukket i variabel tid
    if (input_svigt.gvluk_radio_lukket)
        brødtekst_vl_status := "gv lukket i hjemzone kl. " input_svigt.gv_hjemzonetid " ifm. variabel kørsel"

    ; åbningstid ændret
    if (input_svigt.gvluk_radio_åbningstid)
    {
        brødtekst_vl_status := "åbningstid ændret til " input_svigt.gv_åbningstid_ændret_edit
    }

    ; alm svigt
    if (!input_svigt.gvluk_radio_lukket and !input_svigt.gvluk_radio_åbningstid and !input_svigt.gvluk_radio_slettet)
    {
        brødtekst_garanti_tid := ""
        brødtekst_vl_status := ""
        brødtekst_vm_kontakt := ""
    }
    brødtekst := brødtekst_garanti_tid brødtekst_vl_status brødtekst_vm_kontakt brødtekst_opdeler input_svigt.beskrivelse_edit

    return brødtekst
}


svigt_opret_tekst_emnefelt_variabel(input)
{
    emnefelt := ""
    emnefelt_vognløbsnummer := input.vognløbsnummer_edit
    emnefelt_tid_for_svigt := input.tid_for_svigt
    emnefelt_vognløbsdato := input.vognløbsdato
    emnefelt_tid_for_svigt := input.tid_for_svigt

    if input.årsag
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": " input.årsag " - " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if !input.årsag
        emnefelt := "Svigt VL " emnefelt_vognløbsnummer ": " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    return emnefelt
}

svigt_opret_tekst_brødtekst_variabel(input_svigt)
{
    brødtekst := ""

    ; alm svigt
    brødtekst := input_svigt.beskrivelse_edit

    return brødtekst
}

svigt_opret_tekst_emnefelt_vogngruppe(input)
{
    emnefelt := ""
    emnefelt_vognløbsnummer := input.vognløbsnummer_edit
    emnefelt_tid_for_svigt := input.tid_for_svigt
    emnefelt_vognløbsdato := input.vognløbsdato
    emnefelt_tid_for_svigt := input.tid_for_svigt
    emnefelt_vogngruppe := input.vogngruppe

    if input.årsag
        emnefelt := "Svigt " emnefelt_vogngruppe ", VL" emnefelt_vognløbsnummer ": " input.årsag " - " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato
    if !input.årsag
        emnefelt := "Svigt " emnefelt_vogngruppe ", VL" emnefelt_vognløbsnummer ": " emnefelt_tid_for_svigt " d. " emnefelt_vognløbsdato

    return emnefelt
}
svigt_opret_tekst_brødtekst_vogngruppe(input_svigt)
{
    brødtekst := ""
    brødtekst_vognløb := "vognløb " input_svigt.vognløbsnummer_edit
    brødtekst_beskrivelse := input_svigt.beskrivelse_edit
    brødtekst_opdeler := " — "
    ; alm svigt
    brødtekst := brødtekst_vognløb brødtekst_opdeler brødtekst_beskrivelse

    return brødtekst
}


;; Svigt

; SvigtGIU
; Sluse for forkert input, hvor skal den placeres?

; Vognløbsdata
; tjek et vognløb for garantitid/variabel - ferie osv.
; Indhent VM-data (inklusiv VM-telefon?)

; Svigtdata
; opbyg svigttekst
; Hvordan laves en fornuftig logik?

; Mailfunktion
; Lav/send outlookmail


;; Svigtdata

; To niveauer af mail, emnefelt og brødtekst (to funktioner?)
; Logik for behandling af input

;; Oprettelse af tekst

; Mulige inputs

; VL:
; 1. Variabel
; 2. Garantivogn
; 3. Garantivogn (variabel)

; Vognløbsbehandling, ved GV:
; 1.Åbningstid udskudt, tidspunkt
; 2. VL lukket efter start, tidspnkt
; 3. VL slettet
; 4. Variabel tid fjernet på GV(?) - magen til 1?

; VM-kontakt, ved GV:
; 1. Ja
; 2. Nej

; Øvrige, for alle:
; 1. tid for svigt
; 2. beskrivelse
; 3. Emnefelt ja
; 4. Emnefelt nej


;; Outputs brødtekst
; 1. Alm. svigt, GV
; 3. GV vl slettet
; 2. Alm svigt, GV variabel og variable
; 4. GV vl lukket efter start
; 5. GV åbningstid udskudt

; sluse tjek for type vl
; input
; input_test := Map(
;     "Vl_type", "[1, 2, 3]",
;     ; GV-behandling
;     "GV_behandling", "[0, 1, 2, 3, 4]",
;     "VM_kontakt", "[1, 2]",
;     "VM_tid_for_kontakt", "str_tid_vm",
;     ; øvrigt
;     "Tid_for_svigt", "str_tid_svigt",
;     "Beskrivelse", "str_beskrivelse",
;     "Emnefelt_beskrivelse", "[0, 1]",
;     "Emnefelt_tekst", "str_emnefelt_tekst"
; )

; GV-behandling
; t::


; svigt_opret_tekst_brødtekst_gv_variabel()
; {

; }
; svigt_opret_tekst_brødtekst_variabel()
; {

; }


;; testing

test_reset()
{
    input_test := Object()
    input_test.beskrivelse_edit := "Svigtet beskrives her"
    input_test.forrige_skærmprint_checkbox := 0
    input_test.gv_hjemzonetid := "1231"
    input_test.gv_åbningstid_ændret_edit := "1231"
    input_test.gvluk_radio_lukket := 0
    input_test.gvluk_radio_slettet := 0
    input_test.gvluk_radio_åbningstid := 0
    input_test.vl_type_variabel := 0
    input_test.vl_type_garanti := 0
    input_test.vl_type_gv_variabel := 0
    input_test.vm_kontakt_ja := 0
    input_test.vm_kontakt_nej := 0
    input_test.vm_kontakt_tid := "1454"
    input_test.vognløbsnummer_edit := "31200"
    input_test.årsag := ""
    input_test.tid_for_svigt := "15:06"
    input_test.vognløbsdato := "07-07-2024"

    return input_test
}

test(input)
{
    input := test_reset()
    ; garnativogn lukket, vm kontakt
    input.vl_type_garanti := 1
    input.gvluk_radio_lukket := 1
    input.vm_kontakt_ja := 1
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv lukket"

    input := test_reset()
    ; garnativogn lukket, vm kontakt, årsag
    input.vl_type_garanti := 1
    input.gvluk_radio_lukket := 1
    input.vm_kontakt_ja := 1
    input.årsag := "bil punkteret"
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv lukket"


    input := test_reset()
    ; garnativogn slettet, vm ingen kontakt
    input.gvluk_radio_slettet := 1
    input.vm_kontakt_nej := 1
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv slettet"

    input := test_reset()
    ; garnativogn slettet, vm ingen kontakt, årsag
    input.gvluk_radio_slettet := 1
    input.vm_kontakt_nej := 1
    input.årsag := "bil punkteret"
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv slettet"


    ; garantivogn ændret åbningstid
    input := test_reset()
    input.gvluk_radio_åbningstid := 1
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv ændret åbningstid"


    ; garantivogn alm svigt
    input := test_reset()
    input.vl_type_garanti := 1
    brødtekst := svigt_opret_tekst_brødtekst_gv(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv alm svigt"


    ; variabel svigt
    input := test_reset()
    input.vl_type_variabel := 1
    input.årsag := "ikke i hjemzone"
    brødtekst := svigt_opret_tekst_brødtekst_variabel(input)
    emnefelt := svigt_opret_tekst_emnefelt_variabel(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "variabel svigt"


    input := test_reset()
    ; gv variabel, vl slettet
    input.vl_type_radio_variabel := 1
    input.gvluk_radio_slettet := 1
    brødtekst := svigt_opret_tekst_brødtekst_gv_variabel(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv_variabel(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv variabel slettet"


    input := test_reset()
    ; gv variabel, åbningstid ændret
    input.vl_type_gv_variabel := 1
    input.gvluk_radio_åbningstid := 1
    input.gv_åbningstid_ændret_edit := "14:00"

    brødtekst := svigt_opret_tekst_brødtekst_gv_variabel(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv_variabel(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv variabel åbningstid ændret"


    input := test_reset()
    ; gv variabel, vl lukket
    input.vl_type_gv_variabel := 1
    input.gvluk_radio_lukket := 1
    input.gv_hjemzonetid := "13:21"
    input.vm_kontakt_ja := 1

    brødtekst := svigt_opret_tekst_brødtekst_gv_variabel(input)
    emnefelt := svigt_opret_tekst_emnefelt_gv_variabel(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "gv variabel vl lukket"


    ; vogngruppesvigt
    input := test_reset()
    input.vvl_type_vogngruppe := 1
    input.vognløbsnummer_edit := "5023, 5054, 5034 og 5031"
    input.vogngruppe := "Aarhusstat" ; data skal indhentes
    input.årsag := "en årsag"

    brødtekst := svigt_opret_tekst_brødtekst_vogngruppe(input)
    emnefelt := svigt_opret_tekst_emnefelt_vogngruppe(input)
    MsgBox "Emnefelt: " emnefelt "`nBrødtekst: " brødtekst, "vogngruppe"


}