#Requires AutoHotkey v2.0

sleep_konstant := 1

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
SvigtGUI_vognløb_beskrivelse := SvigtGUI.Add("Text", "x16 y58 h50 w100", "Garantiperiode: blalbalba")
SvigtGUI_gvlukket_groupbox := SvigtGUI.Add("Groupbox", "x150 y5 w140 h130", "Hvis &GV lukket:")
SvigtGUI_gvluk_radio_åbningstid := SvigtGUI.Add("Radio", "vgvluk_radio_åbningstid x160 y30", "&Åbningstid udskudt")
SvigtGUI_gvluk_radio_lukket := SvigtGUI.Add("Radio", "vgvluk_radio_lukket x160 y70", "Lukket &midt på VL")
SvigtGUI_gvluk_radio_slettet := SvigtGUI.Add("Radio", "vgvluk_radio_slettet x160 y110", "VL s&lettet")

SvigtGUI_GV_åbningstid_edit := SvigtGUI.Add("Edit", "vgv_åbningstid x180 y45 w79 h21", "Vl start kl.")
SvigtGUI_GV_hjemzonetid_edit := SvigtGUI.Add("Edit", "vgv_hjemzonetid x180 y85 w79 h21", "Hjemzone kl.")

SvigtGUI_vltype_groupbox := SvigtGUI.Add("Groupbox", "x294 y5 w140 h130", "Type VL:")
SvigtGUI_vltype_radio_gv := SvigtGUI.Add("Radio", "vvltype_garanti x304 y29 h16", "&Garanti")
SvigtGUI_vltype_radio_gv_variabel := SvigtGUI.Add("Radio", "vvltype_gv_variabel x304 y45 w120 h32", "Garantivognløb i variabel tid")
SvigtGUI_vltype_radio_variabel := SvigtGUI.Add("Radio", "vvl_type_variabel x304 y77 h23", "&Variabel")
SvigtGUI_vltype_radio_vogngruppe := SvigtGUI.Add("Radio", "vvl_type_vogngruppe x304 y97 h32", "&Vogngruppe")
SvigtGUI_årsag_tekst := SvigtGUI.Add("Text", "x16 y125 w120 h23", "Årsag (valgfri):")
SvigtGUI_årsag_edit := SvigtGUI.Add("Edit", "vårsag x16 y140 w120 h21")
SvigtGUI_årsag_tekst.SetFont("bold")

SvigtGUI_vm_kontakt_groupbox := SvigtGUI.Add("Groupbox", "x150 y135 w283 h48", "Kontakt til vognmand")
SvigtGUI_vm_kontakt_radio_ja := SvigtGUI.Add("Radio", "vvm_kontakt_ja x160 y152 h23", "Kontaktet")
SvigtGUI_vm_kontakt_radio_nej := SvigtGUI.Add("Radio", "vvm_kontakt_nej x240 y152 h23", "Forgæves kontakt")
SvigtGUI_vm_kontakt_tid_edit := SvigtGUI.Add("Edit", "vvm_kontakt_tid x360 y152 w50", "Ca. kl.")
SvigtGUI_beskrivelse_tekst := SvigtGUI.Add("Text", "x16 y165 h23 w100", "&Beskrivelse:")
SvigtGUI_beskrivelse_edit := SvigtGUI.Add("Edit", "vbeskrivelse x16 y185 w410 h106")
SvigtGUI_beskrivelse_tekst.SetFont("bold")
SvigtGUI_forrige_skærmprint_checkbox := SvigtGUI.Add("CheckBox", "vforrige_skærmprint_checkbox x16 y299", "Brug &forrige skærmprint")
SvigtGUI_vis_mail_button := SvigtGUI.Add("Button", "vsendmail_button x160 y314 w60 h23 +default", "&Vis")
SvigtGUI_send_mail_button := SvigtGUI.Add("Button", "vvismail_button x240 y314 w60 h23 +default", "&Send")

SvigtGUI_vis_mail_button.Onevent("Click", SvigtGUI_vis_mail_funk)
SvigtGUI_vltype_radio_gv.Onevent("Click", SvigtGUI_vltype_radio_gv_funk)
SvigtGUI_vltype_radio_gv_variabel.Onevent("Click", SvigtGUI_vltype_radio_gv_funk)
SvigtGUI_gvluk_radio_åbningstid.Onevent("Click", SvigtGUI_GV_åbningstid_funk)
SvigtGUI_gvluk_radio_lukket.Onevent("Click", SvigtGUI_GV_hjemzonetid_edit_funk)
SvigtGUI_vltype_radio_variabel.Onevent("Click", SvigtGUI_vltype_radio_variabel_funk)
SvigtGUI_gvluk_radio_slettet.Onevent("Click", SvigtGUI_gvluk_radio_slettet_funk)
; SvigtGUI_vltype_radio_gv_variabel.Onevent("Click", SvigtGUI_vltype_gv_funk)
;; Slut autoexec
; Test

SvigtGUI_vis_mail_funk(*)
{
    values := SvigtGUI.Submit()

    return
}
SvigtGUI_gvluk_radio_slettet_funk(*)
{
    
    SvigtGUI_GV_åbningstid_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 1
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 1
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 1

    return
}
SvigtGUI_GV_åbningstid_funk(*)
{
    SvigtGUI_GV_åbningstid_edit.Enabled := 1
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_edit.Focus()

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

SvigtGUI_vltype_radio_gv_variabel_funk(*)
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


SvigtGUI_vltype_radio_gv_funk(*)
{
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 1
    SvigtGUI_gvluk_radio_lukket.Enabled := 1
    SvigtGUI_gvluk_radio_slettet.Enabled := 1

    return
}


SvigtGUI_vltype_radio_variabel_funk(*)
{
    SvigtGUI_gvluk_radio_slettet.Enabled := 0
    SvigtGUI_gvluk_radio_slettet.Value := 0
    SvigtGUI_gvluk_radio_lukket.Enabled := 0
    SvigtGUI_gvluk_radio_lukket.Value := 0
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 0
    SvigtGUI_gvluk_radio_åbningstid.Value := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_edit.Enabled := 0
    SvigtGUI_GV_åbningstid_edit.Value := "Vl start kl."
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Value := 0
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 0
    SvigtGUI_vm_kontakt_radio_nej.Value := 0
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_tid_edit.Value := "Ca. kl."



    return
}
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
    SvigtGUI.Title := "ssdflkjsfd"
    SvigtGUI.Show("w448 h357",)

    return
}

SvigtGUIresetfunk(vltype)
{
    SvigtGUI_gvluk_radio_slettet.Enabled := 0
    SvigtGUI_gvluk_radio_slettet.Value := 0
    SvigtGUI_gvluk_radio_lukket.Enabled := 0
    SvigtGUI_gvluk_radio_lukket.Value := 0
    SvigtGUI_gvluk_radio_åbningstid.Enabled := 0
    SvigtGUI_gvluk_radio_åbningstid.Value := 0
    SvigtGUI_GV_hjemzonetid_edit.Enabled := 0
    SvigtGUI_GV_hjemzonetid_edit.Value := "Hjemzone kl."
    SvigtGUI_GV_åbningstid_edit.Enabled := 0
    SvigtGUI_GV_åbningstid_edit.Value := "Vl start kl."
    SvigtGUI_vm_kontakt_radio_ja.Enabled := 0
    SvigtGUI_vm_kontakt_radio_ja.Value := 0
    SvigtGUI_vm_kontakt_radio_nej.Enabled := 0
    SvigtGUI_vm_kontakt_radio_nej.Value := 0
    SvigtGUI_vm_kontakt_tid_edit.Enabled := 0
    SvigtGUI_vm_kontakt_tid_edit.Value := "Ca. kl."


    SvigtGUI_vltype_radio_vogngruppe.Enabled := 0

    SvigtGUI_forrige_skærmprint_checkbox.Enabled := 0
    if (DllCall("IsClipboardFormatAvailable", "uint", 2))
    {
        SvigtGUI_forrige_skærmprint_checkbox.Enabled := 1
    }

    if (vltype := "variabel")
        {
        SvigtGUI_vltype_radio_variabel.Value := 1
        SvigtGUI_vltype_radio_variabel_funk()
        SvigtGUI_beskrivelse_edit.Focus()
        }
    if (vltype := "gv")
        {
        SvigtGUI_vltype_radio_gv.Value := 1
        SvigtGUI_vltype_radio_gv_funk()
        SvigtGUI_beskrivelse_edit.Focus()
        }
    if (vltype := "gv_variabel")
        {
        SvigtGUI_vltype_radio_gv_variabel.Value := 1
        SvigtGUI_vltype_radio_gv_variabel_funk()
        SvigtGUI_beskrivelse_edit.Focus()
        }



        return
}

P6_var_sleep(sleep_var)
{
    global sleep_konstant
    sleep sleep_konstant * sleep_var

    return
}
;; P6-navigering

; Aktiverer P6-vindue, hvis ikke aktivt
P6_nav_aktiver()
{

    HotIfWinNotActive "PLANET"
    {
        WinActivate "PLANET"
        WinWaitActive "PLANET"
        sleep 100
        SendInput "{esc}" ; registrerer ikke første tryk, når der skiftes til vindue
        ; sleep 300
        return true
    }
    return false
}

; Aktiverer alt-menu i P6, tager op til to taste-sekvenser
P6_nav_alt_menu(tast1, tast2?)
{
    SendInput "{alt}"
    sleep 20
    Sendinput tast1
    if IsSet(tast2)
    {
        sleep 40
        SendInput tast2
        sleep 40
    }

    return
}

;; P6-navigering, vinduer

;
P6_nav_planbillede()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("tp")

    return
}

P6_nav_rejsesøg()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("rr")
    sleep 200
    SendInput "^t"

    return
}

; tager dato (hvis kun en søgedato), alternativt start og slutdato
P6_nav_rejsesøg_hylde(dato, datoslut?)
{
    P6_nav_rejsesøg

    SendInput "!f"
    sleep 10
    if IsSet(datoslut)
        SendInput dato "{tab 2}" datoslut
    else
        SendInput dato "{tab 2}" dato
    SendInput "!h{space}{enter}"

    return
}

P6_nav_bestilling()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("rb")

    return
}

P6_nav_kørselsaftale()
{
    P6_nav_aktiver()
    P6_nav_planbillede()
    P6_nav_alt_menu("tk")

    sleep 40
    SendInput "!{F5}"

    return

}
P6_nav_kundealarm()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("ta")

    return
}


p6_nav_udråb()
{
    p6_nav_aktiver()
    p6_nav_alt_menu("ta", "!u")

    return
}

p6_nav_tal()
{
    p6_nav_aktiver()
    p6_nav_alt_menu("ta", "!t")

    return
}

;

; går til aktive vognløbs vognløbsbillede, return true når indlæst
P6_nav_vognløbsbillede(planbillede_vognløb)
{
    P6_nav_aktiver()

    sleep 30
    SendInput "^{F12}"
    sleep 150

    ; tjek for popup, igangværende ændring i vognløbsbillede
    A_Clipboard := ""
    SendInput "^c"
    ClipWait 0.3
    if (InStr(A_Clipboard, "opdateringern"))
        SendInput "!y"

    A_Clipboard := ""
    SendInput "+{F10}c"
    ClipWait 1
    vognløbsbillede_vognløb := A_Clipboard
    while (vognløbsbillede_vognløb != planbillede_vognløb)
    {
        if (A_Index == 6)
        {
            return false
        }
        SendInput "!l"
        sleep 10
        SendInput "+{F10}c"
        ClipWait 0.3
        vognløbsbillede_vognløb := A_Clipboard
        sleep 500
    }
    return true
}

;
P6_nav_vognløbsbillede_ændr_1(planbillede_kørselsaftale)
{
    SendInput "^æ"
    A_Clipboard := ""
    sleep 40
    SendInput "^c"
    ClipWait 0.5

    ; Tjek for hop til dato ved lukket vognløb
    if (A_Clipboard != "")
        if (InStr(A_Clipboard, A_Year))
            return "lukket"

    ; tjek for færdigindlæst vognløbsbillede
    sleep 100
    A_Clipboard := ""
    SendInput "+{F10}c"
    ClipWait 0.3
    while (A_Clipboard != planbillede_kørselsaftale)
    {
        if (A_Index == 10)
            return false
        A_Clipboard := ""
        SendInput "!k"
        sleep 20
        SendInput "+{F10}c"
        sleep 500
    }
    vognløbsbillede_kørselsaftale := A_Clipboard

    ; Tjek for vogngruppevognløb
    A_Clipboard := ""
    SendInput "{tab 2}"
    sleep 60
    SendInput "+{F10}c"
    ClipWait 0.3
    if (A_Clipboard = "" or A_Clipboard == vognløbsbillede_kørselsaftale)
        vognløbsbillede_vogngruppe := false
    else
        vognløbsbillede_vogngruppe := true

    SendInput "{enter}"

    return

}

P6_nav_vognløbsbillede_ændr_2()
{

    SendInput "{enter}"

}
P6_nav_vognløbsbillede_ændr_afslut()
{

    SendInput "{enter}"

}
;; P6 indhent data

; Henter tlf fra vl hvis intet parameter, indsætter tlf på vognløb hvis der er
P6_hent_data_vognløbsbillede_telefon(telefonnummer?)
{
    {
        SendInput "{enter}!ø{tab 2}"
        sleep 20
        A_Clipboard := ""
        SendInput "+{F10}c"
        ClipWait 0.3
        while (StrLen(A_Clipboard) != 8)
        {
            if (a_index == 6)
                return false
            SendInput "!ø{tab 2}"
            sleep 20
            SendInput "+{F10}c"
            ClipWait 0.3
        }
        if IsSet(telefonnummer)
        {
            SendInput telefonnummer
            sleep 20
            SendInput "{enter}"
        }
        else
        {
            A_Clipboard := ""
            SendInput "+{F10}c"
            ClipWait 0.5
            SendInput "{enter}"
            return A_Clipboard
        }
    }
    return
}

P6_hent_data_vm_telefon()
{
    P6_nav_kørselsaftale()
    SendInput "^æ"
    sleep 40
    SendInput "!a{tab 4}"
    A_Clipboard := ""
    SendInput "^c"
    ClipWait 0.5
    while (StrLen(A_Clipboard) != 8)
    {
        if (a_index == 4)
            return false

        SendInput "!a{tab 4}"
        A_Clipboard := ""
        SendInput "^c"
        ClipWait 0.5
    }

    SendInput "^a"
    return A_Clipboard
}

p6_hent_data_rejsesøg_telefon(telefon)
{
    P6_nav_rejsesøg()
    SendInput "+{tab 2}" telefon
    sleep 100
    SendInput "{enter}"
}
; Fejlbesked ved fejl i indhentning af data, tager navn på forsøgt data som parameter
P6_hent_data_vis_fejlbesked(indhentet_data)
{
    MsgBox "Der er sket en fejl i indhentning af " indhentet_data ", prøv igen.", "Fejl", 16

    return
}

;
P6_ret_data_vognløbsbillede_ændre_sluttid(vognløb, kørselsaftale, sluttid, dato?)
{
    P6_nav_aktiver()
    P6_nav_vognløbsbillede(vognløb)
    P6_nav_vognløbsbillede_ændr_1(kørselsaftale)

    SendInput "{tab 2}"
    if IsSet(dato)
    {
        SendInput dato
        SendInput "{tab}"
        SendInput sluttid
        SendInput "{tab}"
        SendInput dato
        SendInput "{tab}"
        SendInput sluttid
    }
    else
    {
        SendInput FormatTime(, "dd")
        SendInput "{tab}"
        SendInput sluttid
        SendInput "{tab}"
        SendInput FormatTime(, "dd")
        SendInput "{tab}"
        SendInput sluttid
    }

    return

}

; Return array[4], vognløbsnummer som [1]
P6_hent_data_vognløb_vognløbsnummer()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["vognløbsnummer", "", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}
; Return array[4], vognløbsnummer som [1]
; uden at aktivere planbillede
P6_hent_data_vognløb_vognløbsnummer_i_planbillede()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    for i, e in ["vognløbsnummer", "", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], vognløbsdato som [2]
P6_hent_data_vognløb_vognløbsdato()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["", "vognløbsdato", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], vognløbsnummer som [1], vognløbsdato som [2]
P6_hent_data_vognløb_vognløbsnummer_og_vognløbsdato()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["vognløbsnummer", "vognløbsdato", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}
; Return array[4], kørselsaftale som [3], styresystem som [4]
P6_hent_data_vognløb_kørselsaftale_og_styresystem()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["", "", "kørselsaftale", "styresystem"]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], vognløbsnummer som [1], vognløbsdato som [2], kørselsaftale som [3], styresystem som [4]
P6_hent_data_vognløb_alt()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for index, data in ["vognløbsnummer", "vognløbsdato", "kørselsaftale", "styresystem"]
    {
        indhentet_data := P6_hent_data_vognløb_funk(data)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[index] := indhentet_data
            break

        }

        hent_data_vognløb_output[index] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Tager valgt 1 datatype som input, "vognløb", "vognløbsdato", "kørselsaftale", "styresystem"
; Return "fejl" hvis fejl i indhentning
; => str
P6_hent_data_vognløb_funk(valgt_data)
{
    ; [1] planetgenvej, [2] kopieringsgenvej
    hent_data_input :=
        Map(
            "vognløbsnummer", ["!l", "+{F10}c"],
            "vognløbsdato", ["!l{tab}", "^c"],
            "kørselsaftale", ["!k", "+{F10}c"],
            "styresystem", ["!k{tab}", "+{F10}c"]
        )

    hent_data_output := ""


    if (valgt_data != "")
    {
        SendInput hent_data_input[valgt_data][1]
        A_Clipboard := ""
        SendInput hent_data_input[valgt_data][2]
        ClipWait 0.5
        hent_data_output := A_Clipboard
        while (hent_data_output == "")
        {
            if (A_Index == 3 and valgt_data == "kørselsaftale")
            {
                hent_data_output := "vogngruppe"

                return hent_data_output
            }
            if (A_Index == 10 and valgt_data != "kørselsaftale")
            {
                P6_hent_data_vis_fejlbesked(valgt_data)
                hent_data_output := "fejl"

                return hent_data_output
            }
            P6_nav_aktiver()
            P6_nav_planbillede()
            SendInput hent_data_input[valgt_data][1]
            A_Clipboard := ""
            SendInput hent_data_input[valgt_data][2]
            ClipWait 0.2
            hent_data_output := A_Clipboard
        }
    }
    return hent_data_output
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
; 2. Alm svigt, GV variabel og variable
; 3. GV vl slettet
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

input_test := Map(
    "Vl_type", "[1, 2, 3]",
    ; GV-behandling
    "GV_behandling", "[0, 1, 2, 3, 4]",
    "VM_kontakt", "[1, 2]",
    "VM_tid_for_kontakt", "str_tid_vm",
    ; øvrigt
    "Tid_for_svigt", "str_tid_svigt",
    "Beskrivelse", "str_beskrivelse",
    "Emnefelt_beskrivelse", "0",
    "Emnefelt_tekst", "str_emnefelt_tekst"
)
; t::
; svigt_opret_tekst_brødtekst_gv(input_test)
; {

;     brødtekst := ""


;     if (input_test["GV_behandling"] == )
; }
; svigt_opret_tekst_brødtekst_gv_variabel()
; {

; }
; svigt_opret_tekst_brødtekst_variabel()
; {

; }
