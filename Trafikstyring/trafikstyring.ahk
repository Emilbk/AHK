#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook
#InstallMouseHook
;FileEncoding UTF-8
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode, 1 ; matcher så længe et ord er der
#SingleInstance, force
; Define the group: gruppe
GroupAdd, gruppe, PLANET
; GroupAdd, gruppe, ahk_class Chrome_WidgetWin_1
GroupAdd, gruppe, ahk_class AccessBar
GroupAdd, gruppe, ahk_class Agent Main GUI
GroupAdd, gruppe, ahk_class Addressbook
;; lib
#Include, %A_linefile%\..\lib\AHKDb\ahkdb.ahk

;; TODO

; gemt-klip-funktion ved al brug af clipboard
; Trio gå til linie 1 hvis linie 2 aktiv
; forstå pixelsearch
; Tilføj kommentar, der vises når VM ringer op

; hvis vm tlf
;     vis liste over tilknyttede vognløb, med markering for kommentar
;     vælg vl

; scratchpad, med mulighed for at liste vognløb
; tilknyt kommentar til vl (vis i oversigten hvis og hvornår)
; mulighed for timer reminder
; klik for åben vl i planet

;; kendte fejl

;; Globale variabler

brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1) ; brugerens række i databasen
bruger_genvej := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1) ; array med alle brugerens data
genvej_ren := []
genvej_navn := []
;   1       2               3
s := bruger_genvej.41
tlf :=
trio_genvej := "Genvejsoversigt"
vl_repl := []
;   bruger_genvej  telenor_opr     telenor_ahk

;; hotkeydef.
; globale genveje                                           ; Standard-opsætning
Hotkey, % bruger_genvej.4, l_trio_P6_opslag ; !w
Hotkey, % bruger_genvej.30, l_trio_afslut_opkald ; Numpad -
Hotkey, % bruger_genvej.31, l_trio_afslut_opkaldB ; Numpad -
Hotkey, % bruger_genvej.32, l_trio_til_p6 ; +F4
Hotkey, % bruger_genvej.33, l_quitAHK ; +escape
Hotkey, % bruger_genvej.46, l_restartAHK ; +escape
Hotkey, % bruger_genvej.34, l_p6_aktiver ; +!p
Hotkey, % bruger_genvej.47, l_gui_hjælp ; ^½
Hotkey, % bruger_genvej.28, l_trio_opkald_markeret ; !q

Hotkey, IfWinActive, PLANET
Hotkey, % bruger_genvej.38, l_outlook_svigt ; +F1
Hotkey, % bruger_genvej.5, l_p6_initialer ; F2
Hotkey, % bruger_genvej.6, l_p6_initialer_skriv ; +F2
Hotkey, % bruger_genvej.7, l_p6_vis_k_aftale ; F3
Hotkey, % bruger_genvej.8, l_p6_ret_vl_tlf ; +F3
Hotkey, % bruger_genvej.9, l_p6_vaelg_vl ; ^F3
Hotkey, % bruger_genvej.10, l_p6_vaelg_vl ; F4
Hotkey, % bruger_genvej.11, l_p6_vl_ring_op ; +F5
Hotkey, % bruger_genvej.12, l_p6_vm_ring_op ; ^+F5
Hotkey, % bruger_genvej.13, l_p6_vl_luk ; #F5
Hotkey, % bruger_genvej.14, l_p6_alarmer ; F7
Hotkey, % bruger_genvej.15, l_p6_udraabsalarmer ; +F7
Hotkey, % bruger_genvej.16, l_p6_ring_til_kunde ; +F8
Hotkey, % bruger_genvej.17, l_p6_udregn_minut ; #t
Hotkey, % bruger_genvej.18, l_p6_sygehus_ring_op ; ^+s
Hotkey, % bruger_genvej.19, l_p6_central_ring_op ; ^+c
Hotkey, % bruger_genvej.20, l_p6_tekst_til_chf ; ^+t
Hotkey, % bruger_genvej.36, l_flexf_fra_p6 ; +^F
Hotkey, % bruger_genvej.48, l_p6_rejsesog ; F1
Hotkey, % bruger_genvej.50, l_p6_liste_vl ; F1
Hotkey, % bruger_genvej.51, l_p6_vis_liste_vl ; F1
; Hotkey, % bruger_genvej.45, l_sys_inputbox_til_fra ; ^½
Hotkey, IfWinActive

Hotkey, IfWinActive, Planet ; specifikt alarmrepl-infobox
Hotkey, % bruger_genvej.49, l_p6_replaner ; F1
Hotkey, IfWinActive
; Trio
Hotkey, IfWinActive, ahk_group gruppe
Hotkey, % bruger_genvej.22, l_trio_pause ; ^0
Hotkey, % bruger_genvej.23, l_trio_klar ; ^1
Hotkey, % bruger_genvej.24, l_trio_udenov ; ^2
Hotkey, % bruger_genvej.25, l_trio_efterbehandling ; ^3
Hotkey, % bruger_genvej.26, l_trio_alarm ; ^4
Hotkey, % bruger_genvej.27, l_trio_frokost ; ^5
Hotkey, % bruger_genvej.29, l_triokald_til_udklip ; #q
Hotkey, IfWinActive

; flexfinder
Hotkey, IfWinActive, FlexDanmark FlexFinder ;
Hotkey, % bruger_genvej.35, l_flexf_til_p6 ; ~$^LButton
Hotkey, IfWinActive, ,
; outlook
Hotkey, % bruger_genvej.37, l_outlook_ny_mail ; ^+m

Hotkey, IfWinActive, PLANET
Hotkey, IfWinActive, ,

;excel
Hotkey, ifWinActive, Garantivognsoversigt FG8.xlsm
Hotkey, % bruger_genvej.39, l_excel_vl_til_P6_A ; !Lbutton
Hotkey, % bruger_genvej.40, l_excel_vl_til_P6_B ; ^w
Hotkey, IfWinActive, ,

;; GUI
; gui-definitioner

; Ring til sygehus
gui sygehus:+Labelsygehus
Gui sygehus: Font, s9, Segoe UI
Gui sygehus: Add, Button, gsygehusmenu1 vauh x16 y8 w115 h23, &AUH
Gui sygehus: Add, Button, gsygehusmenu1 vrhg x16 y32 w115 h23, RH&G
Gui sygehus: Add, Button, gsygehusmenu1 vrand x16 y56 w115 h23, &Randers Sygehus
Gui sygehus: Add, Button, gsygehusmenu1 vvib x16 y80 w115 h23, &Viborg Sygehus
Gui sygehus: Add, Button, gsygehusmenu1 vhor x16 y104 w115 h23, &Horsens Sygehus
Gui sygehus: Add, Button, gsygehusmenu1 vsil x16 y128 w115 h23, &Silkeborg Sygehus
Gui sygehus: Add, Button, gsygehusmenu1 vpsyk x16 y152 w115 h23, &Psyk
Gui sygehus: Add, Button, gsygehusmenu1 vmisc x16 y176 w115 h23, Andr&e

gui sygehusauh:+Labelsygehus2
Gui sygehusauh: Font, s9, Segoe UI
Gui sygehusauh: Add, Button, gsygehusmenu2 v78450000 x16 y8 w115 h23, &AUH syg.
Gui sygehusauh: Add, Button, gsygehusmenu2 v78452501 x16 y32 w115 h23, &Dialyse
Gui sygehusauh: Add, Button, gsygehusmenu2 v78454955 x16 y56 w115 h35, &Kræ.amb. og Kemo
Gui sygehusauh: Add, Button, gsygehusmenu2 v78454921 x16 y92 w115 h23, &Stråleterapi
Gui sygehusauh: Add, Button, gsygehusmenu2 v78454931 x16 y116 w115 h23, Kræft &sengeafsnit
Gui sygehusauh: Add, Button, gsygehusmenu2 v78454114 x16 y140 w115 h23, &Ortopædkir.
Gui sygehusauh: Add, Button, gsygehusmenu2 v78471000 x16 y164 w115 h23, &Psyk.

gui sygehuspsyk:+Labelsygehus2
Gui sygehuspsyk: Font, s9, Segoe UI
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78471000 x16 y8 w115 h23, &AUH psyk.
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78474500 x16 y32 w115 h23, &RHG psyk.
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78475300 x16 y56 w115 h30, &Randers psyk.
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v20936488 x16 y87 w115 h30, &Holstebro psyk.
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78474000 x16 y111 w115, &Viborg, Silkeborg og Skive psyk.
; Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78474000 x16 y111 w115 h23, &Silkeborg psyk.
; Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78474000 x16 y135 w115 h23, &Psyk.

gui sygehusrhg:+Labelsygehus2
Gui sygehusrhg: Font, s9, Segoe UI
Gui sygehusrhg: Add, Button, gsygehusmenu2 v78430000 x16 y8 w115 h23, RH&G syg.
Gui sygehusrhg: Add, Button, gsygehusmenu2 v78436760 x16 y32 w115 h23, &Dialyse
Gui sygehusrhg: Add, Button, gsygehusmenu2 v78474500 x16 y56 w115 h23, &Psyk.
Gui sygehusrhg: Add, Button, gsygehusmenu2 v78437463 x16 y80 w115 h23, H&erning Stråle.

gui sygehusrand:+Labelsygehus2
Gui sygehusrand: Font, s9, Segoe UI
Gui sygehusrand: Add, Button, gsygehusmenu2 v78420000 x16 y8 w115 h23, Randers syg.
Gui sygehusrand: Add, Button, gsygehusmenu2 v78421590 x16 y32 w115 h23, &Dialyse
Gui sygehusrand: Add, Button, gsygehusmenu2 v78475300 x16 y56 w115 h23, &Psyk.

gui sygehusvib:+Labelsygehus2
Gui sygehusvib: Font, s9, Segoe UI
Gui sygehusvib: Add, Button, gsygehusmenu2 v78430000 x16 y8 w115 h23, &Viborg syg.
Gui sygehusvib: Add, Button, gsygehusmenu2 v78447720 x16 y32 w115 h23, &Dialyse
Gui sygehusvib: Add, Button, gsygehusmenu2 v78474000 x16 y56 w115 h23, &Psyk.

gui sygehussil:+Labelsygehus2
Gui sygehussil: Font, s9, Segoe UI
Gui sygehussil: Add, Button, gsygehusmenu2 v78415000 x16 y8 w115 h23, &Silkeborg syg.
Gui sygehussil: Add, Button, gsygehusmenu2 v78474000 x16 y32 w115 h23, &Psyk.
; Gui sygehussil: Add, Button, gsygehusmenu2 vnogetandet x16 y56 w115 h23, &Noget andet
; Gui sygehussil: Add, Button, gsygehusmenu2 x16 y80 w115 h23, &Vi
; Gui sygehussil: Add, Button, gsygehusmenu2 x16 y104 w115 h23, &o
; Gui sygehussil: Add, Button, gsygehusmenu2 x16 y128 w115 h23, &S

gui sygehushor:+Labelsygehus2
Gui sygehushor: Font, s9, Segoe UI
Gui sygehushor: Add, Button, gsygehusmenu2 v78425000 x16 y8 w115 h23, &Horsens syg.
Gui sygehushor: Add, Button, gsygehusmenu2 v78426160 x16 y32 w115 h23, &Dialyse
Gui sygehushor: Add, Button, gsygehusmenu2 v78425871 x16 y56 w115 h23, &Røntg. før 09
; Gui sygehushor: Add, Button, gsygehusmenu2 x16 y80 w115 h23, &Vi
; Gui sygehushor: Add, Button, gsygehusmenu2 x16 y104 w115 h23, &o
; Gui sygehushor: Add, Button, gsygehusmenu2 x16 y128 w115 h23, &S

gui sygehusmisc:+Labelsygehus2
Gui sygehusmisc: Font, s9, Segoe UI
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78425000 x16 y8 w115 h23, &Brædstrup
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78420000 x16 y32 w115 h23, &Grenå
Gui sygehusmisc: Add, Button, gsygehusmenu2 v20936488 x16 y56 w115 h35, &Holstebro` Psykiatrien
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78437463 x16 y92 w115 h23, H&erning Stråle.
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78419000 x16 y116 w115 h23, Hammel &Neuro.
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78430000 x16 y140 w115 h23, &Lemvig og Tarm
Gui sygehusmisc: Add, Button, gsygehusmenu2 v30463689 x16 y164 w115 h23, &Samsø
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78440000 x16 y188 w115 h23, &Skive
Gui sygehusmisc: Add, Button, gsygehusmenu2 v99157302 x16 y212 w115, Skive - &Fys (efter 12.00)
Gui sygehusmisc: Add, Button, gsygehusmenu2 v78474000 x16 y254 w115 h23, Skive - &Psyk

; Trio_tlf_knap
; Trio_tlf_knap
Gui tlf: +Labeltlf
Gui tlf: -MinimizeBox -MaximizeBox +AlwaysOnTop +Owner -Caption +ToolWindow +hwndhGui
Gui tlf: Font, s12, Segoe UI
Gui tlf: Add, Button, vtlfKopi gtlfKopi x0 y0 w120 h23, Tlf: %tlf_knap%

Gui tlf: Show, x995 y3 w120 h23 NA, Tlf

; Trio-genvej
Gui trio_genvej: +Labeltrio_genvej
Gui trio_genvej: -MinimizeBox -MaximizeBox +AlwaysOnTop +Owner -Caption +ToolWindow +hwndhGui
Gui trio_genvej: Font, s12, Segoe UI
Gui trio_genvej: Add, Button, vtrio_genvej gtrio_genvej x0 y0 h42 w240, %trio_genvej%

; Gui trio_genvej: Show, x1120 y3 w120 h42 w240 NA, %trio_genvej%

; gui repl
Gui repl: Font, s9, Segoe UI
Gui repl: Add, ListBox, Choose1 x78 y21 w120 h364 vvalg, %vl_repl_liste%
Gui repl: Add, Button, x359 y239 w80 h23 Default greplok, &OK
Gui repl: Add, Button, x359 y270 w80 h23 greplslet, &Slet
; Gui repl: Show, w620 h420, Window

;; end autoexec
return
;; GUI-labels
trio_genvej:
    ; Goto, l_gui_hjælp
    MsgBox, , Tillykke!, Du har trykket på knappen!,

tlfKopi:
    {
        clipboard :=
        tlf := Trio_hent_tlf()
        Clipboard := tlf
        ClipWait, 3,
        return
    }

sygehusmenu1:
    GuiControlGet, navn, sygehus: name , % A_GuiControl
    ; navn := SubStr(navn, 2)
    ; MsgBox, , , % navn
    vis_sygehus_2(navn)
    gui cancel
return

sygehusmenu2:
    GuiControlGet, knap2, sygehus%navn%: name, % a_guicontrol
    ; MsgBox, , navn2 på knap , % knap2
    Trio_opkald(knap2)
    gui cancel
    WinActivate, PLANET
    afslut_genvej()
return

sygehusEscape:
sygehusClose:
    gui cancel
    afslut_genvej()
return

sygehus2Escape:
sygehus2Close:
    vis_sygehus_1()
    gui Cancel
    afslut_genvej()
return

; GUI repl
replguiEscape:
replguiClose:
    gui, hide
    return

replOK:
    Gui, Submit
    gui, hide
    p6_vaelg_vl(valg)
; MsgBox, , , % valg
return

replslet:
    Gui, Submit
    gui, hide
    for k, v in vl_repl
        if (valg = v)
            vl.Pop(k)
return

replvl:
    gui Submit
    gui Hide
    P6_planvindue()
    p6_vaelg_vl(%valg%)
return

;; FUNKTIONER
;; P6

afslut_genvej()
{
    GuiControl, trio_genvej:text, Button1, Genvejsoversigt
    mod_up()
    return
}

genvej_beskrivelse(kolonne)
{
    trio_genvej := databaseget("%A_linefile%\..\db\bruger_ops.tsv", 3, kolonne)
    GuiControl, trio_genvej:text, Button1, %trio_genvej%
    return trio_genvej
}
; **
; fix, giver 0-fejl ved esc.
P6_hastighed()
{
    global s
    global brugerrække
    keywait, shift
    InputBox, s, P6-hastighed, Hastighed fra 1-3? `n 1 = hurtig (standard)`, 3 = meget langsom`, kommatal f. eks. = 1.5.`n `n Er nu: %s%
    if (s = "" or s = "0")
    {
        sleep 400
        MsgBox, , Fejl, Kan ikke være nul eller intet.
        return
    }
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 41, s)
    return
}

; True hvis skiftet til P6
P6_aktiver()
{
    IfWinNotActive, PLANET
    {
        WinActivate, PLANET
        WinWaitActive, PLANET
        sleep 100
        SendInput, {esc} ; registrerer ikke første tryk, når der skiftes til vindue
        sleep 300
        return 1
    }
    return 0
}

P6_alt_menu(byref tast1 := "", byref tast2 := "")
{
    ; keywait ctrl, T0.5
    ; keywait alt, T0.5
    sleep 40
    SendInput, %tast1%
    SendInput, %tast2%
    sleep 200
}
; ***
; Åben planbillede
P6_planvindue()
{
    global s
    P6_alt_menu("{alt}", "tp")
}

; ***
; Åben renset rejsesøg
P6_rejsesogvindue(byref telefon := "")
{
    global s

    P6_alt_menu("{alt}", "rr")
    sleep s * 100
    if (telefon = "")
        return
    SendInput, ^t
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, ^r
    Return
}

;  ***
; Vis kørselsaftale for aktivt planbillede
P6_vis_k()
{
    global s
    P6_planvindue()
    P6_alt_menu("!tk")
    sleep s * 40
    SendInput !{F5}
    return
}
; ***
;Kørselsaftale på VL til clipboard
P6_hent_k()
{
    global s
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    Sendinput !tp!k
    clipboard := ""
    Sendinput +{F10}c
    Send, {Ctrl}
    ClipWait
    sleep s * 200
    kørselsaftale := clipboard
    return kørselsaftale
}
;udfyld kørselsaftale
P6_udfyld_k(k:="")
{
    global s
    clipboard := k
    P6_Planvindue()
    sleep s * 40
    SendInput, !k
    sleep 40
    SendInput, {AppsKey}P
    sleep 40
    SendInput, {Enter}
}
; ***
;styresystem til clipboard
P6_hent_s()
{
    global s
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    Sendinput !tp!k{tab}
    clipboard := ""
    Sendinput +{F10}c
    ClipWait
    styresystem := clipboard
    return styresystem
}

P6_udfyld_s(ss:="")
{
    global s
    clipboard := ss
    P6_Planvindue()
    sleep s * 40
    SendInput, !k
    sleep 40
    SendInput, {tab}
    sleep 40
    SendInput, {AppsKey}P
    sleep 40
    SendInput, {Enter}

    return
}
; Hent VL-nummer
P6_hent_vl()
{
    global s
    clipboard := ""
    P6_planvindue()
    SendInput, !l
    sleep 100 ; ikke P6-afhængig
    SendInput, +{F10}c
    ClipWait, 2, 0
    vl := Clipboard
    return vl
}

p6_vl_vindue()
{
    P6_planvindue()
    sleep 30
    vl := P6_hent_vl()
    sleep 30
    SendInput, ^{F12}
    sleep 250
    clipboard :=
    SendInput, ^c
    clipwait 0.5
    if (InStr(clipboard, "opdateringern"))
    {
        SendInput, !y
    }
    clipboard :=
    vl_opslag := clipboard
    tid_start := A_TickCount
    while (vl_opslag != vl)
    {
        Send, +{F10}c
        vl_opslag := clipboard
        sleep 100
        tid_nu := A_TickCount - tid_start
        if (tid_nu > 12000)
        {
            return 0
        }
    }
    vl := clipboard
    return vl
}

p6_vl_vindue_edit()
{
    k_aftale := []
    gemtklip := ClipboardAll

    sendinput ^æ
    clipboard :=
    SendInput, ^c
    clipwait 0.5
    if (InStr(clipboard, A_Year))
    {
        return 1 ; VL lukket
    }
    clipboard :=
    SendInput, +{F10}c
    clipwait 2
    k_aftale.1 := clipboard
    clipboard :=
    SendInput, {tab 2}
    sleep 40
    SendInput, +{F10}c
    clipwait 0.5
    k_aftale.2 := clipboard
    if (k_aftale.1 = k_aftale.2)
    {
        k_aftale.2 := "drift"
    }
    clipboard := gemtklip
    gemtklip :=
    return k_aftale
}
P6_udfyld_vl(vl:="")
{
    global s
    P6_Planvindue()
    sleep s * 40
    SendInput, !l
    sleep s * 200
    SendInput, %vl%
    sleep s * 40
    SendInput, {Enter}
    return
}

; vælg vl, tager VL som parameter
p6_vaelg_vl(byref vl := "")
{
    P6_Planvindue()
    SendInput, !l
    if (vl != "")
    {
        SendInput, %vl%
        sleep 100
        SendInput, {enter}
    }
    return
}

; Udfyld kørselsaftale og styresystem, tager vl(array) som parameter. Kørselaftale = vl.1, Styresystem = vl.2
P6_udfyld_k_og_s(vl:="")
{
    global s
    P6_Planvindue()
    sleep s * 40
    SendInput, !k
    SendInput, {BackSpace} ; ved tp til udfyldt VL er første tastetryk lig med delete
    SendInput, % vl.1
    sleep s * 100
    SendInput, {tab}
    sleep s * 100
    SendInput, % vl.2
    sleep s * 100
    SendInput, {Enter}
    return
}

; ***
; åben alarmvinduet, ny liste alle alarmer, blad til første, col 14
P6_alarmer()
{
    global s

    P6_alt_menu("!ta!k")
    SendInput, ^{up}
    SendInput, +^{Down}
    sleep s * 200
    SendInput, ^{delete}
    SendInput, ^l
    SendInput, ^{up}
    sleep 100 + s * 10
    SendInput, ^{F10}

    return
}

; ***
; åben alarmvinduet, ny liste alle udråbsalarmer, blad til første, col 15
P6_udraabsalarmer()
{
    global s

    P6_alt_menu("!ta!u")
    sleep s * 200
    SendInput, ^{Delete}
    SendInput, ^{Up}
    SendInput, +^{Down}
    sleep s * 40
    SendInput, ^l
    SendInput, ^{Up}
    SendInput, ^{F10}

    return
}

P6_notat(byref tekst:="")
{
    P6_planvindue()
    SendInput, ^n
    sleep 500
    SendInput, %tekst%
    SendInput, !o

    Return

}
; ***
; gå i rent rejsesøg med karet i telefonfelt
P6_rejsesog_tlf(ByRef telefon:=" ")
{
    global s
    P6_rejsesogvindue()
    sleep s * 300
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, ^r

    Return
}
; ***
;

P6_hent_vl_tlf()
{

    global s
    gemt_klip := clipboard

    vl_tilstand := p6_vl_vindue()
    if (vl_tilstand = 0)
    {
        sleep 100
        MsgBox, , For lang tid brugt, Noget er gået galt. Prøv igen.
        afslut_genvej()
        return 0
    }
    vl_tilstand := p6_vl_vindue_edit()
    if (vl_tilstand = 0)
    {
        sleep 100
        MsgBox, , Vl er lukket, Kan ikke trække telefonnummer, vl er afsluttet
        afslut_genvej()
        return 0
    }
    SendInput {Enter}{Enter}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    Clipboard :=
    SendInput {tab}{tab}
    while (StrLen(clipboard) != 8)
    {
        clipboard :=
        SendInput ^c
        ClipWait, 1
        sleep 100
    }
    SendInput {enter}
    vl_tlf := Clipboard
    clipboard := gemt_klip
    Return vl_tlf

}
; ***
; P6 hent VM tlf
P6_hent_vm_tlf()
{
    gemtklip := clipboard
    global s
    P6_vis_k()
    sleep s * 40
    sendinput ^æ
    sleep s * 40
    SendInput !a
    ; sleep * 40
    Clipboard :=
    SendInput {tab}{tab}{tab}{tab}
    while (StrLen(clipboard) != 8)
    {
        clipboard :=
        SendInput ^c
        ClipWait, 1
        sleep 100
    }
    SendInput, {enter}
    SendInput ^a
    vm_tlf := Clipboard
    clipboard := gemtklip
    Return vm_tlf
}

; l_telenor_p6_opslag
P6_hent_vl_fra_tlf(ByRef tlf:="")
{
    if (tlf = "")
    {
        return 0
    }
    række := DataBasefind( "%A_linefile%\..\db\VL_tlf.txt", tlf)
    vl := databaseget("%A_linefile%\..\db\VL_tlf.txt", række.1, 2)
    if (række.1 is number) ; hvorfor virker den ikke med true/false?
    {
        vl := StrSplit(vl, "_") ;vl.1 k, vl.2 s
        Return vl
    }
    return 0
}

; ***
;indsæt clipboard i vl-tlf
P6_ret_tlf_vl(ByRef telefon:=" ")
{
    p6_vl_vindue()
    sleep 100
    sendinput ^æ
    sleep s * 200
    SendInput {Enter}{Enter}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, {enter}
    return
}
; ***
;indsæt clipboard i vl-tlf
P6_ret_tlf_vl_efterfølgende(ByRef telefon:="")
{
    sendinput ^æ
    sleep s * 200
    SendInput {Enter}{Enter}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, {enter}
    return
}

;  ***
;indsæt clipboard i vl-tlf dagen efterfølgende
P6_tlf_vl_dato_efter(ByRef telefon:=" ")
{
    global s
    SendInput, {Tab}
    sleep s * 200
    SendInput, !{right}{AltUp}
    sleep s * 200
    SendInput, ^æ
    sleep s * 200
    SendInput {Enter 2}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, {enter}

    return
}
; ***
; Omskriv til simplere funktion
; Noterer intialer, fjerner dem hvis første ord i notering er initialer
P6_initialer()
{
    global s
    initialer := sys_initialer()
    initialer_udentid := "/mt" A_userName 

    SendInput, {F5} ; for at undgå timeout. Giver det problemer med langsom opdatering?
    sleep s * 40
    sendinput ^n
    sleep s * 1400
    clipboard :=
    SendInput, ^a^c
    ClipWait, 1, 0
    notering := Clipboard
    clipwait 3, 0
    if (substr(notering,1, 6) = initialer_udentid)
    {
        initialer_fjernet := SubStr(notering, 12)
        If (initialer_fjernet) = ""
            initialer_fjernet := " "
        Clipboard :=
        sleep 100
        Clipboard := initialer_fjernet
        ClipWait, 1, 0
        sendinput ^a^v
        sleep s * 200
        SendInput, !o
        return
    }
    if (substr(notering,1, 6) != initialer_udentid)
    {
        Clipboard :=
        sleep s * 40
        clipboard := initialer
        ClipWait, 1, 0
        SendInput, {Left}
        Sendinput ^v
        SendInput, %A_Space%
        sleep s * 100
        SendInput, !o
        return
    }
}

; ** kan gemtklip-funktion skrives bedre?
;Indsæt initialer med efterf. kommentar, behold tidligere klip
P6_initialer_skriv()
{
    global s
    initialer := sys_initialer()
    P6_Planvindue()
    sleep s * 40
    sendinput ^n
    sleep s * 40
    Sendinput %initialer%
    Sendinput %A_space%
    Sendinput {home}
    sleep 2000 ; ikke P6-afhængig
    ; gemtklip := ""
    return
}

;  ***
; Send tekst til chf
P6_tekstTilChf(ByRef tekst:=" ")
{
    global s
    P6_planvindue()
    kørselsaftale := P6_hent_k()
    styresystem := P6_hent_s()
    sleep s * 200
    Sendinput !tt^k
    sleep s * 100
    Sendinput !k
    sleep s * 40
    SendInput, ^t
    Sendinput %kørselsaftale%
    sleep s * 100
    SendInput, {tab}
    sleep 40
    Sendinput %styresystem%
    SendInput, {tab}
    sleep s * 100
    if (tekst != " ")
    {
        SendInput, %tekst%
    }
    Else
        return
    return
}

; ***
; Finder lukketid ud fra sidste stop og tid til hjemzone.
; Input tid for sidste stop, tryk enter. Input tid til hjemzone, tryk enter.
; Hvis tid for sidste stop hjemzone er tom, luk nu + 5 min
; hvis tid til hjemzone stop tom luk til udfyldte tid for sidste stop uden ændringer
; hvis tid for sidste stop og tid til hjemzone udfyldt, luk til tiden fra sidste stop til hjemzone, plus 2 min
P6_input_sluttid()
{
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    p6_input_sidste_slut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,42)
    KeyWait, Ctrl,
    KeyWait, Shift,
    EnvAdd, nu_plus_5, 5, minutes
    FormatTime, nu_plus_5, %nu_plus_5%, HHmm
    if (p6_input_sidste_slut_ops = "1")
    {
        sleep 100
        InputBox, sidste_stop, Sidste stop, Tast tid for sidste stop (4 cifre)
        if (ErrorLevel = "1")
            Return 0
        if (sidste_stop = "")
        {
            return nu_plus_5
        }
        if (StrLen(sidste_stop)!= 4)
        {
            MsgBox, , Fejl i indtastning, Der skal bruges fire tal, i formatet TTMM (f. eks. 1434).
            return 0
        }
        sidste_stop_tjek := A_YYYY A_MM A_DD sidste_stop
        if sidste_stop_tjek is not Time
        {
            MsgBox, , Fejl i indtastning , Det indtastede er ikke et klokkeslæt.,
            return 0
        }
        sidste_stop := A_YYYY A_MM A_DD sidste_stop
        sleep 200
        InputBox, tid_til_hjemzone, Tid til hjemzone, Tid til hjemzone i minutter
        if (ErrorLevel = "1")
            Return 0
        if (tid_til_hjemzone = "" )
        {
            FormatTime, sidste_stop, %sidste_stop%, HHmm
            return sidste_stop
        }
        EnvAdd, sidste_stop, tid_til_hjemzone + 5, minutes
        FormatTime, sidste_stop, %sidste_stop%, HHmm
        return sidste_stop
    }
    if (p6_input_sidste_slut_ops = "0")
    {
        luk := [] 
        Input, sidste_stop, T10, {Enter}{escape}
        if (ErrorLevel = "EndKey:Escape")
            Return 0
        if (ErrorLevel = "Timeout")
        {MsgBox, , Timeout , Det tog for lang tid.
            return 0
        }
        if (sidste_stop = "")
        {
            return nu_plus_5
        }
        luk.Push(sidste_stop)
        if (!InStr(luk.1, "/"))
            {
                luk.3 := "luk"
            }
        if (InStr(luk.1, "/"))
            {
                luk := StrSplit(sidste_stop, "/")
                luk.3 := "åbnluk"
                if (luk.2 = "")
                    {
                    luk.3 := "åbn"
                    }
            }
        if (StrLen(luk.1) != 4)
        {
            MsgBox, , Fejl i indtastning, Der skal bruges fire tal, i formatet TTMM (f. eks. 1434).
            return 0
        }
        if (luk.2 != "" and StrLen(luk.2)!= 4)
        {
            MsgBox, , Fejl i indtastning, Der skal bruges fire tal i luktid, i formatet TTMM (f. eks. 1434).
            return 0
        }
 
        sidste_stop_tjek := A_YYYY A_MM A_DD luk.1
        if sidste_stop_tjek is not Time
        {
            MsgBox, , Fejl i indtastning , Det indtastede er ikke et klokkeslæt.,
            return 0
        }
        sidste_stop_tjek := A_YYYY A_MM A_DD luk.2
        if sidste_stop_tjek is not Time
        {
            MsgBox, , Fejl i indtastning , Den indtastede lukketid er ikke et klokkeslæt.,
            return 0
        }
 
        luk.1 := A_YYYY A_MM A_DD luk.1
        Input, tid_til_hjemzone, T5, {enter}{Escape},
        if (ErrorLevel = "EndKey:Escape")
            Return
        if (ErrorLevel = "Timeout")
        {MsgBox, , Timeout , Det tog for lang tid.
            return 0
        }
        if (tid_til_hjemzone = "" )
        {
            ; formattime kan ikke tage array?
            midl_luk := luk.1
            FormatTime, midl_luk, %midl_luk%, HHmm
            luk.RemoveAt(1)
            luk.InsertAt(1, midl_luk)
            return luk
        }
        midl_luk := luk.1
        EnvAdd, midl_luk , tid_til_hjemzone + 5, minutes
        FormatTime, midl_luk, %midl_luk%, HHmm
        luk.1 := midl_luk
        return luk
    }
}

; skal kun sende, hvis der er en tom køreordre.
; P6_send_slut()
; {
;     P6_planvindue()
;     SendInput, ^{F11}
;     sleep 100
;     clipboard :=
;     SendInput, ^c
;     ClipWait, 1
;     if (clipboard = "")
;     {
;         MsgBox, , , ja,
;     }
;     SendInput, !s{F5}
;     return
; }

; ^e::P6_send_slut()

; læg minuttal til klokkeslæt eller minuttal til minuttal.
P6_udregn_minut()
{
    resultat := []
    tidA := ; HHmm, starttid. Enten fire cifre for klokkeslæt, mellem 1 og 3 cifre for minuttertal.
    tidB := ; mm, tillægstid. Minuttal
    tidC := ; resultat
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    p6_udregn_minut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,44)
    if (p6_udregn_minut_ops = 1)
    {
        sleep 100
        InputBox, tidA, Udgangspunkt, Skriv tiden`, der skal lægges noget til. `nKlokkeslæt: 4 cifre ud i ét`, minuttal: 3 til 1 ciffer ud i ét. `n `n F. eks: `n Klokken 13:34 skrives 1334 `n 231 minutter skrives 231, , , 240
        if (ErrorLevel != 0)
            return "fejl"
        if (tida = "")
            tida := "0"
        sleep 100
        InputBox, tidB, Tid `, der skal lægges til., Skriv tid`, der skal lægges til. Minuttal ud i ét (- foran`, hvis der skal trækkes fra).,
        if (ErrorLevel != 0)
            return "fejl"
        if (tidb = "")
            tidb := "0"
        if (tidb + tida < 0)
        {
            MsgBox, , Lad vær', ,
            return
        }
        if (StrLen(tida) <= "3")
        {
            tid_nul := A_YYYY A_MM A_DD "00" "00"
            EnvAdd, tid_nul, tida , minutes
            EnvAdd, tid_nul, tidb, minutes
            FormatTime, tidC, %tid_nul%, HHmm
            FormatTime, tid_time, %tid_nul%, H
            FormatTime, tid_min, %tid_nul%, m
            FormatTime, result_mid, %tid_nul%, HHmm
            if (tid_time = "0" and tid_min = "1")
            {
                resultat.1 := tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "0" and tid_min >= "1")
            {
                resultat.1 := tid_min " minutter."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "1" and tid_min = "0")
            {
                resultat.1 := tid_time " time."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time > "1" and tid_min = "0")
            {
                resultat.1 := tid_time " timer."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "1" and tid_min = "1")
            {
                resultat.1 := tid_time " time og " tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time >= "1" and tid_min = "1")
            {
                resultat.1 := tid_time " timer og " tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time >= "1" and tid_min >= "1")
            {
                resultat.1 := tid_time " timer og " tid_min " minutter."
                resultat.2 := result_mid
                return resultat
            }
        }
        if (StrLen(tida) = "4")
        {
            tidA := A_YYYY A_MM A_DD tida "00"
            EnvAdd, tidA, tidB, minutes
            FormatTime, tid_time, %tidA%, HH
            FormatTime, tid_min, %tidA%, mm
            FormatTime, result_mid, %tida%, HHmm
            if (tid_time != "00")
            {
                resultat.1 := tid_time ":" tid_min "."
                resultat.2 := result_mid
                return resultat
            }
        }
        return
    }
    if (p6_udregn_minut_ops = 0)
    {
        Input, tida, E , {enter},
        if (ErrorLevel != "Endkey:Enter")
            return
        if (tida = "")
            tida := "0"
        Input, tidb,, {enter}
        if (ErrorLevel != "Endkey:Enter")
            return
        if (tidb = "")
            tidb := "0"
        if (tidb + tida < 0)
        {
            MsgBox, , Skal være et tidspunkt, ,
            return
        }
        if (StrLen(tida) <= "3")
        {
            tid_nul := A_YYYY A_MM A_DD "00" "00"
            EnvAdd, tid_nul, tida , minutes
            EnvAdd, tid_nul, tidb, minutes
            FormatTime, tidC, %tid_nul%, HHmm
            FormatTime, tid_time, %tid_nul%, H
            FormatTime, tid_min, %tid_nul%, m
            FormatTime, result_mid, %tid_nul%, HHmm
            if (tid_time = "0" and tid_min = "1")
            {
                resultat.1 := tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "0" and tid_min >= "1")
            {
                resultat.1 := tid_min " minutter."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "1" and tid_min = "0")
            {
                resultat.1 := tid_time " time."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time > "1" and tid_min = "0")
            {
                resultat.1 := tid_time " timer."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time = "1" and tid_min = "1")
            {
                resultat.1 := tid_time " time og " tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time >= "1" and tid_min = "1")
            {
                resultat.1 := tid_time " timer og " tid_min " minut."
                resultat.2 := result_mid
                return resultat
            }
            if (tid_time >= "1" and tid_min >= "1")
            {
                resultat.1 := tid_time " timer og " tid_min " minutter."
                resultat.2 := result_mid
                return resultat
            }
        }
        if (StrLen(tida) = "4")
        {
            tidA := A_YYYY A_MM A_DD tida "00"
            EnvAdd, tidA, tidB, minutes
            FormatTime, tid_time, %tidA%, HH
            FormatTime, tid_min, %tidA%, mm
            FormatTime, result_mid, %tida%, HHmm
            if (tid_time != "00")
            {
                resultat.1 := tid_time ":" tid_min "."
                resultat.2 := result_mid
                return resultat
            }
        }
        return
    }
}
; luk vl på variabel tid
P6_vl_luk(tid:="")
{
    global s

    vl := p6_vl_vindue()
    k_aftale := p6_vl_vindue_edit()
    sleep 40
    if (k_aftale.1 = 1)
    {
        MsgBox, , VL afsluttet, VL er allerede afsluttet
        SendInput, ^a
        afslut_genvej()
        return 0
    }
    if (k_aftale.2 = "drift" or k_aftale.2 = "" and tid.3 = "luk")
    {
        SendInput, {Enter}{Tab 3}
        SendInput, % tid[1]
        SendInput, {tab}{tab}
        SendInput, % tid[1]
        SendInput, {enter}{enter}
        return
    }
    if (k_aftale.2 = "drift" or k_aftale.2 = "" and tid.3 = "åbnluk")
    {
        SendInput, {Enter}{Tab}
        SendInput, % tid.1
        SendInput, {tab 2}
        SendInput, % tid.2
        SendInput, {tab 2}
        SendInput, % tid.2
        SendInput, {enter}{enter}
        return
    }
    if (k_aftale.2 = "drift" or k_aftale.2 = "" and tid.3 = "åbn")
    {
        SendInput, {Enter}{Tab}
        SendInput, % tid.1
        SendInput, {enter}{enter}
        return
    }
    if (k_aftale.2 != "") ; Skandstat er 7-serien. Ingen driftsVL?
    {
        SendInput, {Enter}
        FormatTime, dato, YYYYMMDDHH24MISS, d
        SendInput, %dato%
        SendInput, {tab}
        SendInput, % tid[1]
        SendInput, {enter}{enter}
        return
    }
    Else ; bruges ikke
    {
        SendInput, {Enter}{Tab 3}
        SendInput, % tid.1
        SendInput, {tab}{tab}
        SendInput, % tid.1
        SendInput, {enter}{enter}
        return
    }
}

; P6 ring op til markeret kunde i VL (telefon i bestilling)
p6_hent_kunde_tlf(ByRef telefon:="")
{
    global s

    SendInput, {enter}
    sleep s * 300
    SendInput, +{tab 2}
    sleep s * 100
    clipboard :=
    SendInput, ^c
    ClipWait, 3,
    telefon := clipboard
    return
}

p6_replaner_gem_vl()
{
    gemtklip := ClipboardAll
    ; global vl_repl
    ; global vl_repl_liste
    clipboard :=
    SendInput, ^c
    clipwait 2
    repl_besked := StrSplit(clipboard, " ")
    SendInput, {enter}
    if (repl_besked.MaxIndex() = 11)
        vl := repl_besked.6
    ; vl_repl.Push(repl_besked.6)
    if (repl_besked.MaxIndex() = 12)
        vl := repl_besked.7
    ; vl_repl.Push(repl_besked.7)

    clipboard := gemtklip
    return vl

}

p6_liste_vl(byref vl := "")
{
    gemtklip := ClipboardAll
    global vl_repl
    global vl_repl_liste

    vl_repl.Push(vl)
    vl_repl_liste := "|"
    for k, v in vl_repl
        vl_repl_liste .= vl_repl[k] . "|"

    clipboard := gemtklip
    return vl_repl

}
#IfWinActive, Planet ; er ikke med stort i repl.vindue
    +enter::
        {
            vl := p6_replaner_gem_vl()
            p6_liste_vl(vl)
            Return
        }
#IfWinActive
#IfWinActive PLANET
    +^l::
        {
            KeyWait, control
            KeyWait, shift
            ; MsgBox, , , Text,
            GuiControl, repl: , listbox1 , %vl_repl_liste%
            Gui repl: Show, w620 h420, Window
            return
        }

    ^l::
        {
            vl := P6_hent_vl()
            sleep 200
            p6_liste_vl(vl)
        }
#IfWinActive
;; Telenor

;; Trio
; ***
; Sæt kopieret tlf i Trio
Trio_opkald(ByRef telefon)
{

    ifWinNotExist, ahk_class AccessBar
    {
        WinActivate, ahk_class Agent Main GUI
        WinWaitActive, ahk_class Agent Main GUI
        sleep 100
        SendInput, {alt}
        sleep 100
        SendInput, v{Down 5}{enter}
    }
    ControlClick, x360 y17, ahk_class AccessBar
    sleep 800
    WinActivate, ahk_class Addressbook
    WinwaitActive, ahk_class Addressbook
    ControlClick, Edit2, ahk_class Addressbook
    SendInput, ^a{del}
    ; sleep 200
    ; SendInput, {NumpadSub}
    sleep 200
    SendInput, %telefon%
    sleep 500
    SendInput, +{enter} ; undgår kobling ved igangværende opkald
    Return
}

; ***
; Læg på i Trio
Trio_afslutopkald()
{
    WinActivate, ahk_class AccessBar
    winwaitactive, ahk_class AccessBar
    sleep 40
    SendInput, {NumpadSub}

    return
}

; **
; Trio hop til efterbehandling
trio_efterbehandling()
{
    WinActivate, ahk_class Agent Main GUI
    winwaitactive, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 8
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; **
; Trio hop til midt uden overløb
trio_udenov()
{
    WinActivate, ahk_class Agent Main GUI
    winwaitactive, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 3
    sleep 100
    SendInput, {F4}
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; **
; Trio hop til alarm
trio_alarm()
{
    WinActivate, ahk_class Agent Main GUI
    winwaitactive, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 7
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; **
; Trio hop til pause
trio_pause()
{
    WinActivate, ahk_class AccessBar
    winwaitactive, ahk_class AccessBar
    sleep 100
    SendInput, {F3}
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; **
; Trio hop til klar
trio_klar()
{
    WinActivate, ahk_class AccessBar
    winwaitactive, ahk_class AccessBar
    Sleep 100
    SendInput, {F4}
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; **
; Trio hop til frokost
trio_frokost()
{
    WinActivate, ahk_class Agent Main GUI
    winwaitactive, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 9
    WinActivate, PLANET
    winwaitactive, PLANET
    Return
}

; Trio skift mellem pause og klar

trio_pauseklar()
{
    WinActivate, ahk_class AccessBar
    winwaitactive, ahk_class AccessBar
    Sleep 200
    SendInput, {F3}
    sleep 400
    SendInput, {F4}
    WinActivate, PLANET
    winwaitactive, PLANET

    Return
}

;  ***
;Træk tlf fra Trio indkomne kald
Trio_hent_tlf()
{
    clipboard := ""
    sleep 200
    Sendinput !+k
    ClipWait, 1
    if (clipboard = "")
    {
        SendInput, !+k
        ClipWait, 1
    }
    Telefon := Clipboard
    trio_tlf_knap(Telefon)
    rentelefon := Substr(Telefon, 4, 8)
    return rentelefon
}

trio_tlf_knap(ByRef tlf := "")
{
    ; global tlf
    ; SendInput, +!k
    ; tlf := "test
    ; tlf := "+4512345678"
    if (SubStr(tlf, 1, 1) = "+")
        tlf_knap := SubStr(tlf, 4, 4) . " " . (SubStr(tlf, 8, 4))
    else
        tlf_knap := SubStr(tlf, 1, 4) . " " . SubStr(tlf, 5, 4)
    sleep 100
    GuiControl, tlf:text, Button1, Tlf: %tlf_knap%
    return
}

;; Flexfinder

; *
; Kørselsaftale til flexfinder
; 244,215
Flexfinder_opslag()
{
    KeyWait, Shift,
    KeyWait, Ctrl
    If (WinExist("FlexDanmark FlexFinder"))
    {
        k_aftale := P6_hent_k()
        k_aftale := SubStr("000" . k_aftale, -3) ; indsætter nuller og tager sidste fire cifre i strengen.
        ; MsgBox, , er 4 , % k_aftale
        sleep 200
        WinActivate, FlexDanmark FlexFinder
        winwaitactive, FlexDanmark FlexFinder
        sleep 40
        SendInput, {Home}
        sleep 400
        SendInput, {PgUp}
        sleep 200
        WinGetPos, X, Y, , , FlexDanmark FlexFinder, , ,
        if(x = "1920" or x = "-1920")
        {
            PixelSearch, Px, Py, 1097, 74, 1202, 123, 0x5B6CF2, 0, Fast ; Virker ikke i fuld skærm. ControlClick i stedet?
            sleep 200
            click %Px% %Py%
            sleep 200
            ControlClick, x322 y100, FlexDanmark FlexFinder
            sleep 40
            SendInput, +{tab}{up}{tab}
            sleep 200
            SendInput, %k_aftale%
            KeyWait, Enter, D, T7
            sleep 200
            WinActivate, PLANET
        }
        Else
        {
            PixelSearch, Px, Py, 90, 190, 1250, 250, 0x5E6FF2, 0, Fast
            sleep 200
            click %Px% %Py%
            sleep 200
            ControlClick, x244 y215, FlexDanmark FlexFinder
            sleep 40
            SendInput, +{tab}{up}{tab}
            sleep 200
            SendInput, %k_aftale%
            KeyWait, Enter, D, T7
            sleep 200
            WinActivate, PLANET
        }
        ; SendInput, {CtrlUp}{ShiftUp} ; for at undgå at de hænger fast
    }
    Else
        MsgBox, , FlexFinder, Flexfinder ikke åben (skal være den forreste fane)
    Return
}

; Klik VL i FlexFinder, slår op i p6
; skal tilpasse Edge også
Flexfinder_til_p6()
{

    vl := {}
    sleep 40
    SendInput, {Home}
    sleep 400
    SendInput, {PgUp}
    BlockInput, Mouse
    WinGetPos, X, Y, , , FlexDanmark FlexFinder, , ,
    if(x = "0")
        PixelGetColor, pixel, 281, 155
    if (pixel = 0xFCFBFB)
    {
        MsgBox, , FlexFinder, Fanenerne "Grupper" og "Tid" i FlexFinder skal være lukket.
    return 0
}
    if (x = 0)
    {
        ; PixelSearch, Px, Py, 90, 190, 1062, 621, 0x7E7974, 0, Fast ; Virker ikke i fuld skærm. ControlClick i stedet?
        ; click %Px%, %Py%
        ; click %Px%, %Py%
        ControlClick, x281 y155, FlexDanmark FlexFinder
        ControlClick, x281 y155, FlexDanmark FlexFinder
        BlockInput, MouseMoveOff
        SendInput, ^c
        sleep 400
        ff_opslag := clipboard
        vl.1 := SubStr(ff_opslag, 1, 4)
        vl.2 := SubStr(ff_opslag, 6, 4)
        vl.2 := StrReplace(vl.2, 0, , , Limit := -1)
        return vl
    }
    else
        PixelGetColor, pixel, 236, 262
    if (pixel = 0xFBFBFB)
    {
        MsgBox, , FlexFinder, Fanenerne "Grupper" og "Tid" i FlexFinder skal være lukket.
        return 0
    }
    Else
    {
        ControlClick, x236 y262, FlexDanmark FlexFinder
        ControlClick, x236 y262, FlexDanmark FlexFinder
        BlockInput, MouseMoveOff
        SendInput, ^c
        sleep 400
        ClipWait, 2, 0
        ff_opslag := clipboard
        vl.1 := SubStr(ff_opslag, 1, 4)
        vl.2 := SubStr(ff_opslag, 6, 4)
        vl.2 := StrReplace(vl.2, 0, , , Limit := -1)
        return vl
    }

}

;; Outlook
; ***
; Åbn ny mail i outlook. Kræver nymail.lnk i samme mappe som script. Kolonne 37
Outlook_nymail()
{
    genvej_mod := sys_genvej_til_ahk_tast(37)
    sys_genvej_keywait(genvej_mod)
    Run, %A_linefile%\..\lib\nymail.lnk, , ,
    WinWaitActive, Ikke-navngivet - Meddelelse (HTML) , , , ,
    Return
}

;; Excel

Excel_vl_til_udklip()
{

    tast := GetKeyState("ctrl", "P")
    if tast = 0
    {
        SendInput, {AltUp}
        SendInput, {LButton}
    }
    clipboard :=
    sleep 800
    SendInput, ^c
    ClipWait, 6
    sleep 200
    SendInput, {Esc} ;
    vl := clipboard
    vl := StrReplace(vl, "`n", "")
    vl := StrReplace(vl, "`r", "")
    if (StrLen(vl) = 5) ; fem c<ifre plus new-line
    {
        return vl
    }
    else
        return 0

}

Excel_udklip_til_p6(byref vl:="")
{
    if vl = 0
    {
        MsgBox, , Klik på vognløb, Du skal klikke på vognløbet,
        return
    }
    Else
    {
        WinActivate, PLANET
        P6_udfyld_vl(vl)
        input, tast, L1 V T4, {Up}{Down}{tab}{LButton}
        if (tast = chr(27))
        {
            sleep 100 ; forhindrer hop tilbage til P6, hvis infobox
            WinActivate, Garantivognsoversigt FG8.xlsm
            return
        }
        if ErrorLevel
        {
            return
        }
    }
    return
}

;; System

; asd

mod_up()
{
    SendInput, {AltUp}{ShiftUp}{CtrlUp}{LWinUp}{RWinUp}
    ; Loop, 0xFF
    ;     IF GetKeyState(Key:=Format("VK{:X}",A_Index))
    ;         SendInput, {%Key% up}
    ; Return
}

; *
; færdigskrives
sys_genveje_opslag()
{
    global bruger_genvej
    global genvej_ren := []
    global genvej_navn := databaseget("%A_linefile%\..\db\bruger_ops.tsv", 1)
    for index, genvej in bruger_genvej
    {
        genvej_ren[index] := StrReplace(genvej, "+", "Shift + ")
        ; genvej_ren[index] := StrReplace(genvej, "!", "Alt + ")
        ; genvej_ren[index] := StrReplace(genvej, "^", "Control + ")
        ; MsgBox, , , % genvej
    }
    for index, genvej in genvej_ren
    {
        ;    genvej_ren[index] := StrReplace(genvej, "+", "Shift + ")
        ; genvej_ren[index] := StrReplace(genvej, "!", "Alt + ")
        genvej_ren[index] := StrReplace(genvej, "^", "Ctrl + ")
        ; MsgBox, , , % genvej
    }
    for index, genvej in genvej_ren
    {
        ; genvej_ren[index] := StrReplace(genvej, "+", "Shift + ")
        genvej_ren[index] := StrReplace(genvej, "!", "Alt + ")
        ; genvej_ren[index] := StrReplace(genvej, "^", "Control + ")
        ; MsgBox, , , % genvej
    }
    for index, genvej in genvej_ren
    {
        ; genvej_ren[index] := StrReplace(genvej, "+", "Shift + ")
        genvej_ren[index] := StrReplace(genvej, "#", "Windows + ")
        ; genvej_ren[index] := StrReplace(genvej, "^", "Control + ")
        ; MsgBox, , , % genvej
    }

    ; MsgBox,% genvej_navn.4 " - " genvej_ren.4 "`n"  genvej_navn.5 " - " genvej_ren.5

    ; MsgBox, , Genvej, % StrReplace(bruger_genvej.30, "+" , "Shift + ")
    return
}
; omdan genvejstaster til AHK-key. Tager genvejens kolonnenummer
sys_genvej_til_ahk_tast(byref kolonne := "")
{
    global bruger_genvej

    genvej_mod := StrSplit(bruger_genvej[kolonne])
    
    for i, e in genvej_mod
        {
        genvej_mod[i] := StrReplace(genvej_mod[i], "!", "alt")
        genvej_mod[i] := StrReplace(genvej_mod[i], "^", "control")
        genvej_mod[i] := StrReplace(genvej_mod[i], "+", "shift")
        genvej_mod[i] := StrReplace(genvej_mod[i], "#", "lwin")
        }
        return genvej_mod   
    ; genvej_mod := StrReplace(genvej_mod, "+", "shift")
    ; genvej_mod := StrReplace(genvej_mod, "^", "control")
    ; win skal være enten left/right - hvordan?
    ; MsgBox, , , % genvej_mod
    ; KeyWait, %genvej_mod%
}
sys_genvej_keywait(byref genvej_mod := "")
{
    genvej_mod1 := genvej_mod.1
    genvej_mod2 := genvej_mod.2
    KeyWait, %genvej_mod1%,
    if (genvej_mod2 = "shift" or genvej_mod2 = "alt" or genvej_mod2 = "control" genvej_mod2 = "lwin")
        keywait, %genvej_mod2%
}
sys_initialer()
{
    FormatTime, tid, ,HHmm ;definerer format på tid/dato
    initialer = /mt%A_userName%%tid%
    return initialer
} 
; l_sys_inputbox_til_fra:

; brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
; p6_udregn_minut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,44)
; p6_vl_slut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,42)
; if (p6_udregn_minut_ops = 0)
;     min_default := 2
; if (p6_udregn_minut_ops = 1)
;     min_default := 1
; if (p6_vl_slut_ops = 0)
;     vl_default := 2
; if (p6_vl_slut_ops = 1)
;     vl_default := 1
; Gui, sys:New
; Gui, sys:default
; Gui Font, s9, Segoe UI
; Gui Add, Text, x9 y32 w115 h23 +0x200, P6 - VL Sluttid
; Gui Add, Text, x8 y64 w123 h23 +0x200, P6 - Minutudregner
; Gui Add, DropDownList, vp6_vl_slut x144 y32 w120 Choose%vl_default%, Med Inputbox|Uden Inputbox|
; Gui Add, DropDownList, vp6_minut x144 y64 w120 Choose%min_default%, Med Inputbox|Uden Inputbox|
; Gui Add, Button, gsysok, &OK

; Gui Show, w307 h332, Window
; Return

; sysok:
; GuiControlGet, p6_vl_slut
; GuiControlGet, p6_minut
; if (p6_vl_slut ="Med Inputbox")
; {
;     p6_vl_ops = 1
;     gui, cancel
;     databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
; }
; if (p6_vl_slut ="Uden Inputbox")
; {
;     p6_vl_ops = 0
;     gui, cancel
;     databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
; }
; if (p6_minut ="Med Inputbox")
; {
;     p6_minut_ops = 1
;     gui, cancel
;     databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
; }
; if (p6_minut ="Uden Inputbox")
; {
;     p6_minut_ops = 0
;     gui, cancel
;     databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
; }
; sleep 200
; WinActivate, PLANET
; return

sysGuiEscape:
sysGuiClose:
    gui, cancel
return

;; Misc
; *
; SygehusGUI
; omskrives
l_p6_sygehus_ring_op:
    genvej_beskrivelse(18)
    vis_sygehus_1()
    mod_up()
return

l_p6_central_ring_op:
    genvej_beskrivelse(19)
    gui, Taxa:Default
    Gui,Add,Button,vtaxa1,&Århus Taxa
    Gui,Add,Button,vtaxa2,Århus Taxa Sk&ole
    Gui,Add,Button,vtaxa3,&Dantaxi
    Gui,Add,Button,vtaxa4,Taxa &Midt
    Gui,Add,Button,vtaxa5,D&K Taxi
    Gui,Show, AutoSize Center , Ring op til central
    taxaknap1:=Func("opkaldtaxa").Bind("89484892")
    taxaknap2:=Func("opkaldtaxa").Bind("89484837")
    taxaknap3:=Func("opkaldtaxa").Bind("96341121")
    taxaknap4:=Func("opkaldtaxa").Bind("97120777")
    taxaknap5:=Func("opkaldtaxa").Bind("87113030")
    GuiControl,+g,taxa1,%taxaknap1%
    GuiControl,+g,taxa2,%taxaknap2%
    GuiControl,+g,taxa3,%taxaknap3%
    GuiControl,+g,taxa4,%taxaknap4%
    GuiControl,+g,taxa5,%taxaknap5%
return
Opkaldtaxa(p*){
    Gui, taxa: Destroy
    telefon := % p.1
    sleep 100
    Trio_opkald(telefon)
    WinActivate, PLANET, , ,
}
TaxaGuiClose:
    gui, Destroy
return

TaxaGuiEscape:
    Gui, Destroy
return

vis_sygehus_1()
{
    Gui, sygehus:Show, w144 h240, Ring til Sygehus
    Return
}
vis_sygehus_2(navn)
{
    if (navn = "misc")
        Gui, sygehus%navn%:Show, w144 h300, AUH
    else
        Gui, sygehus%navn%:Show, w144 h240, AUH
    Return
}

;; Testknap

; ^+e::
;     {
;         ControlClick, x360 y17, ahk_class AccessBar
;         return
;     }
;; HOTKEYS

;; Global
l_quitAHK:
ExitApp
Return

l_p6_aktiver:
    p6_aktiver()
    afslut_genvej()
return
;; PLANET
l_p6_hastighed:
    P6_hastighed()
    afslut_genvej()
return
; skriv/fjern initialer. Kolonne 5
l_p6_initialer: ;; Initialer til/fra
    genvej_mod := sys_genvej_til_ahk_tast(5)
    sys_genvej_keywait(genvej_mod)
    P6_initialer()
    afslut_genvej()
Return
; skriv initialer, fortsæt notat. Kolonne 6
l_p6_initialer_skriv: ; skriv initialer og forsæt notering.
    genvej_mod := sys_genvej_til_ahk_tast(6)
    sys_genvej_keywait(genvej_mod)
    P6_initialer_skriv()
    afslut_genvej()
return

l_p6_vis_k_aftale: ;Vis kørselsaftale for aktivt vognløb
    P6_vis_k()
    afslut_genvej()
Return

l_p6_ret_vl_tlf: ; +F3 - ret vl-tlf til triopkald
    faste_dage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
    uge_dage := ["faste mandage", "faste tirsdage", "faste onsdage", "faste torsdage", "faste fredage", "faste lørdage", "faste søndage"]

    genvej_beskrivelse(8)
    genvej_mod := sys_genvej_til_ahk_tast(8)
    sys_genvej_keywait(genvej_mod)

    ; SendInput, {ShiftUp}{AltUp}{CtrlUp}
    klip := clipboard
    sleep 100
    telefon := Trio_hent_tlf()

    WinActivate, PLANET
    vl := P6_hent_vl()
    ; if (telefon = "")
    ; {
    ;     telefon := "Ikke registreret"
    ; }
    ; else
    ; {
        clipboard := klip
        if (telefon = "")
            telefon := "ikke registreret"
        InputBox, telefon, VL, Skal der bruges et andet telefonnummer end %telefon%?,, 160, 180, X, Y, , Timeout, %telefon%
        if (ErrorLevel = 1 or ErrorLevel = 2)
        {
            afslut_genvej()
            return
        }
        sleep 100
        MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl%?,
        IfMsgBox, no
        {
            afslut_genvej()
            return
        }
        IfMsgBox, Yes
            P6_ret_tlf_vl(telefon)
        sleep s * 100
        Input, næste, L1 V T4
        if (næste = "n")
        {
            Loop
            {
                Clipboard :=
                sleep 20
                SendInput, !l{tab}
                sleep 200
                SendInput, !{right}
                sleep 400
                SendInput, ^c
                clipwait 3
                if (InStr(clipboard, "eksistere"))
                    continue
                if (ErrorLevel = 1)
                    SendInput, ^c
                clipwait 3
                tid_ind := Clipboard
                if (InStr(tid_ind, "senare"))
                {
                    SendInput, {enter}
                    for index, element in faste_dage
                    {
                        SendInput, !l
                        sleep 200
                        sendinput {tab}
                        SendInput, %element% {enter}
                        sleep 200
                        clipboard :=
                        SendInput, {tab}
                        SendInput, ^c
                        clipwait 3
                        if (clipboard = element)
                        {
                            ; MsgBox, , , match
                            tid_ind := uge_dage[index]
                            ramt_dag := index
                            ; MsgBox, , overført, % tid_ind
                            Break
                        }
                        else
                        {
                            ; MsgBox, , ,% clipboard "." element,
                            SendInput, {enter}
                            continue
                        }

                    }

                }
                dato := SubStr(tid_ind, 4, 2) SubStr(tid_ind, 1, 2)
                OutputDebug, % dato
                dato:= A_YYYY dato
                FormatTime, dato, %dato%, dddd dd/MM
                if (StrLen(clipboard) = 2)
                {
                    StringLower, clipboard, clipboard, ; dage med ø skal være lowercase. Hvorfor?
                    for index, element in faste_dage
                    {
                        if (clipboard = element)
                        {
                            ramt_dag := index
                            ; MsgBox, , , % ramt_dag,
                        }
                    }
                    dato := uge_dage[ramt_dag]
                    ; MsgBox, , næste dag , % uge_dag[ramt_dag]
                }
                ; MsgBox, , , % dato,
                sleep 100
                MsgBox, 4, Sikker?, Vil du sætte %telefon% på VL %vl% på %dato%?,
                IfMsgBox, Yes
                {
                    if (ramt_dag = 7)
                    {
                        P6_ret_tlf_vl_efterfølgende(telefon)
                        afslut_genvej()
                        return
                    }
                    sleep 200
                    P6_ret_tlf_vl_efterfølgende(telefon)
                    sleep 200
                    continue
                }
                IfMsgBox, no
                {
                    afslut_genvej()
                    return
                }
            }

        }
    ; }
    afslut_genvej()
return

#IfWinActive ; for at resette indent
; ***
l_p6_søg_vl: ; Søg VL ud fra indgående kald i Trio
    global s
    tlf := Trio_hent_tlf()
    WinActivate, PLANET, , ,
    sleep s * 40
    vl := P6_hent_vl_fra_tlf(tlf)
    if (vl = 0)
    {
        MsgBox, , Tlf ikke registreret , Telefonnummeret er ikke registreret i Ethics., 1
        WinActivate, PLANET, , ,
        SendInput, !tp!l
        afslut_genvej()
        return
    }
    else
        sleep s * 40
    P6_udfyld_k_og_s(vl)
    afslut_genvej()
Return

; ***r
l_trio_til_p6: ;træk tlf til rejsesøg
    global s

    If (IfWinNotExist, PLANET, , , )
        MsgBox, , PLANET, P6 er ikke åben.,
    Else
    {
        telefon := Trio_hent_tlf()
        if (telefon = "")
        {
            MsgBox, , Intet indgående telefonnummer, Der er intet indgående telefonnummer, 1
            afslut_genvej()
            return
        }
        if (telefon = "78410222")
        {

            P6_rejsesogvindue()
            sleep s * 40
            SendInput, ^t
            afslut_genvej()
            return
        }
        Else
        {
            WinActivate, PLANET
            P6_rejsesog_tlf(telefon)
            afslut_genvej()
            return
        }
    }
return

; gå i vl
l_p6_vaelg_vl:
    p6_vaelg_vl()
    afslut_genvej()
return
;træk tlf fra aktiv planbillede, ring op i Trio. Col 11
l_p6_vl_ring_op:
    genvej_beskrivelse(11)
    genvej_mod := sys_genvej_til_ahk_tast(11)
    sys_genvej_keywait(genvej_mod)
    sleep s * 100
    vl_tlf := P6_hent_vl_tlf()
    if (vl_tlf = 0)
    {
        afslut_genvej()
        return
    }
    if (vl_tlf = "")
    {
        MsgBox, 4, Prøv igen?, Tlf-nr ikke opfanget. Prøv igen?
        IfMsgBox, yes
            Goto, l_p6_vl_ring_op
        IfMsgBox, no
            afslut_genvej()
        return
    }
    sleep 200
    Trio_opkald(vl_tlf)
    ; Clipboard = %gemtklip%
    ; gemtklip :=
    sleep 400
    WinActivate, PLANET
    P6_Planvindue()
    afslut_genvej()
return

; ***

; ^+F5 col 12
l_p6_vm_ring_op: ; træk vm-tlf fra aktivt planbillede, ring op i Trio
    genvej_beskrivelse(12)
    genvej_mod := sys_genvej_til_ahk_tast(12)
    sys_genvej_keywait(genvej_mod)

    P6_planvindue()
    sleep s * 100
    vm_tlf := P6_hent_vm_tlf()
    sleep 500
    Trio_opkald(vm_tlf)
    sleep 800
    WinActivate, PLANET

    afslut_genvej()
Return

; P6 - ring op til kunde markeret i Vl (kræver tlf opsat på kundetilladelse)
l_p6_ring_til_kunde:
    p6_hent_kunde_tlf(telefon)
    sleep s * 200
    if (SubStr(telefon, 1, 3) = "888")
    {
        MsgBox, , Telefon ikke tilknyttet, Kunden har ikke telefon tilknyttet.
        afslut_genvej()
        return
    }
    Else
    {
        Trio_opkald(telefon)
        afslut_genvej()
        return
    }
return

; #F5, col 13
l_p6_vl_luk:
    genvej_beskrivelse(13)
    genvej_mod := sys_genvej_til_ahk_tast(13)
    sys_genvej_keywait(genvej_mod)
    gemtklip := ClipboardAll

    tid := P6_input_sluttid()
    if !tid
    {
        afslut_genvej()
        return
    }
    p6_vl_luk(tid)
    sleep 100
    P6_planvindue()
    sleep 200
    SendInput, {F5}
    afslut_genvej()

    clipboard := gemtklip
    gemtklip :=
return

l_p6_udregn_minut:
    ; mod_up()
    tid := P6_udregn_minut()
    tid_tekst := tid.1
    if (tid = "fejl")
    {
        afslut_genvej()
        return
    }
    gui, plustid:New,
    gui, plustid:Default
    Gui Font, s9, Segoe UI
    Gui Add, Button, gok x24 y88 w80 h23 +Default, &OK
    Gui Add, Button, gudklip x144 y88 w80 h23, Til &Udklip
    Gui Add, Text, x72 y24 w120 h23 +0x200 +Center, %tid_tekst%

    sleep 100
    Gui Show, w260 h125, Resultat
Return

ok:
    {
        gui, cancel
        afslut_genvej()
        return
    }
udklip:
    {
        Clipboard := tid.2
        gui, cancel
        afslut_genvej()
        return
    }
plustidGuiEscape:
plustidGuiClose:
    gui, cancel
    afslut_genvej()
return

l_p6_alarmer:
    genvej_beskrivelse(14)
    P6_alarmer()
    afslut_genvej()
return

l_p6_udraabsalarmer:
    genvej_beskrivelse(15)

    P6_udraabsalarmer()
    afslut_genvej()
return
l_p6_tekst_til_chf: ; Send tekst til aktive vognløb
    genvej_mod := sys_genvej_til_ahk_tast(20)
    sys_genvej_keywait(genvej_mod)
    FormatTime, Time, ,HHmm
    initialer = /mt%A_userName%%time%
    initialer_udentid =/mt%A_userName%
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    bruger := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 2)
    ; ctrl_s := chr(19)

    genvej_beskrivelse(20)

    ; KeyWait Alt
    ; keywait Ctrl
    Input valgt, L1 T5 C, {esc},
    if (valgt = "t")
    {
        P6_tekstTilChf() ; tager tekst ("eksempel") som parameter (accepterer variabel)
        afslut_genvej()
        return
    }
    if (valgt = "f")
    {
        gui, f_chf:New
        gui, f_chf:Default
        Gui Font, s9, Segoe UI
        Gui Add, Edit, vf_stop x15 y29 w120 h21,
        Gui Add, Text, x16 y7 w120 h23 +0x200, Forgæves stop
        Gui Add, Edit, vs_stop x214 y32 w120 h21
        Gui Add, Text, x216 y7 w120 h23 +0x200, Sendt stop
        Gui Add, Edit, vk_navn x14 y106 w120 h21
        Gui Add, Text, x16 y86 w120 h23 +0x200, Navn på kunde forg.
        Gui Add, Text, x215 y84 w120 h23 +0x200, Navn på kunde sendt
        Gui Add, Edit, vk_navn2 x216 y103 w120 h21
        Gui Add, Button, gf_chfok x81 y172 w80 h23 +Default, &OK
        Gui Add, Button, gf_annuller x216 y171 w80 h23, &Annuller

        Gui Show,x812 y22 w381 h220, Send tekst om forgæves til chauffør
        Return

        f_annuller:
        f_chfGuiEscape:
        f_chfGuiClose:
            {
                gui, Cancel
                afslut_genvej()
                return
            }
        f_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            ; MsgBox, , , % tekst,
            gui, cancel
            P6_tekstTilChf("Jeg kan ikke ringe dig op. Jeg har meldt st. " f_stop "`, " . k_navn "`, forgæves og sendt st. " s_stop "`, " k_navn2 ", i stedet - Mvh. Midttrafik")
            sleep 500
            MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
            IfMsgBox, Yes
            {
                SendInput, ^s
                ; KeyWait, Ctrl
                sleep 1000
                SendInput, {enter}
                P6_notat("Ingen kontakt til chf. St. " f_stop " forgæves`, " s_stop " og tekst sendt til chf." initialer)
                afslut_genvej()
                return
            }
            IfMsgBox, No
            {
                sleep 200
                MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
                gui, cancel
            }
            afslut_genvej()
        return
    }
    if (valgt = "k")
    {
        gui, k_chf:New
        gui, k_chf:Default
        Gui Font, s9, Segoe UI
        Gui Add, Edit, vf_stop x15 y29 w120 h21,
        Gui Add, Text, x16 y7 w120 h23 +0x200, Kvitteret stop
        Gui Add, Edit, vs_stop x214 y32 w120 h21
        Gui Add, Text, x215 y7 w120 h23 +0x200, Sendt stop
        Gui Add, Edit, vk_navn x14 y106 w120 h21
        Gui Add, Text, x16 y86 w120 h23 +0x200, Navn på kunde kvit.
        Gui Add, Text, x215 y86 w120 h23 +0x200, Navn på kunde sendt
        Gui Add, Text, x120 y137 w120 h23 +0x200, Evt. kvitteret tid.
        Gui Add, Edit, vk_navn2 x216 y103 w120 h21
        Gui Add, Edit, vk_tid x120 y157 w120 h21, Oprindelig kvittering
        Gui Add, Button, gk_chfok x81 y200 w80 h23 +Default, &OK
        Gui Add, Button, gk_annuller x216 y200 w80 h23, &Annuller

        Gui Show, x812 y22 w381 h280, Send tekst om kvittering til chauffør
        Return

        k_annuller:
        k_chfGuiEscape:
        k_chfGuiClose:
            {
                gui, Cancel
                afslut_genvej()
                return
            }
        k_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            GuiControlGet, k_tid, , ,
            gui, cancel
            P6_tekstTilChf("Husk at bede om ny tur ved ankomst. Jeg har bekræftet ankomst ved st. " f_stop "`, " . k_navn "`, og sendt st. " s_stop "`, " k_navn2 " - Mvh. Midttrafik")
            sleep 500
            MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
            IfMsgBox, Yes
            {
                SendInput, ^s
                sleep 1000
                SendInput, {enter}
                if (k_tid != "Oprindelig kvittering")
                {
                    P6_notat("St. " f_stop " ikke kvitteret ved ankomst`, st. " s_stop " og tekst sendt til chf. Oprindeligt kvitt. tid " k_tid initialer " ")
                    return
                }
                else
                    P6_notat("St. " f_stop " ikke kvitteret ved ankomst`, st. " s_stop " og tekst sendt til chf. " initialer " ")
                afslut_genvej()
                return
            }
            IfMsgBox, No
            {
                sleep 200
                MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
                gui, cancel
            }
            afslut_genvej()
        return
    }
    if (valgt == "p")
    {

        P6_tekstTilChf("Er der blevet glemt at kvittere for privatrejsen? Mvh. Midttrafik")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Priv. ikke kvitteret, tekst sendt til chf" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        afslut_genvej()
        return
    }
    if (valgt == "P")
    {

        P6_tekstTilChf("Jeg kan ikke ringe dig op, din privatrejse er ikke kvitteret. Vognløbet er låst, ring til driften, hvis du er ude at køre.")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Priv. ikke kvitteret, ingen kontakt til chf. Tekst sendt om VL-lås" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        afslut_genvej()
        return
    }
    if (valgt == "w")
    {

        P6_tekstTilChf("Der er ikke bedt om vognløb start. Huske at bede om første køreordre ved opstart, uanset om der ligger ture eller ej. Mvh. Midttrafik")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_planvindue()
            vl := P6_hent_vl()
            p6_liste_vl(vl)
            P6_notat("WakeUp sendt" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            afslut_genvej()
            return
        }
    }
    if (valgt == "W")
    {

        P6_tekstTilChf("Jeg kan ikke ringe dig op, der er ikke trykket for første køreordre. Ring til driften, hvis du er ude at køre, ellers bliver vognløbet lukket.")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør? Husk at låse VL,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Ingen kontakt til chf, tekst sendt, VL låst" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            afslut_genvej()
            return
        }
    }
    if (valgt == "r")
    {
        tlf := P6_hent_vl_tlf()
        P6_tekstTilChf("Jeg kan ikke ringe dig op på telefonnummer " tlf ". Ring til driften, 70112210. Mvh Midttrafik.")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør? 
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Ingen kontakt til chf, tekst sendt (ring til driften)" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            afslut_genvej()
            return
        }
    }
    if (valgt = "a")
    {

        P6_tekstTilChf("Jeg kan ikke ringe dig op. Tryk for opkald igen, hvis du stadig gerne vil ringes op. Mvh. Midttrafik")
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Tal forgæves, tekst sendt" initialer " ")
            gui, cancel
            afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        afslut_genvej()
        return
    }
    afslut_genvej()
    return
#IfWinActive ; udelukkende for at resette indentering i auto-formatering

;; Trio
l_trio_klar: ;Trio klar
    trio_klar()
Return

l_trio_pause: ;Trio pause
    trio_pause()
Return

l_trio_udenov: ;Trio Midt uden overløb
    trio_udenov()
    trio_klar()
Return

l_trio_efterbehandling: ;Trio efterbehandling
    trio_efterbehandling()
    trio_pauseklar()
Return

l_trio_alarm: ;Trio alarm bruger.9
    trio_alarm()
Return

l_trio_frokost: ;Trio frokostr. bruger.10
    trio_frokost()
Return

l_triokald_til_udklip: ; trækker indkommende kald til udklip, ringer ikke op.
    clipboard := Trio_hent_tlf()
    afslut_genvej()
Return

; Telenor accepter indgående kald, søg planet
l_trio_P6_opslag: ; brug label ist. for hotkey, defineret ovenfor. Bruger.4
    genvej_beskrivelse(3)
    genvej_mod := sys_genvej_til_ahk_tast(4)
    sys_genvej_keywait(genvej_mod)
    SendInput, % bruger_genvej[3] ; opr telenor-genvej
    sleep 40
    SendInput, % bruger_genvej[3] ; Misser den af og til?
    sleep 40
    telefon := Trio_hent_tlf()
    sleep 40
    P6_aktiver()
    if (telefon = "")
    {
        MsgBox, , , Intet indgående telefonnummer el. hemmeligt nummer, 1
        P6_aktiver()
        sleep 100
        p6_vaelg_vl()
        afslut_genvej()
        return
    }
    vl := P6_hent_vl_fra_tlf(telefon)
    if vl
    {
        sleep 200
        P6_udfyld_k_og_s(vl)
        afslut_genvej()
        Return
    }
    if (telefon = "78410222" OR telefon ="78410224") ; mangler yderligere?
    {
        ; MsgBox, ,CPR, CPR, 1
        sleep 200
        P6_rejsesogvindue()
        SendInput, ^t
        afslut_genvej()
        return
    }
    Else
    {
        sleep 200
        P6_rejsesogvindue(telefon)
        afslut_genvej()
        return
    }
; Opkald på markeret tekst. Kolonne 28
l_trio_opkald_markeret: ; Kald det markerede nummer i trio, global. Bruger.12
    genvej_mod := sys_genvej_til_ahk_tast(28)
    sys_genvej_keywait(genvej_mod)
    clipboard := ""
    SendInput, ^c
    ClipWait, 2, 0
    telefon := clipboard
    sleep 300
    Trio_opkald(telefon)
    afslut_genvej()
Return

; Minus på numpad afslutter Trioopkald global (Skal der tilbage til P6?)
l_trio_afslut_opkald:
l_trio_afslut_opkaldB:
    mod_up()
    Trio_afslutopkald()
    sleep 200
    WinActivate, PLANET
Return

;; Flexfinder
l_flexf_fra_p6:
    mod_up()
    Flexfinder_opslag()
    afslut_genvej()
Return
; slå VL op i FF. Kolonne 36
l_flexf_til_p6:
    genvej_mod := sys_genvej_til_ahk_tast(36)
    sys_genvej_keywait(genvej_mod)
    sleep 200
    vl :=Flexfinder_til_p6()
    if !vl
    {
        afslut_genvej()
        return
    }
    Else
    {
        P6_aktiver()
        sleep s * 200
        P6_udfyld_k_og_s(vl)
        sleep 400 ; skal optimeres
        WinActivate, FlexDanmark FlexFinder, , ,
        afslut_genvej()
        Return
    }

;; Outlook
l_outlook_ny_mail: ; opretter ny mail. Bruger.16
    genvej_beskrivelse(37)
    Outlook_nymail()
    afslut_genvej()
Return

;; Excel til vl. 
l_excel_vl_til_P6_A:
l_excel_vl_til_P6_B:
    mod_up()
    vl := Excel_vl_til_udklip()
    sleep 400
    SendInput, {Esc}
    Excel_udklip_til_p6(vl)
return
;; HOTSTRINGS

::vllp::Låst, ingen kontakt til chf, privatrejse ikke udråbt
::bsgs::Glemt slettet retur
::rgef::Rejsegaranti, egenbetaling fjernet
::vlaok::Alarm st OK
::vlik::
    {
        ; hent st og tid - gui
        SendInput, St. %stop% ank. %tid%, ikke kvitteret
    }

::/in::
    FormatTime, tid, ,HHmm ;definerer format på tid/dato
    initialer = /mt%A_userName%%tid%
    Sendinput %initialer%
return

l_restartAHK: ; AHK-reload
    SendInput, {CtrlUp}
    Reload
    sleep 2000
Return

^+a::databaseview("%A_linefile%\..\db\bruger_ops.tsv")

;; GUI-hjælp

;hjælp GUI
l_gui_hjælp:
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1) ; brugerens række i databasen
    bruger_genvej := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1) ; array med alle brugerens data
    p6_udregn_minut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,44)
    p6_vl_slut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,42)
    p6_hastighed_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,41)
    genvej_ren := []
    genvej_navn := []

    global genvej_ren
    global genvej_navn
    global hk :=

    if (p6_udregn_minut_ops = 0)
        min_default := 2
    if (p6_udregn_minut_ops = 1)
        min_default := 1
    if (p6_vl_slut_ops = 0)
        vl_default := 2
    if (p6_vl_slut_ops = 1)
        vl_default := 1
    sys_genveje_opslag()

    Gui Font, s9, Segoe UI
    Gui Color, 0xC0C0C0
    Gui Add, StatusBar,, Status Bar
    Gui Add, Tab3, x0 y0 w748 h642 0x54010240, Oversigt|Genvejsoversigt P6|Genvejsoversigt Trio|Opsætning|Hjælp|Misc
    Gui Tab, Opsætning
    ; Fix placerinc
    Gui Add, edit, vp6_hastighed_ops x144 y32 w120, %p6_hastighed_ops%
    Gui Add, Text, x285 y32 h23 +0x200, Tilpas efter P6-langsomhed. 1 = hurtigst. Skal bruge punktum (eks. 1.2)
    Gui Add, Text, x10 y32 h23 +0x200, P6 - Hastighed
    Gui Add, Text, x12 y64 h23 +0x200, P6 - Luk VL
    Gui Add, Text, x285 y64 h23 +0x200 , Vælg om der skal bruges en popup der kan skrives i til funktionen Luk VL.
    Gui Add, Text, x285 y96 h23 +0x200, Vælg om der skal bruges en popup der kan skrives i til funktionen Minutudregner.
    Gui Add, Text, x10 y96 h23 +0x200, P6 - Minutudregner
    Gui Add, DropDownList, vp6_vl_slut x144 y64 w120 Choose%vl_default%, Med Inputbox|Uden Inputbox|
    Gui Add, DropDownList, vp6_minut x144 y96 w120 Choose%min_default%, Med Inputbox|Uden Inputbox|
    Gui Add, Button, gsysok, &OK
    ; Gui Tab, Genvejsoversigt
    ; Gui Font
    ; Gui Font, s12 Bold
    ; Gui Add, Text, x0 y0 w748 h642 +0x200, Generelt
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Tab, Genvejsoversigt Trio
    Gui Font
    Gui Font, s14 Bold q4, Segoe UI
    Gui Add, Text, x16 y32 w120 h23 +0x200, Trio
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Text, x8 y64 w227 h23 +0x200, % genvej_navn.3
    Gui Add, Text, x272 y64 w260 h23 +0x200, % genvej_ren.3
    Gui Add, Text, x8 y88 w227 h23 +0x200, % genvej_navn.22
    Gui Add, Text, x272 y88 w227 h23 +0x200, % genvej_ren.22
    Gui Add, Text, x8 y112 w227 h23 +0x200, % genvej_navn.23
    Gui Add, Text, x272 y112 w227 h23 +0x200, % genvej_ren.23
    Gui Add, Text, x8 y136 w227 h23 +0x200, % genvej_navn.24
    Gui Add, Text, x272 y136 w227 h23 +0x200, % genvej_ren.24
    Gui Add, Text, x8 y160 w227 h23 +0x200, % genvej_navn.25
    Gui Add, Text, x272 y160 w227 h23 +0x200, % genvej_ren.25
    Gui Add, Text, x8 y184 w227 h23 +0x200, % genvej_navn.26
    Gui Add, Text, x272 y184 w227 h23 +0x200, % genvej_ren.26
    Gui Add, Text, x8 y208 w227 h23 +0x200, % genvej_navn.27
    Gui Add, Text, x272 y208 w227 h23 +0x200, % genvej_ren.27
    Gui Add, Text, x8 y232 w227 h23 +0x200, % genvej_navn.28
    Gui Add, Text, x272 y232 w227 h23 +0x200, % genvej_ren.28
    Gui Add, Text, x8 y256 w227 h23 +0x200, % genvej_navn.29
    Gui Add, Text, x272 y256 w227 h23 +0x200, % genvej_ren.29
    Gui Add, Text, x8 y280 w227 h23 +0x200, % genvej_navn.30
    Gui Add, Text, x272 y280 w227 h23 +0x200, % genvej_ren.30
    Gui Add, Text, x8 y304 w227 h23 +0x200, % genvej_navn.31
    Gui Add, Text, x272 y304 w227 h23 +0x200, % genvej_ren.31
    Gui Add, Text, x8 y328 w227 h23 +0x200, % genvej_navn.32
    Gui Add, Text, x272 y328 w227 h23 +0x200, % genvej_ren.32
    ; Gui Add, Text, x8 y352 w227 h23 +0x200, % genvej_navn.33
    ; Gui Add, Text, x272 y352 w227 h23 +0x200, % genvej_ren.33
    Gui Add, Text, x8 y56 w198 h2 +0x10
    Gui Tab, Genvejsoversigt P6
    Gui Font
    Gui Font, s14 Bold q4, Segoe UI
    Gui Add, Text, x16 y32 w120 h23 +0x200, Planet
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Text, x8 y56 w198 h2 +0x10
    Gui Add, Text, x8 y64 w227 h23 +0x200, % genvej_navn.4
    Gui Add, Text, x248 y64 w260 h23 +0x200, % genvej_ren.4
    Gui Add, Text, x8 y88 w227 h23 +0x200, % genvej_navn.5
    Gui Add, Text, x248 y88 w260 h23 +0x200, % genvej_ren.5
    Gui Add, Text, x8 y112 w227 h23 +0x200, % genvej_navn.6
    Gui Add, Text, x248 y112 w260 h23 +0x200, % genvej_ren.6
    Gui Add, Text, x8 y136 w227 h23 +0x200, % genvej_navn.7
    Gui Add, Text, x248 y136 w260 h23 +0x200, % genvej_ren.7
    Gui Add, Text, x8 y160 w227 h23 +0x200, % genvej_navn.8
    Gui Add, Text, x248 y160 w260 h23 +0x200, % genvej_ren.8
    Gui Add, Text, x8 y184 w227 h23 +0x200, % genvej_navn.9
    Gui Add, Text, x248 y184 w260 h23 +0x200, % genvej_ren.9
    Gui Add, Text, x8 y208 w227 h23 +0x200, % genvej_navn.10
    Gui Add, Text, x248 y208 w260 h23 +0x200, % genvej_ren.10
    Gui Add, Text, x8 y232 w227 h23 +0x200, % genvej_navn.11
    Gui Add, Text, x248 y232 w260 h23 +0x200, % genvej_ren.11
    Gui Add, Text, x8 y256 w227 h23 +0x200, % genvej_navn.12
    Gui Add, Text, x248 y256 w260 h23 +0x200, % genvej_ren.12
    Gui Add, Text, x8 y280 w227 h23 +0x200, % genvej_navn.13
    Gui Add, Text, x248 y280 w260 h23 +0x200, % genvej_ren.13
    Gui Add, Text, x8 y304 w227 h23 +0x200, % genvej_navn.14
    Gui Add, Text, x248 y304 w260 h23 +0x200, % genvej_ren.14
    Gui Add, Text, x8 y328 w227 h23 +0x200, % genvej_navn.15
    Gui Add, Text, x248 y328 w260 h23 +0x200, % genvej_ren.15
    Gui Add, Text, x8 y352 w227 h23 +0x200, % genvej_navn.16
    Gui Add, Text, x248 y352 w260 h23 +0x200, % genvej_ren.16
    Gui Add, Text, x8 y376 w227 h23 +0x200, % genvej_navn.17
    Gui Add, Text, x248 y376 w260 h23 +0x200, % genvej_ren.17
    Gui Add, Text, x8 y400 w227 h23 +0x200, % genvej_navn.18
    Gui Add, Text, x248 y400 w260 h23 +0x200, % genvej_ren.18
    Gui Add, Text, x8 y424 w227 h23 +0x200, % genvej_navn.32
    Gui Add, Text, x248 y424 w260 h23 +0x200, % genvej_ren.32
    Gui Add, Text, x8 y448 w227 h23 +0x200, % genvej_navn.34
    Gui Add, Text, x248 y448 w260 h23 +0x200, % genvej_ren.34
    Gui Add, Text, x8 y472 w227 h23 +0x200, % genvej_navn.36
    Gui Add, Text, x248 y472 w260 h23 +0x200, % genvej_ren.36
    Gui Add, Text, x8 y496 w227 h23 +0x200, % genvej_navn.38
    Gui Add, Text, x248 y496 w260 h23 +0x200, % genvej_ren.38
    ; Gui Add, Text, x8 y520 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x248 y520 w260 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x8 y544 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x248 y544 w260 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x8 y568 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x248 y568 w260 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x8 y592 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x248 y592 w260 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x344 y64 w227 h23 +0x200,% genvej_navn.3
    ; Gui Add, Text, x344 y88 w227 h23 +0x200,% genvej_navn.3
    ; Gui Add, Text, x344 y112 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y136 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y160 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y184 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y208 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y232 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y256 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y280 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y304 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y328 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y352 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y376 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y400 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y424 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y472 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y496 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y520 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y544 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y568 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y592 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x344 y448 w227 h23 +0x200, % genvej_navn.3
    ; Gui Add, Text, x248 y448 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y472 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y496 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y520 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y544 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y568 w97 h23 +0x200, % genvej_ren.3
    ; Gui Add, Text, x248 y592 w97 h23 +0x200, % genvej_ren.3
    Gui Tab, Oversigt
    Gui Font
    Gui Font, s14 Bold q4, Segoe UI
    Gui Add, Text, x16 y32 w120 h23 +0x200, Generelt
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Text, x8 y56 w198 h2 +0x10
    Gui Add, Text, x8 y64 h23 +0x200, Skift mellem faner med pil højre/venstre. Genveje gælder som udgangspunkt kun når vinduet er p6 (ellers anført).
    Gui Add, Text, x8 y100 w227 h23 +0x200, % genvej_navn.33
    Gui Add, Text, x248 y100 w260 h23 +0x200, % genvej_ren.33
    Gui Add, Text, x8 y128 w227 h23 +0x200, % genvej_navn.46
    Gui Add, Text, x248 y128 w260 h23 +0x200, % genvej_ren.46
    Gui Add, Text, x8 y156 w227 h23 +0x200, % genvej_navn.47
    Gui Add, Text, x248 y156 w260 h23 +0x200, % genvej_ren.47
    Gui Tab, Misc
    Gui Tab

    Gui Show, w747 h670, AHK
Return

Return
gui, Submit, nohide

sysok:
    GuiControlGet, p6_vl_slut
    GuiControlGet, p6_minut
    GuiControlGet, p6_hastighed_ops
    gui, destroy
    if (p6_vl_slut ="Med Inputbox")
    {
        p6_vl_ops = 1
        databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
    }
    if (p6_vl_slut ="Uden Inputbox")
    {
        p6_vl_ops = 0
        databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
    }
    if (p6_minut ="Med Inputbox")
    {
        p6_minut_ops = 1
        databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
    }
    if (p6_minut ="Uden Inputbox")
    {
        p6_minut_ops = 0
        databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
    }
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 41, p6_hastighed_ops)

GuiEscape:
genvejGuiClose:
    gui, destroy
return
; Opret svigt på VL. Kolonne 38
l_outlook_svigt: ; tag skærmprint af P6-vindue og indsæt i ny mail til planet
    FormatTime, dato, , d/MM
    ; FormatTime, tid, , HH:mm
    trio_genvej := global genvej_navn := databaseget("%A_linefile%\..\db\bruger_ops.tsv", 3, 38)
    GuiControl, trio_genvej:text, Button1, %trio_genvej%
    ; svigt := []
    genvej_mod := sys_genvej_til_ahk_tast(38)
    sys_genvej_keywait(genvej_mod)
    gemtklip := ClipboardAll
    vl := P6_hent_vl()
    clipboard :=
    sleep 500
    SendInput, !{PrintScreen}
    sleep 500
    ; ClipWait, 3,
    klip := ClipboardAll
    ; clipwait 3, 1 ;; bedre løsning?
    gui, svigt:new
    gui, svigt:default
    Gui Font, w600
    Gui Add, Text, x16 y0 w120 h23 +0x200, Vognløbs&nummer
    Gui Font
    Gui Add, Edit, vVL x16 y24 w120 h21, %vl%
    Gui Font, s9, Segoe UI
    Gui Font, w600
    Gui Add, Text, x161 y0 w118 h25 +0x200, &Lukket? (Vælg én)
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, CheckBox, vlukket x160 y24 w39 h23, &Ja
    Gui Add, Edit, vtid x200 y24 w79 h21, Hjemzone kl.
    Gui Add, CheckBox, vhelt x160 y48 w120 h23, Ja, og VL &slettet:
    Gui Add, Edit, vtid_slet x170 y68 h21, Åbningstid garanti
    ; Gui Add, CheckBox, vhelt2 x160 y72 w120, GV garanti &slettet i variabel tid ; nødvendig?
    Gui Font, s9, Segoe UI
    Gui Font, w600
    Gui Add, Text, x304 y0 w120 h23 +0x200, Garanti eller Var.
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Radio, x304 y24 w120 h16, &Garanti
    Gui Add, Radio, x304 y40 w120 h32, G&arantivognløb i variabel tid
    Gui Add, Radio, vtype x304 y72 w120 h23, &Variabel
    Gui Font, w600
    Gui Add, Text, x16 y48 w120 h23 +0x200, &Årsag
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Edit, vårsag x16 y72 w120 h21
    Gui Font, w600
    Gui Add, Text, x8 y96 h23 +0x200, &Beskrivelse
    Gui Font
    Gui Font, s9, Segoe UI
    Gui Add, Edit, vbeskrivelse x8 y120 w410 h126
    Gui Add, CheckBox, vgemt_ja x20 y261, Brug &forrige skærmklip
    Gui Add, Button, gsvigtok x176 y256 w80 h23, &OK

    Gui Show, w448 h297, Svigt
    ControlFocus, Button1, Svigt
    mod_up()
; ^Backspace::Send +^{Left}{Backspace}
Return
svigtok:
    gui, submit
    ; MsgBox, , , % beskrivelse
    ; GuiControlGet, tid
    ; GuiControlGet, årsag
    ; GuiControlGet, beskrivelse
    ; GuiControlGet, lukket
    ; GuiControlGet, helt
    ; GuiControlGet, vl
    ; MsgBox, , Lukket kl, % tid
    ; MsgBox, , Garantitid, % tid_slet
    beskrivelse := StrReplace(beskrivelse, "`n", " ")
    if (lukket = 1 and helt = 1)
    {
        sleep 100
        MsgBox, 48 , Vælg kun én, Vælg enten lukket eller slettet VL
        sleep 100
        Gui Show, w448 h297, Svigt
        return
    }
    if (lukket = 1 and StrLen(tid) != 4)
    {
        sleep 100
        MsgBox, 48 , Klokkeslæt skal være firecifret, Klokkeslæt skal være firecifret (intet kolon).
        sleep 100
        Gui Show, w448 h297, Svigt
        SendInput, !l{tab}^a
        return
    }
    if (StrLen(tid) = 4)
    {
        timer := SubStr(tid, 1, 2)
        min := SubStr(tid, 3, 2)
        tid_tjek := A_YYYY A_MM A_DD timer min
        if tid_tjek is not Time
        {
            sleep 100
            MsgBox, 48 , Klokkeslæt ikke gyldigt , Skal være et gyldigt tidspunkt
            sleep 100
            Gui Show, w448 h297, Svigt
            SendInput, ^a
            return
        }
        tid := timer ":" min
    }
    if (helt = 1 and StrLen(tid_slet) != 4)
    {
        sleep 100
        MsgBox, 48 , Klokkeslæt for åbningstid skal være firecifret, Klokkeslæt skal være firecifret (intet kolon).
        sleep 100
        Gui Show, w448 h297, Svigt
        SendInput, !s{space}{tab}
        return
    }
    if (StrLen(tid_slet) = 4)
    {
        timer := SubStr(tid_slet, 1, 2)
        min := SubStr(tid_slet, 3, 2)
        tid_tjek := A_YYYY A_MM A_DD timer min
        if tid_tjek is not Time
        {
            sleep 100
            MsgBox, 48 , Åbningstid ikke korrekt , Klokkeslæt for åbningstid skal være et gyldigt tidspunkt
            sleep 100
            Gui Show, w448 h297, Svigt
            SendInput, !s{space}{tab}
            return
        }
        tid_slet := timer ":" min
    }
    if (type = 0)
    {
        sleep 100
        MsgBox, 48 , Mangler VL-type, Husk at krydse af i typen af VL.
        sleep 100
        Gui Show, w448 h297, Svigt
        return
    }
    if (type = 1)
        vl_type := "GV"
    if (type = 2)
        vl_type := "(Variabel tid)"
    if (type = 3)
        vl_type :=
    if (beskrivelse = "")
    {
        sleep 100
        MsgBox, 48 , Udfyld beskrivelse, Mangler beskrivelse af svigtet,
        sleep 100
        Gui Show, w448 h297, Svigt
        SendInput, !b
        return
    }
    ; MsgBox, , beskrivelse , % beskrivelse
    ; MsgBox, , type , % type
    ; MsgBox, , tid , % tid
    ; MsgBox, , årsag , % årsag
    ; MsgBox, , helt , % helt
    ; MsgBox, , vl , % dato
    if (type = 1 and lukket = 1 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 1 , % emnefelt,
        ; beskrivelse := "GV lukket kl. " tid ": " . beskrivelse
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, destroy
    }
    if (type = 1 and lukket = 1 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 2, % emnefelt,
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, destroy
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - d. " dato
        ; MsgBox, , 3, % emnefelt,
        gui, destroy
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        gui, destroy
    }
    if (type = 1 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 5, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse

        gui, destroy
    }
    if (type = 1 and helt = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - ikke startet op d. " dato
        ; MsgBox, , 5.1, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse
        gui, destroy
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag !="")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - " dato
        ; MsgBox, , 6, % emnefelt,
        gui, destroy
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " d. " dato
        ; MsgBox, , 7, % emnefelt,
        gui, destroy
    }
    if (type = 2 and lukket = 0 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 7.1, % emnefelt,
        beskrivelse := "GV slettet i variabel kørsel. Garantitid start: " tid_slet " — " . beskrivelse
        gui, destroy
    }
    if (type = 2 and lukket = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 8, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, destroy
    }
    if (type = 2 and lukket = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 9, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, destroy
    }
    if (type = 3 and årsag != "")
    {
        emnefelt := "Svigt VL " vl ": " årsag " - d. " dato
        ; MsgBox, , 10, % emnefelt,
        gui, destroy
    }
    if (type = 3 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        ; MsgBox, , 11, % emnefelt,
        gui, destroy
    }
    Outlook_nymail()
    sleep 200
    SendInput, planet
    sleep 1000
    SendInput, {enter}
    sleep 250
    SendInput, {Tab 2}
    SendInput, %emnefelt%
    sleep 400
    SendInput, {tab} %beskrivelse%
    sleep 500
    SendInput, {Enter 2}
    sleep 40
    if (gemt_ja = 1)
    {
        clipboard := gemtklip
        sleep 200
        SendInput, ^v
    }
    if (gemt_ja = 0)
    {
        clipboard := klip
        SendInput, ^v
    }
    SendInput, {Home}
    ; sleep 2000
    ; Clipboard = %gemtklip%
    ; ClipWait, 2, 1
    ; gemtklip :=
    ; MsgBox, , , % emnefelt "`n" beskrivelse
    gemtklip :=
    afslut_genvej()
Return

svigtGuiEscape:
svigtGuiClose:
    afslut_genvej()
    Gui, destroy
Return

test()
{
    WinGetTitle, tlf, ahk_exe Miralix OfficeClient.exe
    MsgBox, , , % tlf,
}

l_p6_rejsesog:
    P6_rejsesogvindue()
    afslut_genvej()
return

#IfWinActive, PLANET
::/ankc::
{
    initialer := sys_initialer()
    Input, st, , {enter}{Escape}
    Input, tid, , {enter}{Escape}
    P6_notat("st. " st " ank. " tid ", chf informerer kunde" initialer " ")
}
#IfWinActive