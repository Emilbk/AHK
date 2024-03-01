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
GroupAdd, trafikstyringsgruppe, PLANET
; GroupAdd, gruppe, ahk_class Chrome_WidgetWin_1
GroupAdd, trafikstyringsgruppe, ahk_class AccessBar
GroupAdd, trafikstyringsgruppe, ahk_class Agent Main GUI
GroupAdd, trafikstyringsgruppe, ahk_class Addressbook
GroupAdd, trafikstyringsgruppe, ahk_class Transparent Windows Client
;; lib
#Include, %A_linefile%\..\lib\AHKDb\ahkdb.ahk
#Include, %A_linefile%\..\lib\JSON.ahk
#Include, %A_linefile%\..\lib\ImagePut-master\ImagePut (for v1).ahk
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
if (brugerrække = 0)
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", "xyz", ,1)
bruger_genvej := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1) ; array med alle brugerens data
genvej_ren := []
genvej_navn := []
valg :=
vl :=
tid :=
;   1       2               3
s := bruger_genvej.41
tlf :=
trio_genvej := "Genvejsoversigt"
outlook := ComObjCreate("Outlook.application")
vl_repl := []
;; VL-liste-read
vl_liste_tekst := "db\vl_liste\" A_UserName . "_vl_liste.txt"
; tjek dato for modification, hvis ikke samme dag slet data
FileGetTime, vl_liste_tekst_dato, %vl_liste_tekst%, M
FormatTime, vl_liste_tekst_dato, vl_liste_tekst_dato, ddM
FormatTime, tid_idag, YYYYMMDDHH24MISS, ddM
if (tid_idag = vl_liste_tekst_dato)
{
    ; MsgBox, , , %vl_liste_tekst_dato% er i dag, 1
}
else
{
    ; MsgBox, , , %vl_liste_tekst_dato% er ikke i dag,

}
FileAppend, , %vl_liste_tekst%
FileRead, vl_liste_array_json, %vl_liste_tekst%
if (vl_liste_array_json = "")
    vl_liste_array := []
else if (vl_liste_array_json != "")
    vl_liste_array := json.load(vl_liste_array_json)

SetTimer, note_tjek_tid, 60000
;   bruger_genvej  telenor_opr     telenor_ahk
; FileRead, vl_repl_liste, %vl_repl_tekst%

;; hotkeydef.
; globale genveje                                           ; Standard-opsætning
Hotkey, % bruger_genvej.4, l_trio_P6_opslag ; !w
Hotkey, % bruger_genvej.30, l_trio_afslut_opkald ; Numpad -
Hotkey, % bruger_genvej.31, l_trio_afslut_opkaldB ; Numpad -
Hotkey, % bruger_genvej.32, l_trio_til_p6 ; +F4
Hotkey, % bruger_genvej.33, l_quitAHK ; +escape
Hotkey, % bruger_genvej.46, l_restartAHK ; +^r
Hotkey, % bruger_genvej.34, l_p6_aktiver ; +!p
Hotkey, % bruger_genvej.47, l_gui_hjælp ; ^½
Hotkey, % bruger_genvej.28, l_trio_opkald_markeret ; !q

Hotkey, IfWinActive, PLANET
Hotkey, % bruger_genvej.38, l_outlook_svigt ; +F1
Hotkey, % bruger_genvej.70, l_outlook_genåben ; +F1
Hotkey, % bruger_genvej.5, l_p6_initialer ; F2
Hotkey, % bruger_genvej.6, l_p6_initialer_skriv ; +F2
Hotkey, % bruger_genvej.7, l_p6_vis_k_aftale ; F3
Hotkey, % bruger_genvej.8, l_p6_ret_vl_tlf ; +F3
Hotkey, % bruger_genvej.9, l_p6_vaelg_vl ; ^F3
Hotkey, % bruger_genvej.10, l_p6_vaelg_vl ; F4
Hotkey, % bruger_genvej.61, l_p6_vaelg_vl_liste ; ^+Down
Hotkey, % bruger_genvej.11, l_p6_vl_ring_op ; +F5
Hotkey, % bruger_genvej.12, l_p6_vm_ring_op ; ^+F5
Hotkey, % bruger_genvej.13, l_p6_vl_luk ; #F5
Hotkey, % bruger_genvej.62, l_p6_laas_vl ; #F5
Hotkey, % bruger_genvej.14, l_p6_alarmer ; F7
Hotkey, % bruger_genvej.15, l_p6_udraabsalarmer ; +F7
Hotkey, % bruger_genvej.69, l_p6_billede_gui ; +F7
; Hotkey, % bruger_genvej.16, l_p6_ring_til_kunde ; +F8
Hotkey, % bruger_genvej.17, l_p6_udregn_minut ; #t
Hotkey, % bruger_genvej.18, l_p6_sygehus_ring_op ; ^+s
Hotkey, % bruger_genvej.19, l_p6_central_ring_op ; ^+c
Hotkey, % bruger_genvej.20, l_p6_tekst_til_chf ; ^+t
Hotkey, % bruger_genvej.36, l_flexf_fra_p6 ; +^F
Hotkey, % bruger_genvej.48, l_p6_rejsesog ; F1
Hotkey, % bruger_genvej.50, l_p6_liste_vl ; ^å
Hotkey, % bruger_genvej.67, l_p6_vis_liste_fra_planbillede
;Hotkey, % bruger_genvej.63, l_p6_liste_vl_notat ; ^+F10
Hotkey, % bruger_genvej.51, l_p6_vis_liste_vl ; F1
Hotkey, % bruger_genvej.55, l_p6_initialer_slet_eget ; +^n
Hotkey, % bruger_genvej.59, l_p6_initialer_skift_eget ; +^n
Hotkey, % bruger_genvej.56, l_p6_tag_alarm ; F1
Hotkey, % bruger_genvej.58, l_p6_cpr_til_bestillingsvindue ; ^F1
Hotkey, % bruger_genvej.66, l_p6_tjek_andre_rejser ; +^F
; Hotkey, % bruger_genvej.45, l_sys_inputbox_til_fra ; ^½
Hotkey, IfWinActive

Hotkey, IfWinActive, Planet Version ; specifikt alarmrepl-infobox
Hotkey, % bruger_genvej.56, l_p6_tag_alarm_vl_box ; F1
Hotkey, % bruger_genvej.49, l_p6_replaner_liste_vl ; F1
Hotkey, % bruger_genvej.60, l_p6_replaner_opslag_vl ; F1
Hotkey, IfWinActive

Hotkey, IfWinActive, Vognløbsnotering ; specifikt alarmrepl-infobox
Hotkey, % bruger_genvej.57, l_p6_notat_igen ; F1
Hotkey, IfWinActive
; Trio
Hotkey, IfWinActive, ahk_group trafikstyringsgruppe
Hotkey, % bruger_genvej.22, l_trio_pause ; ^0
Hotkey, % bruger_genvej.23, l_trio_klar ; ^1
Hotkey, % bruger_genvej.24, l_trio_udenov ; ^2
Hotkey, % bruger_genvej.25, l_trio_efterbehandling ; ^3
Hotkey, % bruger_genvej.26, l_trio_alarm ; ^4
Hotkey, % bruger_genvej.27, l_trio_frokost ; ^5
Hotkey, % bruger_genvej.64, l_trio_linie1 ; ^5
Hotkey, % bruger_genvej.65, l_trio_linie2 ; ^5
Hotkey, IfWinActive
Hotkey, % bruger_genvej.29, l_triokald_til_udklip ; #q

; flexfinder
Hotkey, IfWinActive, FlexDanmark FlexFinder ;
Hotkey, % bruger_genvej.35, l_flexf_til_p6 ; ~$^LButton
Hotkey, IfWinActive, ,
; outlook
Hotkey, % bruger_genvej.37, l_outlook_ny_mail ; ^+m

;excel
Hotkey, ifWinActive, Garantivognsoversigt FG8.xlsm
Hotkey, % bruger_genvej.39, l_excel_vl_til_P6_A ; !Lbutton
Hotkey, % bruger_genvej.40, l_excel_vl_til_P6_B ; ^w
Hotkey, IfWinActive, ,

Hotkey, ifWinActive, MD0121
Hotkey, % bruger_genvej.52, l_excel_mange_ture ; !Lbutton
Hotkey, % bruger_genvej.53, l_excel_p6_id ; !Lbutton
Hotkey, % bruger_genvej.54, l_excel_p6_cpr ; !Lbutton
Hotkey, IfWinActive, ,
;; Trio-setup
if not WinExist("ahk_class Agent Main GUI")
{
    run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Trio Enterprise\Contact Center\Agent Client.lnk"
}
if not WinExist("ahk_class AccessBar")
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Vis, Skrivebordsværktøjslinie
}
if not WinExist("ahk_class Addressbook")
{
    ControlClick, x368 y68, ahk_class Agent Main GUI , , ,, ,,
}

;; GUI
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
Gui sygehuspsyk: Add, Button, gsygehusmenu2 v78474500 x16 y32 w115 h23, RH&G psyk.
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
Gui sygehusrand: Add, Button, gsygehusmenu2 v78420000 x16 y8 w115 h23, &Randers syg.
Gui sygehusrand: Add, Button, gsygehusmenu2 v78421590 x16 y32 w115 h23, &Dialyse
Gui sygehusrand: Add, Button, gsygehusmenu2 v78475300 x16 y56 w115 h23, &Psyk.

gui sygehusvib:+Labelsygehus2
Gui sygehusvib: Font, s9, Segoe UI
Gui sygehusvib: Add, Button, gsygehusmenu2 v78440000 x16 y8 w115 h23, &Viborg syg.
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

Gui trio_genvej: Show, x1120 y3 w120 h42 w240 NA, %trio_genvej%
; Gui trio_genvej: Show, x1120 y3 w120 h42 w240 NA, %trio_genvej%

;; gui vl-liste
gui vl_liste: +labelvl_liste
gui vl_liste: font, s9, segoe ui
gui vl_liste: add, listbox, x18 y24 w200 h449 HWNDListbox1id vvalg1 gvlryd1 multi,
gui vl_liste: add, listbox, x224 y24 w200 h449 HWNDListbox2id vvalg2 gvlryd2 multi,
gui vl_liste: add, listbox, x430 y24 w200 h449 HWNDListbox3id vvalg3 gvlryd3 multi,
gui vl_liste: add, listbox, x636 y24 w200 h449 HWNDListbox4id vvalg4 gvlryd4 multi,
gui vl_liste: add, listbox, x842 y24 w250 h449 HWNDListbox5id vvalg5 gvlryd5 multi,
gui vl_liste: add, listbox, x1098 y24 w200 h449 HWNDListbox6id vvalg6 gvlryd6 multi,
gui vl_liste: add, text, x18 y0 w120 h23 +0x200, &Replaneret
gui vl_liste: add, text, x224 y0 w120 h23 +0x200, Wakeup
gui vl_liste: add, text, x430 y0 w120 h23 +0x200, Privatrejse
gui vl_liste: add, text, x636 y0 w120 h23 +0x200, Huskeliste
gui vl_liste: add, text, x842 y0 w120 h23 +0x200, Låst
gui vl_liste: add, text, x1097 y0 w120 h23 +0x200, Kvitteret for chauffør
gui vl_liste: add, button, x78 y475 w80 h23 gvl_liste_ryd1 vbox1, ryd
gui vl_liste: add, button, x284 y475 w80 h23 gvl_liste_ryd2 vbox2, ryd
gui vl_liste: add, button, x490 y475 w80 h23 gvl_liste_ryd3 vbox3, ryd
gui vl_liste: add, button, x696 y475 w80 h23 gvl_liste_ryd4 vbox4, ryd
gui vl_liste: add, button, x927 y475 w80 h23 gvl_liste_ryd5 vbox5, ryd
gui vl_liste: add, button, x1158 y475 w80 h23 gvl_liste_ryd6 vbox6, ryd
gui vl_liste: add, button, x180 y536 w131 h23 gvl_liste_tilføj_note, tilføj/vis &note
gui vl_liste: add, button, x320 y536 w131 h23 gvl_liste_OBS, OBS
gui vl_liste: add, button, x460 y536 w131 h23 gvl_liste_opslag, &opslag
gui vl_liste: add, button, x600 y536 w131 h23 gvl_liste_opslag_slet, opslag og s&let
gui vl_liste: add, button, x740 y536 w131 h23 gvl_liste_slet, &slet
gui vl_liste: add, button, x880 y536 w131 h23 gvl_liste_slet_alt_alle, slet &alt
gui vl_liste: add, button, x1020 y536 w131 h23 gvl_liste_liste, l&iste

;; GUI P6-billeder
Gui p6_billede: Font, s9, Segoe UI
Gui p6_billede: Add, Radio, x16 y32 w181 h23 vp6_billede_adresse, &Adresse
Gui p6_billede: Add, Radio, x16 y56 w180 h23 vp6_billede_vg, Vogngruppeskema, &dag
Gui p6_billede: Add, Radio, x16 y80 w181 h23 vp6_billede_styringsystem, &Styringssystem
Gui p6_billede: Add, Radio, x16 y104 w178 h23 vp6_billede_vogngruppe, &Vogngruppe
Gui p6_billede: Add, Radio, x16 y128 w179 h23 vp6_billede_vogngrupppe_fast, Vogngruppeskema, &fast
Gui p6_billede: Add, Radio, x16 y152 w175 h23 vp6_billede_liste_vl, &Liste Vognløb
Gui p6_billede: Add, Radio, x16 y176 w178 h23 vp6_billede_betaler, &Betaler
Gui p6_billede: Add, Radio, x16 y200 w178 h23 vp6_billede_lange_rejser, Liste lange &rejser
Gui p6_billede: Add, Button, x40 y272 w42 h23 gp6_billede_ok, &OK
Gui p6_billede: Add, Button, x104 y272 w55 h23 gp6_billedeescape, Afbryd
Gui p6_billede: Add, Text, x16 y8 w120 h23 +0x200, Hvilket billede vil du se?

gui note: +Labelnote
Gui note: Font, s9, Segoe UI
Gui note: Add, Edit, x16 y8 w438 h206 vnote_note
Gui note: Add, Button, x8 y240 w80 h23 gnote_ok, &Gem
Gui note: Add, Button, x104 y240 w80 h23 gnote_opslag, &Opslag VL
Gui note: Add, Button, x200 y240 w80 h23 gnote_slet, &Slet note
Gui note: Add, checkbox, x319 y240 w76 h21 vnote_reminder, &Reminder
Gui note: Add, Edit, x394 y240 w50 h21 number vnote_tid, 

;; GUI vl-note

;; END AUTOEXEC
Return

p6_billede_ok:
gui p6_billede: Submit
if (p6_billede_adresse = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{esc}{alt}", "gga")
        return
    }
if (p6_billede_vg = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "td")
        return
    }
if (p6_billede_styringsystem = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "ts")
        return
    }
if (p6_billede_vogngruppe = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "geg")
        return
    }
if (p6_billede_vogngruppe_fast = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "gef")
        return
    }
if (p6_billede_liste_vl = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "tv")
        return
    }
if (p6_billede_betaler = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "gøb")
        return
    }
if (p6_billede_lange_rejser = 1)
    {
        KeyWait, alt
        P6_aktiver()
        P6_alt_menu("{alt}", "ti")
        return
    }
return

p6_billedeEscape:
p6_billedeClose:
gui p6_billede: hide
Return


#IfWinActive VL-liste
    Enter::
    NumpadEnter::
        fokus := GUIfokus()
        if (InStr(fokus, "listbox"))
        {
            Gosub, vl_liste_opslag
            return
        }
        Else
        {
            SendInput, {enter}
            return
        }
    +enter::
        {
            fokus := GUIfokus()
            if (InStr(fokus, "listbox"))
            {
                Gosub, vl_liste_opslag_slet
                return
            }
        }
#IfWinActive

; Omskriv
#IfWinActive, PLANET
#IfWinActive, VL-liste
    F8::
        gui vl_liste: Hide
    return
#IfWinActive

#IfWinActive, VL-liste
    F5::
        vlListe_opdater_gui()
    return

    w::
        {
            GuiControl, vl_liste: Focus, listbox2
            GuiControl, vl_liste: Choose, Listbox2, 1
            Gosub, vlryd2
            return
        }

    h::
        {
            GuiControl, vl_liste: Focus, listbox4
            GuiControl, vl_liste: Choose, Listbox4, 1
            Gosub, vlryd4
            return
        }

    å::
        {
            GuiControl, vl_liste: Focus, listbox5
            GuiControl, vl_liste: Choose, Listbox5, 1
            Gosub, vlryd5
            return
        }

    r::
        {
            GuiControl, vl_liste: Focus, listbox1
            GuiControl, vl_liste: Choose, Listbox1, 1
            Gosub, vlryd1
            return
        }

    p::
        {
            GuiControl, vl_liste: Focus, listbox3
            GuiControl, vl_liste: Choose, Listbox3, 1
            Gosub, vlryd3
            return
        }
    k::
        {
            GuiControl, vl_liste: Focus, listbox6
            GuiControl, vl_liste: Choose, Listbox6, 1
            Gosub, vlryd6
            return
        }
#IfWinActive
;; gui-label vl-list
vl_listeescape:
vl_listeclose:
    gui vl_liste: hide
return

vlryd1:
    GuiControl, vl_liste: Choose, Listbox2 , 0
    GuiControl, vl_liste: Choose, Listbox3 , 0
    GuiControl, vl_liste: Choose, Listbox4 , 0
    GuiControl, vl_liste: Choose, Listbox5 , 0
    GuiControl, vl_liste: Choose, Listbox6 , 0
return
vlryd2:
    GuiControl, vl_liste: Choose, Listbox1 , 0
    GuiControl, vl_liste: Choose, Listbox3 , 0
    GuiControl, vl_liste: Choose, Listbox4 , 0
    GuiControl, vl_liste: Choose, Listbox5 , 0
    GuiControl, vl_liste: Choose, Listbox6 , 0
return
vlryd3:
    GuiControl, vl_liste: Choose, Listbox1 , 0
    GuiControl, vl_liste: Choose, Listbox2 , 0
    GuiControl, vl_liste: Choose, Listbox4 , 0
    GuiControl, vl_liste: Choose, Listbox5 , 0
    GuiControl, vl_liste: Choose, Listbox6 , 0
return
vlryd4:
    GuiControl, vl_liste: Choose, Listbox1 , 0
    GuiControl, vl_liste: Choose, Listbox2 , 0
    GuiControl, vl_liste: Choose, Listbox3 , 0
    GuiControl, vl_liste: Choose, Listbox5 , 0
    GuiControl, vl_liste: Choose, Listbox6 , 0
return
vlryd5:
    GuiControl, vl_liste: Choose, Listbox2 , 0
    GuiControl, vl_liste: Choose, Listbox3 , 0
    GuiControl, vl_liste: Choose, Listbox4 , 0
    GuiControl, vl_liste: Choose, Listbox1 , 0
    GuiControl, vl_liste: Choose, Listbox6 , 0
return
vlryd6:
    GuiControl, vl_liste: Choose, Listbox2 , 0
    GuiControl, vl_liste: Choose, Listbox3 , 0
    GuiControl, vl_liste: Choose, Listbox4 , 0
    GuiControl, vl_liste: Choose, Listbox1 , 0
    GuiControl, vl_liste: Choose, Listbox5 , 0
return
vl_liste_ryd1:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox1")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
return
vl_liste_ryd2:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox2")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
return
vl_liste_ryd3:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox3")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
return
vl_liste_ryd4:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox4")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
return
vl_liste_ryd5:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox5")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
return
vl_liste_ryd6:
    vl_liste_midl := []
    for i,e in vl_liste_array
    { for i2,e2 in e
        if (i2 = 8 and e2 != "listbox6")
        {
            ;    vl_liste_array.RemoveAt(i)
            vl_liste_midl[i] := vl_liste_array[i]
        }
    }
    vl_liste_array := vl_liste_midl
    vl_liste_array_til_json_tekst()
    vlListe_opdater_gui()
Return

noteEscape:
noteClose:
WinGetActiveTitle, note_vinduetitel
vlListe_opdater_gui()
gui note: hide
sleep 100
if (InStr(note_vinduetitel, "reminder") or InStr(note_vinduetitel, "liste"))
    {
        return
    }
gui vl_liste: show
Return

vl_liste_add_note(valg, note_note)
{
    tid := StrSplit(valg, ",")
tid := Regexreplace(tid[2], "\D")
tid := SubStr(tid, 1, 2) ":" SubStr(tid, 3, 2)
valg := SubStr(valg, 1, 5)  
valg := Regexreplace(valg, "\D")
note_note := 
GuiControl, note:, note_note , %note_note%
for i,e in vl_liste_array
    {
        if (vl_liste_array[i][1] = valg and vl_liste_array[i][8] = listbox and SubStr(vl_liste_array[i][3], 1, 5) = tid)
            {
                note_note := vl_liste_array[i][5]
                GuiControl, note:, edit2, % vl_liste_array[i][10]
            if (vl_liste_array[i][10] = "")
            {
                GuiControl, note:, note_reminder, 0
            }            
            Else
                {
                GuiControl, note:, note_reminder, 1
                }
                break
            }   
    }
GuiControl, note:, note_note, %note_note%
sleep 100
Gui note: Show, w477 h277, Note VL %valg%
ControlFocus, Edit1 , Note
sleep 100

return
}
;; lav
vl_liste_tilføj_note:
; Note-gui
    valg :=
    Gui vl_liste: Submit
    Gui vl_liste: Hide
    for i,e in [valg1, valg2, valg3, valg4, valg5, valg6]
        {
            if (e != "")
                {
                    valg := e
                    listbox := "listbox" . i
                    break
                }
        }
; vl_liste_add_note(valg, note_note)

tid := StrSplit(valg, ",")
tid := Regexreplace(tid[2], "\D")
tid := SubStr(tid, 1, 2) ":" SubStr(tid, 3, 2)
valg := SubStr(valg, 1, 5)  
valg := Regexreplace(valg, "\D")
note_note := 
GuiControl, note:, note_note , %note_note%
for i,e in vl_liste_array
    {
        if (vl_liste_array[i][1] = valg and vl_liste_array[i][8] = listbox and SubStr(vl_liste_array[i][3], 1, 5) = tid)
            {
                note_note := vl_liste_array[i][5]
                GuiControl, note:, edit2, % vl_liste_array[i][10]
            if (vl_liste_array[i][10] = "")
            {
                GuiControl, note:, note_reminder, 0
            }            
            Else
                {
                GuiControl, note:, note_reminder, 1
                }
                break
            }   
    }
GuiControl, note:, note_note, %note_note%
sleep 100
Gui note: Show, w477 h277, Note VL %valg%
ControlFocus, Edit1 , Note
sleep 100

Return
note_ok:
gui note: submit
vlListe_note(note_reminder, note_tid, note_note)
;     if (note_reminder = 1 and StrLen(note_tid) != 4)
;         {
;             MsgBox, 16 , Fejl i indtastet tidspunkt, Der skal bruges fire tal, i formatet TTMM (f. eks. 1434).
;             gui note: Show
;             return 
;         }
;     note_tid_tjek := A_YYYY A_MM A_DD note_tid
;         if note_tid_tjek is not Time
;         {
;             MsgBox, 16 , Fejl i indtastning af tidspunkt , Det indtastede er ikke et klokkeslæt.,
;             gui note: show
;             return
;         }
;     for i,e in vl_liste_array
;     if (vl_liste_array[i][8] = listbox and vl_liste_array[i][1] = valg and SubStr(vl_liste_array[i][3], 1, 5) = tid) 
;         {
;             if (note_reminder = 1)
;                 {
;                     vl_liste_array[i][10] := note_tid
;                     vl_liste_array[i][7] := " (R)"
;                 }
;             vl_liste_array[i][6] := " (N)"
;             vl_liste_array[i][5] := note_note
;             if (note_note = "")
;                 vl_liste_array[i][6] := ""
;             gui note: hide
;             vl_liste_array_til_json_tekst()
;             P6_aktiver()
;             return
;         }
; P6_aktiver()
if (note_reminder = 1)
{
    vlliste_note_reminder(note_reminder, note_tid)
}
return


note_opslag:
gui note: hide
p6_vaelg_vl(valg)
return

note_tjek_tid:
FormatTime, note_tid_ur, YYYYMMDDHH24MISS, HHmm
note_vl := [[]]
note_count := 0
for i,e in vl_liste_array
    {
        if (vl_liste_array[i][10] = note_tid_ur)
            {
                note_count += 1
                note_vl[note_count].Push(vl_liste_array[i][1])
                note_vl[note_count].push(vl_liste_array[i][5])
                vl_liste_array[i][10] := ""
                vl_liste_array[i][7] := ""
            }
    }
for i,e in note_vl
    {
    if (e[i] != "")
    {
        vl := e[1]
        note := e[2]
        Gui note: show, ,Reminder vognløb %vl%
        ControlFocus, edit2
        GuiControl, note:, Edit1, %note%
        GuiControl, note:, Edit2, 
        GuiControl, note:, note_reminder, 0 
        
    }
}
return

vl_liste_obs:

    vl_liste_opslag_array := []
    valg :=
    Gui vl_liste: Submit, NoHide
    vl_liste_opslag_array.Push(valg1, valg2, valg3, valg4, valg5, valg6)
    for i,e in [valg1, valg2, valg3, valg4, valg5, valg6]
        {
            if (i != "")
                listbox := listbox . i
        }
    for i,e in vl_liste_opslag_array
        if (e != "")
        {
            valg := vl_liste_opslag_array[i]
        }
    if (valg = "")
    {
        MsgBox, , Vælg en markering, Der skal laves en markering, 2
        gui vl_liste: show
        return
    }
    vl_liste_valg_vl := StrSplit(valg, ",")
    tid_ind := RegExReplace(vl_liste_valg_vl.2, "\D")
    tid_korrigeret := SubStr(tid_ind, 1, 2) ":" SubStr(tid_ind, 3 , 2)
    ; tjek på indkomne listbox(hvordan?), vl og tid - slet i array

    for i,e in vl_liste_array
        for i2,e2 in e
        {
            if (i2 = 1 and e2 = vl_liste_valg_vl.1 and SubStr(e.3, 1,5) = tid_korrigeret and !InStr(vl_liste_array[i][3], "(!)"))
            {
                vl_liste_array[i][3] := vl_liste_array[i][3] . " (!)"
                vl_liste_array_til_json_tekst()
                vlListe_opdater_gui()
                return
            }
            if (i2 = 1 and e2 = vl_liste_valg_vl.1 and SubStr(e.3, 1,5) = tid_korrigeret and InStr(vl_liste_array[i][3], "(!)"))
            {
                vl_liste_array[i][3] := SubStr(vl_liste_array[i][3], 1 , -4)
                vl_liste_array_til_json_tekst()
                vlListe_opdater_gui()
                return
            }
}
return


note_slet:
{
    for i,e in vl_liste_array
    if (vl_liste_array[i][8] = listbox and vl_liste_array[i][1] = valg and SubStr(vl_liste_array[i][3], 1, 5) = tid)
        {
            vl_liste_array[i](6) := ""
            vl_liste_array[i][5] := ""
            vl_liste_array_til_json_tekst()
            gui note: hide
            P6_aktiver()
            return
        }
return
}

vl_liste_vis_note:
Return
vl_liste_liste:
    vl_liste_opslag_array := []
    valg :=
    array := 0
    Gui vl_liste: Submit
    Gui vl_liste: Hide
    if (InStr(valg1, "|"))
    {
        valg1 := StrSplit(valg1, "|")
        array := 1
    }
    if (InStr(valg2, "|"))
    {
        valg2 := StrSplit(valg2, "|")
        array := 1
    }
    if (InStr(valg3, "|"))
    {
        valg3 := StrSplit(valg3, "|")
        array := 1
    }
    if (InStr(valg4, "|"))
    {
        valg4 := StrSplit(valg4, "|")
        array := 1
    }
    if (InStr(valg5, "|"))
    {
        valg5 := StrSplit(valg5, "|")
        array := 1
    }
    vl_liste_opslag_array.Push(valg1, valg2, valg3, valg4, valg5)
    for i,e in vl_liste_opslag_array
        if (e != "")
        {
            valg := vl_liste_opslag_array[i]
        }
    if (valg = "")
    {
        sleep 500
        MsgBox, , Vælg en markering, Der skal laves en markering, 2
        sleep 500
        gui vl_liste: show
        return
    }
    if (array = 0)
    {
        listevl_array := []
        for i,e in valg
            listevl_array[i] := StrSplit(valg[i], ",")
    }
    if (array = 1)
    {
        listevl_array := []
        for i,e in valg
            listevl_array[i] := StrSplit(valg[i], ",")
    }
Return

vl_liste_opslag:
    vl_liste_opslag_array := []
    valg :=
    array := 0
    Gui vl_liste: Submit
    Gui vl_liste: Hide
    if (InStr(valg1, "|"))
    {
        valg1 := StrSplit(valg1, "|")
        array := 1
    }
    if (InStr(valg2, "|"))
    {
        valg2 := StrSplit(valg2, "|")
        array := 1
    }
    if (InStr(valg3, "|"))
    {
        valg3 := StrSplit(valg3, "|")
        array := 1
    }
    if (InStr(valg4, "|"))
    {
        valg4 := StrSplit(valg4, "|")
        array := 1
    }
    if (InStr(valg5, "|"))
    {
        valg5 := StrSplit(valg5, "|")
        array := 1
    }
    if (InStr(valg6, "|"))
    {
        valg6 := StrSplit(valg6, "|")
        array := 1
    }
    vl_liste_opslag_array.Push(valg1, valg2, valg3, valg4, valg5, valg6)
    for i,e in vl_liste_opslag_array
        if (e != "")
        {
            valg := vl_liste_opslag_array[i]
        }
    if (valg = "")
    {
        sleep 50
        MsgBox, , Vælg en markering, Der skal laves en markering, 2
        gui vl_liste: show
        return
    }
    if (array = 0)
    {
        vl_liste_valg_vl := StrSplit(valg, ",")
        p6_vaelg_vl(vl_liste_valg_vl[1])
    }
    if (array = 1)
    {
        listevl_array := []
        for i,e in valg
            listevl_array[i] := StrSplit(valg[i], ",")
    }
Return
vl_liste_opslag_slet:
    vl_liste_opslag_array := []
    valg :=
    Gui vl_liste: Submit
    vl_liste_opslag_array.Push(valg1, valg2, valg3, valg4, valg5)
    for i,e in vl_liste_opslag_array
        if (e != "")
        {
            valg := vl_liste_opslag_array[i]
        }
    if (valg = "")
    {
        MsgBox, , Vælg en markering, Der skal laves en markering, 2
        gui vl_liste: show
        return
    }
    vl_liste_valg_vl := StrSplit(valg, ",")
    tid_ind := RegExReplace(vl_liste_valg_vl.2, "\D")
    tid_korrigeret := SubStr(tid_ind, 1, 2) ":" SubStr(tid_ind, 3 , 2)
    ; tjek på indkomne listbox(hvordan?), vl og tid - slet i array

    for i,e in vl_liste_array
        for i2,e2 in e
        {
            if (i2 = 1 and e2 = vl_liste_valg_vl.1 and SubStr(e.3, 1,5) = tid_korrigeret)
            {
                vl_liste_array.RemoveAt(i)
                vl_liste_array_til_json_tekst()
                p6_vaelg_vl(vl_liste_valg_vl.1)
                return
            }
        }

Return
vl_liste_slet:
    vl_liste_opslag_array := []
    valg :=
    Gui vl_liste: Submit, NoHide
    vl_liste_opslag_array.Push(valg1, valg2, valg3, valg4, valg5)
    for i,e in vl_liste_opslag_array
        if (e != "")
        {
            valg := vl_liste_opslag_array[i]
            listbox := "listbox" . i
        }
    if (valg = "")
    {
        MsgBox,, Vælg en markering, Der skal laves en markering, 2
        gui vl_liste: show
        return
    }
    ; tjek for multivalg
    if (InStr(valg, "|"))
        {
            valg_split := StrSplit(valg, "|")
            for i,e in valg_split
                {
                    valg_split[i] := SubStr(valg_split[i], 1, 5)
                    valg_split[i] := RegExReplace(valg_split[i], "\D")

                }
        for i,e in valg_split
            {
                valg := valg_split[i]
                for i,e in vl_liste_array
                    {
                        if (vl_liste_array[i][1] = valg and vl_liste_array[i][8] = listbox)
                            {
                                vl_liste_array.RemoveAt(i)
                            }
                    }
            }
        vl_liste_array_til_json_tekst()
        vlListe_opdater_gui()
        return
        }
    vl_liste_valg_vl := StrSplit(valg, ",")
    tid_ind := RegExReplace(vl_liste_valg_vl.2, "\D")
    tid_korrigeret := SubStr(tid_ind, 1, 2) ":" SubStr(tid_ind, 3 , 2)
    ; tjek på indkomne listbox(hvordan?), vl og tid - slet i array

    for i,e in vl_liste_array
        for i2,e2 in e
        {
            if (i2 = 1 and e2 = vl_liste_valg_vl.1 and SubStr(e.3, 1,5) = tid_korrigeret)
            {
                vl_liste_array.RemoveAt(i)
                vl_liste_array_til_json_tekst()
                vlListe_opdater_gui()
                return
            }
        }

Return
vl_liste_slet_alt:
Return
vl_liste_slet_alt_alle:
    gui vl_liste: hide
    FileDelete, %vl_liste_tekst%
    FileAppend, , %vl_liste_tekst%
    vl_liste_array := []
Return
;; end autoexec
return
;; GUI-labels
trio_genvej:
    GetKeyState, tjek_key, Shift, 
    if (tjek_key = "D")
        {
            SendInput, {ShiftUp}
            goto l_restartAHK
        }
    MsgBox, , Knap, Knap, 
    return
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
    tjek := Trio_opkald(knap2)
    if (tjek = 0)
    {
        sys_afslut_genvej()
        return
    }
    gui cancel
    WinActivate, PLANET
    sleep 3000
    trio_klar()
    sys_afslut_genvej()
return

sygehusEscape:
sygehusClose:
    gui cancel
    sys_afslut_genvej()
return

sygehus2Escape:
sygehus2Close:
    vis_sygehus_1()
    gui Cancel
    sys_afslut_genvej()
return

; GUI repl
replguiEscape:
replguiClose:
    gui, hide
return

replvl_opslag:
    Gui, Submit
    gui, hide
    valgt_vl := StrSplit(valg, ",")
    p6_vaelg_vl(valgt_vl.1)
return

replvl_opslag_slet:
    Gui, Submit
    gui, hide
    valgt_vl := StrSplit(valg, ",")
    for k, v in vl_repl
        if (valg = v)
            vl_repl.removeat(k)
    vl_repl_liste := "|"
    for k, v in vl_repl
        vl_repl_liste .= vl_repl[k] . "|"
    FileDelete, %vl_repl_tekst%
    FileAppend, %vl_repl_liste%, %vl_repl_tekst%
    p6_vaelg_vl(valgt_vl.1)
return

replslet:
    Gui, Submit
    gui, hide
    FileRead, vl_repl_liste, %vl_repl_tekst%
    vl_repl_split := StrSplit(vl_repl_liste, "|")
    for k, v in vl_repl_split
        if (valg = v)
            vl_repl.removeat(k)

    vl_repl_liste := "|"
    for k, v in vl_repl
        vl_repl_liste .= vl_repl[k] . "|"
    FileDelete, %vl_repl_tekst%
    FileAppend, %vl_repl_liste%, %vl_repl_tekst%
return

replsletalt:
    Gui, Submit
    gui, hide
    vl_repl := []
    vl_repl_liste := "|"
    for k, v in vl_repl
        vl_repl_liste .= vl_repl[k] . "|"
    FileDelete, %vl_repl_tekst%
    FileAppend, %vl_repl_liste%, %vl_repl_tekst%
return

replvl:
    gui Submit
    gui Hide
    p6_vaelg_vl(%valg%)
return
; P6_billede-labels
p6_billede_adresse:
P6_aktiver()
SendInput, !gga
return
p6_billede_vg:
P6_aktiver()
SendInput, !td
return
;; FUNKTIONER
;; P6
sys_afslut_genvej()
{
    GuiControl, trio_genvej:text, Button1, Genvejsoversigt
    mod_up()
    return
}

sys_genvej_beskrivelse(kolonne)
{
    trio_genvej := databaseget("%A_linefile%\..\db\bruger_ops.tsv", 3, kolonne)
    GuiControl, trio_genvej:text, Button1, %trio_genvej%
    return trio_genvej
}
; henter GUI-control, der har fokus
GUIfokus()
{
    ControlGetFocus, GUIfokus
    return GUIfokus
}
sys_genvej_start(kolonne)
{
    genvej_mod := sys_genvej_til_ahk_tast(kolonne)
    sys_genvej_keywait(genvej_mod)
    sys_genvej_beskrivelse(kolonne)
    return
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
    ; SendInput, {esc}
    sleep 20
    Sendinput %tast1%
    sleep 40
    SendInput, %tast2%
    sleep 40
}
; ***
; Åben planbillede
P6_planvindue()
{
    global s
    P6_aktiver()
    P6_alt_menu("{esc}{alt}", "tp")
}

; ***
; Åben renset rejsesøg
P6_rejsesogvindue(byref telefon := "")
{
    global s
    P6_alt_menu("{alt}", "rr")
    sleep s * 10 + 100
    if (telefon = "")
        return
    SendInput, ^t
    SendInput {tab}{tab}
    SendInput, %telefon%
    SendInput, ^r
    Return
}

p6_bestillingsvindue()
{
    P6_aktiver()
    P6_alt_menu("{alt}", "rb")
    return
}

;  ***
; Vis kørselsaftale for aktivt planbillede
P6_vis_k()
{
    global s
    P6_alt_menu("{alt},", "tk")
    sleep s * 40
    SendInput !{F5}
    return
}

; Tjek for om der er skrevet nyt VL-notat i mellemtiden
;         clipboard :=
; SendInput, {enter}
; sleep 80
; SendInput, ^a^c
; ClipWait, 2
; ny_tekst := clipboard
; SendInput, !a
; sleep 10
; SendInput, ^n
; sleep 200
; clipboard :=
; SendInput, ^a^c
; clipwait, 2
; gammel_tekst := clipboard
; sleep 40
; ; MsgBox, , , % gammel_tekst
; if gammel_tekst = %ny_tekst%
; {
;     p6_notat(ny_tekst, 1)
;     return
; }
; Else
; {
;     MsgBox, , , Der er kommet et nyt notat på vl
;     ; SendInput, {tab}{enter}
;     ; sleep 100
;     ; P6_notat(clipboard, 1)
;     Return
; }
#IfWinActive
; Kørselsaftale på VL til clipboard
P6_hent_k()
{
    global s
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    ; Sendinput !tp!k
    P6_planvindue()
    SendInput, !k
    clipboard := ""
    Sendinput +{F10}c
    ClipWait 1
    sleep s * 200
    loop_test := 0
    while (clipboard = "")
    {
        P6_planvindue()
        SendInput, !k
        clipboard := ""
        Sendinput +{F10}c
        ClipWait 1
        sleep s * 400
    }
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
    Sendinput !k{tab}
    clipboard := ""
    Sendinput +{F10}c
    ClipWait 1
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
    P6_planvindue()
    SendInput, !l
    clipboard := ""
    sleep 50 ; ikke P6-afhængig
    SendInput, +{F10}c
    ClipWait, 1, 0
    vl := clipboard
    loop_test := 0
    while (vl = "")
    {
        P6_planvindue()
        SendInput, !l
        sleep 500
        SendInput, +{F10}c
        ClipWait, 1, 0
        vl := clipboard
        loop_test += 1
        if (loop_test > 7)
        {
            MsgBox, 16, Fejl, Der er sket en fejl - Prøv igen
            return 0
        }
    }
    return vl
}

p6_vl_vindue()
{
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 30
    SendInput, ^{F12}
    sleep 150
    clipboard :=
    SendInput, ^c
    clipwait 0.3
    if (InStr(clipboard, "opdateringern")) ; tjek for tidligere vl-vindue stadig åbent. OBS ikke slåfejl
    {
        SendInput, !y
    }
    clipboard :=
    loop_test := 0
    Send, +{F10}c
    ClipWait, 1
    vl_opslag := clipboard
    while (vl_opslag != vl)
    {
        SendInput, !l
        Send, +{F10}c
        ClipWait, 1
        vl_opslag := clipboard
        sleep 400
        loop_test += 1
        if (loop_test > 15)
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
        return "lukket" ; VL lukket
    }
    clipboard :=
    SendInput, +{F10}c
    clipwait 0.5
    loop_test := 0
    while (clipboard = "")
    {
        SendInput, !k
        sleep 400
        clipboard :=
        SendInput, +{F10}c
        clipwait 1.5
        loop_test += 1
        if (loop_test > 10)
        {
            MsgBox, 16 , Fejl, Der er sket en fejl - Prøv igen
            return 0
        }
    }
    k_aftale.1 := clipboard
    clipboard :=
    ; tjek om drift eller vogngruppe
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

        sleep 200
        SendInput, {enter}
    }
    return
}
p6_vaelg_vl_liste(byref vl := "")
{
    global listevl_array

    vl := listevl_array[1][1]

    P6_Planvindue()
    SendInput, !l
    if (listevl_array.MaxIndex() = "")
    {
        MsgBox, , Liste VL, Listen er tom, 2
        return
    }
    if (vl != "")
    {
        listevl_array.RemoveAt(1)
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

    P6_planvindue()
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

    P6_planvindue()
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
; tager 1, tekst og 2, tjek for "ny notat-funktion" (hvis = 1)
P6_notat(byref tekst:="", tjek:="")
{
    if (tjek = 1)
    {
        SendInput, ^n
        sleep 500
        SendInput, ^a
        sleep 40
        sendraw, %tekst%
        SendInput, !o
    }
    Else
    {
        P6_planvindue()
        SendInput, ^n
        sleep 500
        sendraw, %tekst%
        SendInput, !o
    }
    Return

}
p6_notat_igen()
{
    clipboard :=
    SendInput, {enter}
    sleep 80
    SendInput, ^a^c
    ClipWait, 2
    ny_tekst := clipboard
    sleep 40
    SendInput, !a
    sleep 40
    p6_notat(ny_tekst, 1)
    Return
}
p6_notat_hotstr(notat := "")
{
    initialer := sys_initialer()
    if (InStr(notat, "st."))
    {
        Input, st, , {enter}{Escape}
        if (ErrorLevel = "Endkey:Escape")
            return
        if (st = "")
            {
                MsgBox, 16, Intet input, Intet input - er numlock slået til?
                return
            }
    }
    if (InStr(notat, "ankomst_tid"))
    {
        Input, ankomst_tid, , {enter} {Escape}
        if (ErrorLevel = "Endkey:Escape")
            return
        if (ankomst_tid = "")
            {
                MsgBox, 16, Intet input, Intet input - er numlock slået til?
                return
            }
    }
    if (InStr(notat, "repl_tid"))
    {
        Input, repl_tid, , {enter} {Escape}
        if (ErrorLevel = "Endkey:Escape")
            return
        if (repl_tid = "")
            {
                MsgBox, 16, Intet input, Intet input - er numlock slået til?
                return
            }
    }
    if (InStr(notat, "pause_tid"))
    {
        Input, pause_tid, , {enter} {Escape}
        if (ErrorLevel = "Endkey:Escape")
            return
        if (pause_tid = "")
            {
                MsgBox, 16, Intet input, Intet input - er numlock slået til?
                return
            }
    }

    notat := StrReplace(notat, "st." , "st. " st)
    notat := StrReplace(notat, "ankomst_tid" , "ank. " ankomst_tid)
    notat := StrReplace(notat, "repl_tid" , "" repl_tid)
    notat := StrReplace(notat, " initialer" , initialer " ")
    notat := StrReplace(notat, "pause_tid" , pause_tid " ")
    P6_notat(notat)
    ; "st. " %st% " ank. " tid ", chf informerer kunde" initialer " "
    return
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

p6_cpr_til_bestillingsvindue()
{
    clipboard :=
    SendInput, ^a^c
    ClipWait, 1
    cpr := clipboard
    p6_bestillingsvindue()
    sleep 100
    SendInput, ^t^v{enter}
    return
}

p6_tjek_andre_rejser()
{
    SendInput, ^{F9}
    sleep 200
    SendInput, !r{F5}{down}
    return
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
        MsgBox, 16 , For lang tid brugt, Noget er gået galt. Prøv igen.
        sys_afslut_genvej()
        return 0
    }
    vl_tilstand := p6_vl_vindue_edit()
    if (vl_tilstand = "lukket")
    {
        sleep 100
        MsgBox, , Vl er lukket, Kan ikke trække telefonnummer, vl er afsluttet
        sys_afslut_genvej()
        return 0
    }
    if (vl_tilstand = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 100
    SendInput {Enter}{Enter}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    Clipboard :=
    SendInput {tab}{tab}
    loop_test := 0
    clipboard :=
    SendInput ^c
    ClipWait, 1
    while (StrLen(clipboard) != 8)
    {
        SendInput, !ø{tab 2}
        clipboard :=
        SendInput ^c
        ClipWait, 1
        sleep 400
        loop_test += 1
        if (loop_test > 5)
        {
            MsgBox, 16, Fejl, Der er sket en fejl - Prøv igen
            return 0
        }
    }
    SendInput ^a
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
    SendInput {tab 4}
    loop_test := 0
    clipboard :=
    SendInput ^c
    ClipWait, 1.5
    while (StrLen(clipboard) != 8)
    {
        SendInput, !a{tab4}
        clipboard :=
        SendInput ^c
        ClipWait, 1.5
        sleep 400
        loop_test += 1
        if (loop_test > 10)
        {
            MsgBox, 16, Fejl, Der er sket en fejl - Prøv igen
            return "fejl"
        }
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
    vl := p6_vl_vindue()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
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
P6_initialer_slet_eget()
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
    clipwait 3, 0
    notering := Clipboard
    notering_split := StrSplit(notering, initialer_udentid)
    if (InStr(notering_split.1, "/"))
    {
        SendInput, !o
        return
    }
    notering_split.RemoveAt(1)
    notering_split.1 := SubStr(notering_split.1, 5)
    for i, e in notering_split
    {
        if (i > 1)
        {
            notering_split[i] := initialer_udentid . notering_split[i]
        }
    }
    for i, e in notering_split
    {
        notering_endelig := notering_endelig . notering_split[i]
    }
    clipboard := % notering_endelig
    if (clipboard = "")
        clipboard := " "
    SendInput, ^v
    sleep 50
    SendInput, !o
    sleep 200
    clipboard := notering

}
P6_initialer_skift_eget()
{
    global s
    initialer := sys_initialer()
    initialer_udentid := "/mt" A_userName

    Input, oprindelig_tekst, , {escape} {enter},
    Input, ny_tekst, , {escape} {enter},
    SendInput, {F5} ; for at undgå timeout. Giver det problemer med langsom opdatering?
    sleep s * 40
    sendinput ^n
    sleep s * 1400
    clipboard :=
    SendInput, ^a^c
    ClipWait, 1, 0
    notering := Clipboard
    clipwait 3, 0
    notering_split := StrSplit(notering, initialer_udentid)
    notering_split.1 := StrReplace(notering_split.1, oprindelig_tekst , ny_tekst )
    if (InStr(notering_split.1, "/"))
    {
        SendInput, !o
        return
    }
    for i, e in notering_split
    {
        if (i > 1)
        {
            notering_split[i] := initialer_udentid . notering_split[i]
        }
    }
    for i, e in notering_split
    {
        notering_endelig := notering_endelig . notering_split[i]
    }
    clipboard := % notering_endelig
    SendInput, ^v
    sleep 50
    SendInput, !o
    sleep 200
    clipboard := notering

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
P6_tekstTilChf(ByRef tekst:=" ", kørselsaftale := "", styresystem := "")
{
    global s
    gemtklip := ClipboardAll
    P6_aktiver()
    if (kørselsaftale = "")
        kørselsaftale := P6_hent_k()
    if (styresystem = "")
        styresystem := P6_hent_s()
    systjek := p6_tekst_tjek_for_system(styresystem)
    if systjek = 1
        return 1
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
    {
        clipboard := gemtklip
        return
    }
    clipboard := gemtklip
    return
}

; ***
; Finder lukketid ud fra sidste stop og tid til hjemzone.
; Input tid for sidste stop, tryk enter. Input tid til hjemzone, tryk enter.
; Hvis tid for sidste stop hjemzone er tom, luk nu + 5 min
; hvis tid til hjemzone stop tom luk til udfyldte tid for sidste stop uden ændringer
; hvis tid for sidste stop og tid til hjemzone udfyldt, luk til tiden fra sidste stop til hjemzone, plus 2 min
p6_tekst_tjek_for_system(styresystem)
{
    for i,e in ["2" , "4" , "6" , "7" , "8" , "10" , "11" , "13" , "14" , "16" , "17" , "18" , "19" , "20"] 
    {
    if (styresystem = e)
        {
        MsgBox, 16 , Styresystem %styresystem% , Dette styresystem kan ikke modtage tekstbeskeder, 
        sys_tjek := 1
        return sys_tjek
     }
    }
}

P6_input_sluttid()
{
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    p6_input_sidste_slut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,42)
    KeyWait, Ctrl,
    KeyWait, Shift,
    EnvAdd, nu_plus_5, 5, minutes
    FormatTime, nu_plus_5, %nu_plus_5%, HHmm
    FormatTime, dato, YYYYMMDDHH24MISS, ddMM
    if (p6_input_sidste_slut_ops = "1")
    {
        luk := []
        luk.InsertAt(4, dato)
        sleep 100
        InputBox, sidste_stop, Sidste stop, Tast tid for sidste stop. Enter uden noget giver luk nu. `n (4 cifre)
        if (ErrorLevel = "1")
            Return 0
        if (sidste_stop = "")
        {
            luk.1 := nu_plus_5
            return luk
        }
        luk.1 := sidste_stop
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
        if (StrLen(luk.1)!= 4)
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
        sleep 200
        InputBox, tid_til_hjemzone, Tid til hjemzone, Tid til hjemzone i minutter. Enter uden noget giver luk til tidspunktet, der blev indtastet før.
        if (ErrorLevel = "1")
            Return 0
        if (tid_til_hjemzone = "" )
        {
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
    if (p6_input_sidste_slut_ops = "0")
    {
        luk := []
        Input, sidste_stop, T10, {Enter}{escape}
        luk.InsertAt(4, dato)
        if (ErrorLevel = "EndKey:Escape")
            Return 0
        if (ErrorLevel = "Timeout")
        {MsgBox, , Timeout , Det tog for lang tid.
            return 0
        }
        if (sidste_stop = "")
        {
            luk.1 := nu_plus_5
            return luk
        }
        luk.InsertAt(1, sidste_stop)
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
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    k_aftale := p6_vl_vindue_edit()
    if (k_aftale = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 40
    if (k_aftale = 1)
    {
        MsgBox, , VL afsluttet, VL er allerede afsluttet
        SendInput, ^a
        sys_afslut_genvej()
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

p6_replaner_hent_vl()
{
    gemtklip := ClipboardAll
    ; global vl_repl
    ; global vl_repl_liste
    FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
    clipboard :=
    SendInput, ^c
    clipwait 2
    if (clipboard = "")
        {
            SendInput, ^c
            ClipWait, 1    
        }
    if (InStr(clipboard, "planlægges"))
    {
        SendInput, {enter}
        return 0
    }
    repl_besked := StrSplit(clipboard, " ")
    SendInput, {enter}
    if (repl_besked.MaxIndex() = 11)
        vl := repl_besked.6
    ; vl_repl.Push(repl_besked.6)
    if (repl_besked.MaxIndex() = 12)
        vl := repl_besked.7
    ; vl_repl.Push(repl_besked.7)/mtebk1200
    clipboard := gemtklip
    return vl

}

; del af vl_liste_laas_vl
vl_liste_laas_tjek(vl)
{
    global vl_liste_array

    for i, e in vl_liste_array
    {
        if (e[8] = "listbox5" and e.1 = vl) and if (!InStr(e.3, "låst"))
        {
            return i
        }
    }
    return 0
}
; Tjekker for VL i oversigt og sætter låst/låst op status
vl_liste_laas_vl(vl)
{
    global vl_liste_array

    laas_i := vl_liste_laas_tjek(vl)
    if (laas_i = 0)
    {
        vl_array := vlliste_laast_lav_array(vl)
        vlliste_vl_array_til_liste(vl_array)
    }
    Else
    {
        vl_array := vlliste_laast_op_lav_array(vl,laas_i)
    }
    ; vlListe_dan_liste("listbox5")
    return laas_i
}
; sætter lås på VL, sender til oversigt, sætter notat på
p6_laas_vl()
{
    vl := p6_vl_vindue()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 90
    vl_liste_laas_tjek(vl)
    tjek := p6_vl_vindue_edit()
    if (tjek = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 90
    p6_vl_vindue_laas(vl)
    return
}
; Sender kun lås til oversigt
p6_marker_vl_laas_minimal()
{
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 50
    vl_laas := vl_liste_laas_vl(vl)
    return
}
; sender lås til oversigt og sætter notat
p6_marker_vl_laas()
{
    initialer := sys_initialer()
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 50
    vl_laas := vl_liste_laas_vl(vl)
    if (vl_laas = 0)
        P6_notat("låst" initialer " ")
    if (vl_laas != 0 )
        P6_notat("låst op" initialer " ")
    return
}
; del af p6_laas_vl, låser åbent VL-vindue
p6_vl_vindue_laas(vl)
{

    SendInput, {enter 2}{space}!v^{Up}
    sleep 50
    vl_laas := vl_liste_laas_vl(vl)
    initialer := sys_initialer()
    if (vl_laas = 0)
        SendInput, låst %initialer% {Space}
    if (vl_laas != 0)
        SendInput, låst op %initialer% {space}
    sleep 50
    SendInput, {enter}
    P6_planvindue()
    SendInput, {f5}!o
}
; konverter vl_liste_array til JSON, dump i tekst
vl_liste_array_til_json_tekst()
{
    global vl_liste_array
    global vl_liste_tekst

    vl_liste_array_json := json.Dump(vl_liste_array)
    FileDelete, %vl_liste_tekst%
    FileAppend, %vl_liste_array_json%, %vl_liste_tekst%
    return
}
; p6_liste_vl(vl_arr)
; {
;     gemtklip := ClipboardAll
;     global vl_repl
;     global vl_repl_liste
;     global vl_repl_tekst
;     vl_repl := []
;     vl_repl.Push(vl_arr.1)
;     ; vl_repl_liste := "|"
;     for k, v in vl_repl
;         vl_repl_liste .= vl_repl[k] . "|"
;     FileDelete, %vl_repl_tekst%
;     FileAppend, %vl_repl_liste%, %vl_repl_tekst%
;     clipboard := gemtklip
;     return vl_repl

; }
;; RYD OP
p6_vl_til_liste(vl_arr)
{
    gemtklip := ClipboardAll

    FileRead, vl_array_fil, vl_array.txt
    vl_array_fil.Push(vl_arr.1)
    ; vl_repl_liste := "|"
    FileDelete, vl_array.txt
    FileAppend, %vl_repl_liste%, vl_array.txt

    clipboard := gemtklip
    return vl_repl

}
vlListe_dan_liste(listbox)
{
    global vl_liste_array
    vl_liste_str := "|"

    for i,e in vl_liste_array
        for i2, e2 in vl_liste_array[i]
            if (e.8 = listbox)
            {
                {
                    if (i2 = 4) or if (i2 = 5 and e2 = "") or if (i2 = 8)
                        {}
                        else if (i2 = 5 and e2 != 0)
                        {
                            if (vl_liste_array[i][6] = "")
                                vl_liste_array[i].InsertAt(6, " (N)")
                        }
                        ; if (i2 = 5 and e2 != 0)
                        ; if (i2 = 5 and e2 = 0)
                        ; vl_liste_array.InsertAt(6, "")
                        else if (i2 = 5 and e2 = 0) or (i2 = 10)
                        {

                        }
                        else
                            vl_liste_str := vl_liste_str . e2
                }
            }
    ; vl_liste_array_json := JSON.Dump(vl_liste_array)
    ; vl_liste_array_json_read := json.load(vl_liste_array_json)
    ; FileDelete,  % vl_liste_tekst
    ; FileAppend, % vl_liste_array_json, % vl_liste_tekst

    return vl_liste_str
}

vlListe_vis_gui()
{
    listbox1 := vlListe_dan_liste("listbox1")
    listbox2 := vlListe_dan_liste("listbox2")
    listbox3 := vlListe_dan_liste("listbox3")
    listbox4 := vlListe_dan_liste("listbox4")
    listbox5 := vlListe_dan_liste("listbox5")
    listbox6 := vlListe_dan_liste("listbox6")
    GuiControl, vl_liste: , ListBox1, %listbox1%
    GuiControl, vl_liste: , ListBox2, %listbox2%
    GuiControl, vl_liste: , ListBox3, %listbox3%
    GuiControl, vl_liste: , ListBox4, %listbox4%
    GuiControl, vl_liste: , ListBox5, %listbox5%
    GuiControl, vl_liste: , ListBox6, %listbox6%
    Gui vl_liste: Show, w1372 h574, VL-liste
    sleep 80
    GuiControl Choose, Listbox1 , 0
    Return

}
vlListe_opdater_gui()
{
    listbox1 := vlListe_dan_liste("listbox1")
    listbox2 := vlListe_dan_liste("listbox2")
    listbox3 := vlListe_dan_liste("listbox3")
    listbox4 := vlListe_dan_liste("listbox4")
    listbox5 := vlListe_dan_liste("listbox5")
    listbox6 := vlListe_dan_liste("listbox6")
    GuiControl, vl_liste: , ListBox1, %listbox1%
    GuiControl, vl_liste: , ListBox2, %listbox2%
    GuiControl, vl_liste: , ListBox3, %listbox3%
    GuiControl, vl_liste: , ListBox4, %listbox4%
    GuiControl, vl_liste: , ListBox5, %listbox5%
    GuiControl, vl_liste: , ListBox6, %listbox6%
    Return

}
vlliste_replaner_lav_array(vl)
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd'. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", repl. "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox1"

    return vl_liste
}
vlliste_kvittering_lav_array(vl := "")
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", kvittering "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox6"

    return vl_liste
}
vlliste_wakeup_lav_array(vl := "")
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd'. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", WakeUp "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox2"

    return vl_liste
}
vlliste_priv_lav_array(vl)
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd'. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", priv. OBS "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox3"

    return vl_liste
}
vlliste_listet_lav_array(vl := "")
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd'. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", huskeliste "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox4"

    return vl_liste
}
vlliste_laast_lav_array(vl := "")
{
    vl_liste := []

    FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm 'd'. dd
    FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

    vl_liste[1] := vl
    vl_liste[2] := ", låst "
    vl_liste[3] := vl_replaner_tidspunkt_vis
    vl_liste[4] := vl_replaner_tidspunkt_intern
    vl_liste[5] := note
    vl_liste[6] :=
    vl_liste[11] := "|"
    vl_liste[8] := "listbox5"

    return vl_liste
}
vlliste_laast_op_lav_array(vl, laas_i)
{
    global vl_liste_array

    FormatTime, laas_op_tid, YYYYMMDDHH24MISS, HH:mm
    vl_liste_array[laas_i][3] := vl_liste_array[laas_i][3] . ", låst op " laas_op_tid
    return vl_liste
}

; vlliste_repl_vl_til_liste()
; {
;     global vl_liste_array

;     if (vl[1] = 0)
;         return
;     vl_liste_array.Push(vl)

;     return
; }
vlliste_vl_array_til_liste(vl_array)
{
    global vl_liste_array

    if (vl_array[1] = 0)
        return
    vl_liste_array.Push(vl_array)
    vl_liste_array_til_json_tekst()
    return
}

vlListe_note(note_reminder, note_tid, note_note)
{
     global vl_liste_array
    global tid
    global vl
    global valg
    global listbox

   if (note_reminder = 1 and StrLen(note_tid) != 4)
        {
            MsgBox, 16 , Fejl i indtastet tidspunkt, Der skal bruges fire tal, i formatet TTMM (f. eks. 1434).
            gui note: Show
            return 
        }
    note_tid_tjek := A_YYYY A_MM A_DD note_tid
        if note_tid_tjek is not Time
        {
            MsgBox, 16 , Fejl i indtastning af tidspunkt , Det indtastede er ikke et klokkeslæt.,
            gui note: show
            return
        }
    for i,e in vl_liste_array
    if (vl_liste_array[i][8] = listbox and vl_liste_array[i][1] = valg and SubStr(vl_liste_array[i][3], 1, 5) = tid) 
        {
            if (note_reminder = 1)
                {
                    vl_liste_array[i][10] := note_tid
                    vl_liste_array[i][7] := " (R)"
                }
          if (note_reminder = 0)
                {
                    vl_liste_array[i][10] := ""
                    vl_liste_array[i][7] := ""
                }
            vl_liste_array[i][6] := " (N)"
            vl_liste_array[i][5] := note_note
            if (note_note = "")
                vl_liste_array[i][6] := ""
            gui note: hide
            vl_liste_array_til_json_tekst()
            P6_aktiver()
            return
        }
P6_aktiver()
return
}

vlliste_note_reminder(note_reminder, note_tid)
{
    return
}



vlliste_vis_note_fra_planbillede()
{
    global vl_liste_array
    global tid
    global vl
    global valg
    global listbox

    P6_aktiver()
    vl := P6_hent_vl()
    if (vl = 0)
        return
    valg := vl
    listbox := "listbox4"
    fundet := 0
    for i,e in vl_liste_array
        {
            if (vl_liste_array[i][1] = vl and vl_liste_array[i][8] = listbox)
                {
                    fundet := 1
                }
        }
    if (fundet = 0)
        {
    FormatTime, vl_tid , YYYYMMDDHH24MISS, HH:mm
    vl_array := vlliste_listet_lav_array(vl)
    vlliste_vl_array_til_liste(vl_array)
        }
for i,e in vl_liste_array
    {
        if (vl_liste_array[i][1] = vl and vl_liste_array[i][8] = listbox)
            {
                note_note := vl_liste_array[i][5]  
                tid := SubStr(vl_liste_array[i][3], 1, 5)
                if (vl_liste_array[i][10] = "")
                    {
                        note_reminder = 0
                        GuiControl, note:, edit2, 
                    }
                if (vl_liste_array[i][10] != "")
                    {
                        note_reminder = 1
                        GuiControl, note: , note_reminder, 1
                        GuiControl, note:, edit2, % vl_liste_array[i][10]
                    }
            }
    }
for i,e in vl_liste_array
    {
        if (vl_liste_array[i][1] = valg and vl_liste_array[i][8] = listbox and SubStr(vl_liste_array[i][3], 1, 5) = tid)
            {
                note_note := vl_liste_array[i][5]
                break
            }
    }
GuiControl, note:, note_note, %note_note%
        GuiControl, note:, Edit2, 
        GuiControl, note:, note_reminder, 0 
gui note: show, , Note VL %valg% til huskeliste
ControlFocus, Edit1
}

;; Telenor


;; Trio
; ***
; Sæt kopieret tlf i Trio
Trio_opkald(ByRef telefon)
{
    ifWinNotExist, ahk_class Addressbook
    {
        ; ControlClick, x368 y68, ahk_class Agent Main GUI , , ,, ,, ; Main vindue
        ControlClick, x365 y18, Trio Agent, , ,, ,, ; Skrivebordsværkstøjsline
        sleep 100
    }
    trio_pause()
    sleep 100
    SendInput, {CtrlUp}{AltUp}
    if (telefon = "")
    {
        MsgBox, , , Der er ikke lavet en markering af telefonnummer
        trio_klar()
        return
    }
    ControlGetText, tlf_test, Edit2, Trio Attendant
    sleep 100
    loop_test := 0
    controlsend, Edit2, ^a{delete} ,ahk_class Addressbook
    sleep 100
    ControlGetText, tlf_test, Edit2, Trio Attendant
    while (tlf_test != "")
    {
        controlsend, Edit2, ^a{delete} ,ahk_class Addressbook
        sleep 100
        ControlGetText, tlf_test, Edit2, Trio Attendant
        if (loop_test > 10)
        {
            MsgBox, 16, Fejl, Der er sket en fejl - Prøv igen
            return 0
        }
    }
    sleep 80
    controlsend, Edit2, %telefon%, ahk_class Addressbook
    sleep 100
    ControlGetText, kobl_test, Button1, Trio Attendant
    GuiControl, trio_genvej:text, Button1, Ringer op til %telefon%
    if (kobl_test = "Koble")
    {
        controlsend, , {ShiftDown}{enter}{ShiftUp}, ahk_class Addressbook
        return
    }
    Else
    {
        controlsend, , {enter}, ahk_class Addressbook
        Return
    }
}

; ***
; Læg på i Trio
Trio_afslutopkald()
{
    ; ControlFocus, Button1, Addressbook
    ; ControlGetText, opkaldsstatus, Button1, Trio Attendant
    ; sleep 200
    ; MsgBox, , , % opkaldsstatus
    ControlSend, , {NumpadSub}, ahk_class Agent Main GUI
    ; WinActivate, ahk_class AccessBar
    ; winwaitactive, ahk_class AccessBar
    ; sleep 40
    ; SendInput, {NumpadSub}

    return
}
Trio_linie1()
{
    ; ControlFocus, Button1, Addressbook
    ; ControlGetText, opkaldsstatus, Button1, Trio Attendant
    ; sleep 200
    ; MsgBox, , , % opkaldsstatus
    ControlSend, , {F6}, ahk_class Agent Main GUI
    ; WinActivate, ahk_class AccessBar
    ; winwaitactive, ahk_class AccessBar
    ; sleep 40
    ; SendInput, {NumpadSub}

    return
}
Trio_linie2()
{
    ; ControlFocus, Button1, Addressbook
    ; ControlGetText, opkaldsstatus, Button1, Trio Attendant
    ; sleep 200
    ; MsgBox, , , % opkaldsstatus
    ControlSend, , {F7}, ahk_class Agent Main GUI
    ; WinActivate, ahk_class AccessBar
    ; winwaitactive, ahk_class AccessBar
    ; sleep 40
    ; SendInput, {NumpadSub}

    return
}

; **
; Trio hop til efterbehandling
trio_efterbehandling()
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Fil, Rolle, 9&
    ; WinActivate, ahk_class Agent Main GUI
    ; winwaitactive, ahk_class Agent Main GUI
    ; sleep 40
    ; SendInput, !f
    ; sleep 40
    ; SendInput, o
    ; sleep 40
    ; SendInput, 8
    ; WinActivate, PLANET
    ; winwaitactive, PLANET
    Return
}

; **
; Trio hop til midt uden overløb
trio_udenov()
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Fil, Rolle, 4&
    ; WinActivate, ahk_class Agent Main GUI
    ; winwaitactive, ahk_class Agent Main GUI
    ; sleep 40
    ; SendInput, !f
    ; sleep 40
    ; SendInput, o
    ; sleep 40
    ; SendInput, 3
    ; sleep 100
    ; SendInput, {F4}
    ; WinActivate, PLANET
    ; winwaitactive, PLANET
    Return
}

; **
; Trio hop til alarm
trio_alarm()
{
    WinMenuSelectItem, Trio Attendant, , Fil, Rolle, 8&
    ; WinActivate, ahk_class Agent Main GUI
    ; winwaitactive, ahk_class Agent Main GUI
    ; sleep 40
    ; SendInput, !f
    ; sleep 40
    ; SendInput, o
    ; sleep 40
    ; SendInput, 7
    ; WinActivate, PLANET
    ; winwaitactive, PLANET
    Return
}

; **
; Trio hop til pause
trio_pause()
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Fil, Pause
    ; WinActivate, ahk_class AccessBar
    ; winwaitactive, ahk_class AccessBar
    ; sleep 100
    ; SendInput, {F3}
    ; WinActivate, PLANET
    ; winwaitactive, PLANET
    Return
}

; **
; Trio hop til klar
trio_klar()
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Fil, Klar
    ; WinActivate, ahk_class AccessBar
    ; winwaitactive, ahk_class AccessBar
    ; Sleep 100
    ; SendInput, {F4}
    ; WinActivate, PLANET
    ; winwaitactive, PLANET
    Return
}

; **
; Trio hop til frokost
trio_frokost()
{
    WinMenuSelectItem, ahk_class Agent Main GUI, , Fil, Rolle, 10&
    ;winwaitactive, ahk_class Agent Main GUI
    ;sleep 40
    ;SendInput, !f
    ;sleep 40
    ;SendInput, o
    ;sleep 40
    ;SendInput, 9
    ;WinActivate, PLANET
    ;winwaitactive, PLANET
    Return
}

; Trio skift mellem pause og klar

trio_pauseklar()
{
    trio_pause()
    sleep 900
    trio_klar()

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
    sys_genvej_start(36)
    If (WinExist("FlexDanmark FlexFinder"))
    {
        k_aftale := P6_hent_k()
        k_aftale := SubStr("000" . k_aftale, -3) ; indsætter nuller og tager sidste fire cifre i strengen (kun i spil når mindre end fire cifre ind).
        sty_sys := P6_hent_s()
        sty_sys := SubStr("000" . sty_sys, -3) ; indsætter nuller og tager sidste fire cifre i strengen (kun i spil når mindre end fire cifre ind).
        opslag := k_aftale "_" sty_sys
        ; MsgBox, , er 4 , % k_aFtale
        sleep 200
        WinActivate, FlexDanmark FlexFinder
        winwaitactive, FlexDanmark FlexFinder
        sleep 40
        SendInput, {Home}
        sleep 400
        SendInput, {PgUp}
        sleep 200
        WinGetPos, W_X, W_Y, , , FlexDanmark FlexFinder, , ,
        if(W_X = "1920" or W_X = "-1920")
        {
            PixelSearch, Px, Py, 1097, 74, 1202, 123, 0x5B6CF2, 0, Fast ; Virker ikke i fuld skærm. ControlClick i stedet?
            sleep 200
            click %Px% %Py%
            sleep 200
            ControlClick, x322 y100, FlexDanmark FlexFinder
            sleep 40
            SendInput, +{tab}{up}{tab}
            sleep 200
            SendInput, %opslag%
            sleep 500
            SendInput, {enter}
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
            SendInput, %opslag%
            sleep 500
            SendInput, {enter}
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
; Outlook
; ***
; Åbn ny mail i outlook. Kræver nymail.lnk i samme mappe som script. Kolonne 37
Outlook_nymail()
{
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

excel_p6_kundenummer()
{
    clipboard :=
    SendInput, ^c
    clipwait 2
    kundenr := StrSplit(clipboard, "`r`n")
    sleep 500
    SendInput, {Down}
    sleep 100
    P6_aktiver()
    P6_rejsesogvindue()
    sleep 50
    SendInput, ^t
    sleep 50
    SendInput, +{tab}
    sleep 150
    SendInput, % kundenr[1]
    sleep 100
    SendInput, ^r

    return
}
excel_p6_id()
{
    clipboard :=
    SendInput, ^c
    clipwait 2
    p6_id := StrSplit(clipboard, " ", "`n")
    sleep 500
    SendInput, {Down}
    sleep 100
    P6_aktiver()
    P6_rejsesogvindue()
    sleep 50
    SendInput, ^t
    sleep 50
    SendInput, !n{Down}{Tab}
    sleep 150
    SendInput, % p6_id[1]
    sleep 100
    SendInput, ^r

    return
}
excel_p6_cpr()
{
    clipboard :=
    SendInput, ^c
    clipwait 2
    cpr := StrSplit(clipboard, " ", "`n")
    sleep 500
    SendInput, {Down}
    sleep 100
    P6_aktiver()
    P6_rejsesogvindue()
    sleep 50
    SendInput, ^t
    sleep 150
    SendInput, % cpr[1]
    sleep 100
    SendInput, ^r

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
    genvej_mod := []
    genvej := RegExReplace(bruger_genvej[kolonne], "[\^!\^\+\#]")
    
    genvej_mod_midl := StrSplit(bruger_genvej[kolonne])
    genvej_mod.Push(genvej_mod_midl[1], genvej_mod_midl[2])
    genvej_mod.InsertAt(3, genvej)
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
    genvej_mod3 := genvej_mod.3
    KeyWait, %genvej_mod1%,
    if (genvej_mod2 = "shift" or genvej_mod2 = "alt" or genvej_mod2 = "control" or genvej_mod2 = "lwin")
        keywait, %genvej_mod2%
    if (genvej_mod3 != "")
        keywait, %genvej_mod3%
       
    return
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
    sys_genvej_start(18)
    vis_sygehus_1()
return

l_p6_central_ring_op:
    sys_genvej_start(19)
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
    tjek := Trio_opkald(telefon)
    if (tjek = 0)
    {
        sys_afslut_genvej()
        return
    }
    WinActivate, PLANET, , ,
    sleep 3000
    trio_klar()
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
    sys_afslut_genvej()
return
;; PLANET
l_p6_hastighed:
    P6_hastighed()
    sys_afslut_genvej()
return
; skriv/fjern initialer. Kolonne 5
l_p6_initialer: ;; Initialer til/fra
    sys_genvej_start(5)
    P6_initialer()
    sys_afslut_genvej()
Return
l_p6_initialer_slet_eget:
    sys_genvej_start(55)
    P6_initialer_slet_eget()
    sys_afslut_genvej()
Return
l_p6_initialer_skift_eget:
    sys_genvej_start(59)
    P6_initialer_skift_eget()
    sys_afslut_genvej()
Return
; skriv initialer, fortsæt notat. Kolonne 6
l_p6_initialer_skriv: ; skriv initialer og forsæt notering.
    sys_genvej_start(6)
    P6_initialer_skriv()
    sys_afslut_genvej()
return
; gentag notat. Kolonne 57
l_p6_notat_igen: ; skriv initialer og forsæt notering.
    sys_genvej_start(57)
    P6_notat_igen()
    sys_afslut_genvej()
return

l_p6_vis_k_aftale: ;Vis kørselsaftale for aktivt vognløb
    P6_vis_k()
    sys_afslut_genvej()
Return

l_p6_ret_vl_tlf: ; +F3 - ret vl-tlf til triopkald
    faste_dage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
    uge_dage := ["faste mandage", "faste tirsdage", "faste onsdage", "faste torsdage", "faste fredage", "faste lørdage", "faste søndage"]

    sys_genvej_start(8)
    sleep 100
    klip := clipboard
    sleep 200
    telefon := Trio_hent_tlf()

    WinActivate, PLANET
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
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
        sys_afslut_genvej()
        return
    }
    sleep 100
    MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl%?,
    IfMsgBox, no
    {
        sys_afslut_genvej()
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
                    sys_afslut_genvej()
                    return
                }
                sleep 200
                P6_ret_tlf_vl_efterfølgende(telefon)
                sleep 200
                continue
            }
            IfMsgBox, no
            {
                sys_afslut_genvej()
                return
            }
        }

    }
    ; }
    sys_afslut_genvej()
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
        sys_afslut_genvej()
        return
    }
    else
        sleep s * 40
    P6_udfyld_k_og_s(vl)
    sys_afslut_genvej()
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
            sys_afslut_genvej()
            return
        }
        if (telefon = "78410222")
        {

            P6_rejsesogvindue()
            sleep s * 40
            SendInput, ^t
            sys_afslut_genvej()
            return
        }
        Else
        {
            WinActivate, PLANET
            P6_rejsesog_tlf(telefon)
            sys_afslut_genvej()
            return
        }
    }
return

; gå i vl 
l_p6_vaelg_vl:
    p6_vaelg_vl()
    sys_afslut_genvej()
return
l_p6_vaelg_vl_liste:
    p6_vaelg_vl_liste()
    sys_afslut_genvej()
return
;træk tlf fra aktiv planbillede, ring op i Trio. Col 11
l_p6_vl_ring_op:
sys_genvej_start(11)
sleep s * 100
    vl_tlf := P6_hent_vl_tlf()
    if (vl_tlf = 0)
    {
        sys_afslut_genvej()
        return
    }
    if (vl_tlf = "")
    {
        MsgBox, 4, Prøv igen?, Tlf-nr ikke opfanget. Prøv igen?
        IfMsgBox, yes
            Goto, l_p6_vl_ring_op
        IfMsgBox, no
            sys_afslut_genvej()
        return
    }
    sleep 200
    tjek := Trio_opkald(vl_tlf)
    if (tjek = 0)
    {
        sys_afslut_genvej()
        return
    }
    ; Clipboard = %gemtklip%
    ; gemtklip :=
    sleep 400
    WinActivate, PLANET
    P6_Planvindue()
    sleep 3000
    trio_klar()
    sys_afslut_genvej()
return

; ***

; ^+F5 col 12
l_p6_vm_ring_op: ; træk vm-tlf fra aktivt planbillede, ring op i Trio
    sys_genvej_start(12)
    P6_planvindue()
    sleep s * 100
    vm_tlf := P6_hent_vm_tlf()
    if (vm_tlf = "fejl")
    {

        sys_afslut_genvej()
        return
    }
    sleep 500
    tjek := Trio_opkald(vm_tlf)
    if (tjek := 0)
    {
        sys_afslut_genvej()
        return
    }
    sleep 800
    WinActivate, PLANET
    sleep 3000
    trio_klar()
    sys_afslut_genvej()
Return

; P6 - ring op til kunde markeret i Vl (kræver tlf opsat på kundetilladelse)
l_p6_ring_til_kunde:
    p6_hent_kunde_tlf(telefon)
    sleep s * 200
    if (SubStr(telefon, 1, 3) = "888")
    {
        MsgBox, , Telefon ikke tilknyttet, Kunden har ikke telefon tilknyttet.
        sys_afslut_genvej()
        return
    }
    Else
    {
        tjek := Trio_opkald(telefon)
        if (tjek := 0)
        {
            trio_klar()
            sys_afslut_genvej()
            return
        }
        sleep 3000
        trio_klar()
        sys_afslut_genvej()
        return
    }
return
l_p6_laas_vl:
    sys_genvej_beskrivelse(62)
    genvej_mod := sys_genvej_til_ahk_tast(62)
    sys_genvej_keywait(genvej_mod)

    p6_marker_vl_laas()
return
; #F5, col 13
l_p6_vl_luk:
    sys_genvej_start(13)
    gemtklip := ClipboardAll

    tid := P6_input_sluttid()
    if !tid
    {
        sys_afslut_genvej()
        return
    }
    p6_vl_luk(tid)
    sleep 100
    P6_planvindue()
    sleep 200
    SendInput, {F5}
    sys_afslut_genvej()

    clipboard := gemtklip
    gemtklip :=
return

l_p6_udregn_minut:
    sys_genvej_start(17)
    tid := P6_udregn_minut()
    tid_tekst := tid.1
    if (tid = "fejl")
    {
        sys_afslut_genvej()
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
l_p6_cpr_til_bestillingsvindue:
    {
        sys_genvej_start(58)
        p6_cpr_til_bestillingsvindue()
        return
    }
ok:
    {
        gui, cancel
        sys_afslut_genvej()
        return
    }
udklip:
    {
        Clipboard := tid.2
        gui, cancel
        sys_afslut_genvej()
        return
    }
plustidGuiEscape:
plustidGuiClose:
    gui, cancel
    sys_afslut_genvej()
return
l_p6_billede_gui:
{
    sys_genvej_start(69)
    Gui p6_billede: Show, w218 h303, Alt F9   
    sys_afslut_genvej()
    return
}
l_p6_tjek_andre_rejser:
    {
        sys_genvej_start(66)
        p6_tjek_andre_rejser()
        sys_afslut_genvej()
        return
    }



l_p6_alarmer:
    sys_genvej_start(14)
    P6_alarmer()
    sys_afslut_genvej()
return

l_p6_udraabsalarmer:
    sys_genvej_start(15)
    P6_udraabsalarmer()
    sys_afslut_genvej()
return

l_p6_tag_alarm:
    sys_genvej_start(56)
    p6_tag_alarm()
    sys_afslut_genvej()
Return

l_p6_tag_alarm_vl_box:
    sys_genvej_beskrivelse(56)
    p6_tag_alarm_vl_box()
    sys_afslut_genvej()
Return

; Replaner og gem i liste, kolonne 49
l_p6_replaner_liste_vl:
    sys_genvej_start(49)
    vl := p6_replaner_hent_vl()
    if (vl = 0)
        return
    vl_array := vlliste_replaner_lav_array(vl)
    vlliste_vl_array_til_liste(vl_array)
    sys_afslut_genvej()
return
; Replaner og gå til VL, kolonne 60
l_p6_replaner_opslag_vl:
    sys_genvej_start(60)
    vl := p6_replaner_hent_vl()
    if (vl = 0)
        return
    sleep 200
    p6_vaelg_vl(vl)
    sys_afslut_genvej()
return

; Gem aktiv vl på liste, kolonne 50
l_p6_liste_vl:
    sys_genvej_start(50)
    FormatTime, vl_tid , YYYYMMDDHH24MISS, HH:mm
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    vl_array := vlliste_listet_lav_array(vl)
    vlliste_vl_array_til_liste(vl_array)
    sys_afslut_genvej()
return

; vist VL-liste, kolonne 51
l_p6_vis_liste_vl:

    sys_genvej_start(51)
    vlListe_vis_gui()
    sys_afslut_genvej()

return

l_p6_vis_liste_fra_planbillede:
{
    sys_genvej_start(67)
    vlliste_vis_note_fra_planbillede()
    sys_afslut_genvej()
    return
}


l_p6_tekst_til_chf: ; Send tekst til aktive vognløb
    sys_genvej_start(20)
    FormatTime, tid, ,HHmm
    initialer = /mt%A_userName%%tid%
    initialer_udentid =/mt%A_userName%
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    bruger := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 2)
    ; ctrl_s := chr(19)
    gemtklip := ClipboardAll

    ; KeyWait Alt
    ; keywait Ctrl
    Input valgt, L1 T5 C, {esc},
    if (ErrorLevel = "EndKey:Escape")
        {
        sys_afslut_genvej()
        return
        }

    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    kørselsaftale := P6_hent_k()
    styresystem := P6_hent_s()
    clipboard := gemtklip
    ; loop_test := 0
    ; while (vl = "")
    ;    {
    ;     sleep 400
    ;     vl := P6_hent_vl()
    ;     loop_test += 1
    ;     if (loop_test > 10)
    ;         {
    ;             MsgBox, 16 , Fejl, Der er sket en fejl - prøv igen,
    ;             return
    ;         }
    ;    }
    if (valgt = "t")
    {
        P6_tekstTilChf( , kørselsaftale, styresystem) ; tager tekst ("eksempel") som parameter (accepterer variabel)
        sys_afslut_genvej()
        return
    }
    if (valgt = "f")
        {
    sys_tjek := p6_tekst_tjek_for_system(styresystem)
    if (sys_tjek = 1)
        {      
        sys_afslut_genvej()
        return
        }
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
                sys_afslut_genvej()
                return
            }
        f_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            ; MsgBox, , , % tekst,
            gui, cancel
            P6_tekstTilChf("Jeg kan ikke ringe dig op. Jeg har meldt st. " f_stop "`, " . k_navn "`, forgæves og sendt st. " s_stop "`, " k_navn2 ", i stedet - Mvh. Midttrafik", kørselsaftale, styresystem)
            sleep 500
            MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
            IfMsgBox, Yes
            {
                SendInput, ^s
                ; KeyWait, Ctrl
                sleep 1000
                SendInput, {enter}
                P6_notat("Ingen kontakt til chf. St. " f_stop " forgæves`, " s_stop " og tekst sendt til chf." initialer)
                sys_afslut_genvej()
                return
            }
            IfMsgBox, No
            {
                sleep 200
                MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
                gui, cancel
            }
            sys_afslut_genvej()
        return
    }
}
if ( valgt == "k")
    {
    systjek := p6_tekst_tjek_for_system(styresystem)
    if (systjek = 1)
        {
        sys_afslut_genvej()    
        return
        }
        InputBox, stop, St. nummer, Hvilket stop?
        if ErrorLevel
            {
                sys_afslut_genvej()   
                Return
            }
        InputBox, tid, FlexFinder ankomst, Hvornår faktisk ankommet? 4 cifre
        if ErrorLevel
            {
                sys_afslut_genvej()
                Return
            }
        P6_tekstTilChf("Er der glemt at bede om ny tur v. ankomst? Der skal altid trykkes for næste køreordre ved ankomst på en adresse, uanset om det er en afhentning eller en aflevering. Mvh. Midttrafik", kørselsaftale, styresystem)
            sleep 500
            MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
            IfMsgBox, Yes
            {
                SendInput, ^s
                ; KeyWait, Ctrl
                sleep 1000
                SendInput, {enter}
                P6_notat("St. " stop " ikke kvitteret, ankommet " tid " jf. FF. Tekst sendt til chf, bed om næste køreordre" initialer)
                sys_afslut_genvej()
                return
            }
            IfMsgBox, No
            {
                sleep 200
                MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
                gui, cancel
                sys_afslut_genvej()
                return
            }
            sys_afslut_genvej()

            return
 
    }      
    if (valgt == "K")
        {
    systjek := p6_tekst_tjek_for_system(styresystem)
    if (systjek = 1)
        {
        sys_afslut_genvej()    
        return
        }
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
                sys_afslut_genvej()
                return
            }
        k_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            GuiControlGet, k_tid, , ,
            gui, cancel
            P6_tekstTilChf("Husk at bede om ny tur ved ankomst. Jeg har bekræftet ankomst ved st. " f_stop "`, " . k_navn "`, og sendt st. " s_stop "`, " k_navn2 " - Mvh. Midttrafik", kørselsaftale, styresystem)
            sleep 500
            MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
            IfMsgBox, Yes
            {
                SendInput, ^s
                sleep 1000
                SendInput, {enter}
                FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
                vl_array := vlliste_kvittering_lav_array(vl)
                vlliste_vl_array_til_liste(vl_array)
                if (k_tid != "Oprindelig kvittering")
                {
                    P6_notat("St. " f_stop " ikke kvitteret ved ankomst`, st. " s_stop " og tekst sendt til chf. Oprindeligt kvitt. tid " k_tid initialer " ")
                    return
                }
                else
                    P6_notat("St. " f_stop " ikke kvitteret ved ankomst`, st. " s_stop " og tekst sendt til chf. " initialer " ")
                sys_afslut_genvej()
                return
            }
            IfMsgBox, No
            {
                sleep 200
                MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
                gui, cancel
            }
            sys_afslut_genvej()
        return
    }
}
    if (valgt == "p")
    {

        sys_tjek := P6_tekstTilChf("Er der blevet glemt at kvittere for privatrejsen? Mvh. Midttrafik", kørselsaftale ,styresystem)
        sleep 500
        if (sys_tjek = 1)
            {
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_priv_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            P6_notat("Priv. ikke kvitteret" initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return
            }
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_priv_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            P6_notat("Priv. ikke kvitteret, tekst sendt til chf" initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        sys_afslut_genvej()
        return
    }
    if (valgt == "P")
    {
    systjek := p6_tekst_tjek_for_system(styresystem)
    if (systjek = 1)
        {
        P6_notat("Priv. ikke kvitteret, ingen kontakt til chf. Låst" initialer " ")
        sys_afslut_genvej()
        return
        }
    {

        P6_tekstTilChf("Jeg kan ikke ringe dig op, din privatrejse er ikke kvitteret. Vognløbet er låst, ring til driften, hvis du er ude at køre.", kørselsaftale , styresystem)
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            P6_notat("Priv. ikke kvitteret, ingen kontakt til chf. Låst, tekst sendt om VL-lås" initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        sys_afslut_genvej()
        return
    }
    }
    if (valgt == "w")
    {
        sys_tjek := P6_tekstTilChf("Der er ikke bedt om vognløb start. Huske at bede om første køreordre ved opstart, uanset om der ligger ture eller ej. Mvh. Midttrafik", kørselsaftale, styresystem)
        sleep 500
        if (sys_tjek = 1)
            {
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_wakeup_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            P6_notat(initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return 
            }
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør?,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            SendInput, {enter}
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_wakeup_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            P6_notat("WakeUp sendt" initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            sys_afslut_genvej()
            return
        }
    }
    if (valgt == "W")
    {
        sys_tjek := P6_tekstTilChf("Jeg kan ikke ringe dig op, der er ikke trykket for første køreordre. Vognløbet er nu låst, ring til driften, hvis du er ude at køre.", kørselsaftale , styresystem)
        if (sys_tjek = 1)
           {
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_laast_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            P6_notat("Ingen kontakt til chf, VL låst" initialer " ")
            gui, cancel
            sys_afslut_genvej()
           } 
        sleep 500
        MsgBox, 4, Send til chauffør?, Send tekst til chauffør? Husk at låse VL,
        IfMsgBox, Yes
        {
            sleep 200
            SendInput, ^s
            sleep 1000
            FormatTime, tid, YYYYMMDDHH24MISS, HH:mm
            vl_array := vlliste_laast_lav_array(vl)
            vlliste_vl_array_til_liste(vl_array)
            SendInput, {enter}
            P6_notat("Ingen kontakt til chf, tekst sendt, VL låst" initialer " ")
            gui, cancel
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            sys_afslut_genvej()
            return
        }
    }
    if (valgt == "r")
    {
    systjek := p6_tekst_tjek_for_system(styresystem)
    if (systjek = 1)
        {      
        MsgBox, 16 , Styresystem %styresystem% , Dette styresystem kan ikke modtage tekstbeskeder, 
        P6_notat("Ingen kontakt til chf" initialer " ")
        sys_afslut_genvej()
            return
        }
        tlf := P6_hent_vl_tlf()
        P6_tekstTilChf("Jeg kan ikke ringe dig op på telefonnummer " tlf ". Ring til driften, 70112210. Mvh Midttrafik.", kørselsaftale, styresystem)
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
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
            sys_afslut_genvej()
            return
        }
    }
    if (valgt = "a")
    {
        sys_tjek := P6_tekstTilChf("Jeg kan ikke ringe dig op. Tryk for opkald igen, hvis du stadig gerne vil ringes op. Mvh. Midttrafik", kørselsaftale, styresystem)
        if (sys_tjek = 1)
            {
            P6_notat("Tal forgæves" initialer " ")
            sys_afslut_genvej()
            return
            }
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
            sys_afslut_genvej()
            return
        }
        IfMsgBox, No
        {
            sleep 200
            MsgBox, , Ikke sendt, Tekst er ikke blevet sendt,
            gui, cancel
        }
        sys_afslut_genvej()
        return
    }
    if (valgt = "n")
    {
        ; Jeg kan ikke ringe dig op, jeg har sendt dig en ny tur
    }
    sys_afslut_genvej()
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
l_trio_linie1: ;Trio efterbehandling
    trio_linie1()
Return
l_trio_linie2: ;Trio efterbehandling
    trio_linie2()
Return

l_trio_alarm: ;Trio alarm bruger.9
    trio_alarm()
Return

l_trio_frokost: ;Trio frokostr. bruger.10
    trio_frokost()
Return

l_triokald_til_udklip: ; trækker indkommende kald til udklip, ringer ikke op.
    sys_genvej_start(29)
    clipboard := Trio_hent_tlf()
    sys_afslut_genvej()
Return

; Telenor accepter indgående kald, søg planet
l_trio_P6_opslag: ; brug label ist. for hotkey, defineret ovenfor. Bruger.4
    sys_genvej_start(4)
    if (!WinExist("--- ahk_exe Miralix OfficeClient.exe") and !WinExist("+ ahk_exe Miralix OfficeClient.exe"))
        {
            Trio_afslutopkald()
            sys_afslut_genvej()
            return
        }
    if (WinExist("+4570112210 ahk_exe Miralix OfficeClient.exe"))
        {
            SendInput, % bruger_genvej[68] ; Misser den af og til?
            sys_afslut_genvej()
            return
        }
    
    if (WinExist("--- ahk_exe Miralix OfficeClient.exe") OR WinExist("+ ahk_exe Miralix OfficeClient.exe"))
        {
    ControlGetText, koble_test, Button1, Trio Attendant
    SendInput, % bruger_genvej[68] ; Misser den af og til?
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
        sys_afslut_genvej()
        return
    }
    vl := P6_hent_vl_fra_tlf(telefon)
    if vl
    {
        sleep 200
        P6_udfyld_k_og_s(vl)
        sys_afslut_genvej()
        Return
    }
    if (telefon = "78410222" OR telefon ="78410224") ; mangler yderligere?
    {
        ; MsgBox, ,CPR, CPR, 1
        sleep 200
        P6_rejsesogvindue()
        SendInput, ^t
        sys_afslut_genvej()
        return
    }
    Else
    {
        sleep 200
        P6_rejsesogvindue(telefon)
        sys_afslut_genvej()
        return
    }
    return
        }        

; Opkald på markeret tekst. Kolonne 28
l_trio_opkald_markeret: ; Kald det markerede nummer i trio, global. Bruger.12
    sys_genvej_start(28)
    SendInput, {click}
    sleep 100
    SendInput, {Click}
    sleep 200
    clipboard := ""
    SendInput, ^c
    ClipWait, 1.3, 0
    if (clipboard = "")
        {
    SendInput, {click}
    sleep 100
    SendInput, {Click}
    sleep 200
    clipboard := ""
    SendInput, ^c
    ClipWait, 1.3, 0
        }
    telefon := clipboard
    telefon := RegExReplace(telefon, "\D")
    GuiControl, trio_genvej:text, Button1, Ringer op til %telefon%
    sleep 300
    tjek := Trio_opkald(telefon)
    if (tjek := 0)
    {
        trio_klar()
        sys_afslut_genvej()
        return
    }
    sleep 3100 ; for at genvejsbeskrivelsen bliver der - et problem?
    trio_klar()
    sys_afslut_genvej()
Return

; Minus på numpad afslutter Trioopkald global (Skal der tilbage til P6?)
l_trio_afslut_opkald:
l_trio_afslut_opkaldB:
    sys_genvej_start(30)
    Trio_afslutopkald()
    sys_afslut_genvej()
Return

;; Flexfinder
l_flexf_fra_p6:
    sys_genvej_keywait(36)
    Flexfinder_opslag()
    sys_afslut_genvej()
Return
; slå VL op i FF. Kolonne 36
l_flexf_til_p6:
    sys_genvej_start(35)
    sleep 200
    vl :=Flexfinder_til_p6()
    if !vl
    {
        sys_afslut_genvej()
        return
    }
    Else
    {
        P6_aktiver()
        sleep s * 200
        P6_udfyld_k_og_s(vl)
        sleep 400 ; skal optimeres
        WinActivate, FlexDanmark FlexFinder, , ,
        sys_afslut_genvej()
        Return
    }

;; Outlook
l_outlook_ny_mail: ; opretter ny mail. Bruger.16
    sys_genvej_start(37)
    Outlook_nymail()
    sys_afslut_genvej()
Return

;; Excel til vl.
l_excel_vl_til_P6_A:
l_excel_vl_til_P6_B:
    sys_genvej_start(40)
    vl := Excel_vl_til_udklip()
    sleep 400
    SendInput, {Esc}
    Excel_udklip_til_p6(vl)
return

l_excel_mange_ture:
    sys_genvej_start(52)
    excel_p6_kundenummer()
    sys_afslut_genvej()
return

l_excel_p6_id:
    sys_genvej_start(53)
    excel_p6_id()
    sys_afslut_genvej()
return

l_excel_p6_cpr:
    sys_genvej_start(54)
    excel_p6_cpr()
    sys_afslut_genvej()
return

l_restartAHK: ; AHK-reload
    sys_genvej_start(46)
    Reload
    sleep 2000
Return

;; GUI-hjælp

;hjælp GUI
l_gui_hjælp:
    sys_genvej_start(47)
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
    sys_afslut_genvej()
return
l_outlook_genåben: ; tag skærmprint af P6-vindue og indsæt i ny mail til planet
    sys_genvej_start(69)
    FormatTime, dato, , d-MM-y
    ; FormatTime, tid, , HH:mm
    ; svigt := []
    gemtklip := ClipboardAll
    SendInput, ^{F10}
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    k_aft := P6_hent_k()
    sty_sys := P6_hent_s()
    k_aftale := k_aft  "_" sty_sys
    clipboard :=
    p6_vl_vindue()
    p6_vl_vindue_edit()
    clipboard := 
    SendInput, {enter}{tab}^c
    ClipWait, 1
    aabningstid := clipboard
    clipboard := 
    SendInput, {enter}!v+{up}
    sleep 200
    SendInput, ^c
    ClipWait, 1
    vl_notat := clipboard
    sleep 500
    SendInput, !{PrintScreen}
    sleep 500
    SendInput, ^a
    FileRead, gv_svigt, %A_linefile%\..\db\gv_svigt.txt
    gv_svigt := StrSplit(gv_svigt, ["`n"])
    for i, e in gv_svigt
        {
        gv_svigt[i] := SubStr(gv_svigt[i], 1 , -1)
        gv_svigt[i] := StrSplit(gv_svigt[i], "`t")
        }
    for i,e in gv_svigt
        {
            if (k_aftale = gv_svigt[i][1])
                opr_vl := gv_svigt[i][2]
        }
    gemtklip := ClipboardAll
    clipboard :=
    sleep 500
    SendInput, !{PrintScreen}
    ClipWait, 1, 
    
    if (vl = opr_vl)
        emnefelt := "VL " vl " genåbnet som VL " vl " d. " dato " kl. " aabningstid
    if (vl != opr_vl)
        emnefelt := "VL " opr_vl " genåbnet som VL " vl " d. " dato " kl. " aabningstid
    outlook_template := A_ScriptDir . "\lib\svigt_template.oft"
    svigt_template := outlook.createitemfromtemplate(outlook_template)

    udklip := ImagePutFile(clipboardall, "genåbnet.png")
    udklip_navn := SubStr(udklip, 3)
    udklip_lok := A_ScriptDir "\" udklip_navn
    
    svigt_template.attachments.add(udklip_lok)
    svigt_template.to := "planet@midttrafik.dk"
    svigt_template.subject := emnefelt
    html_tekst =
    (
    </o:shapelayout></xml><![endif]--></head><body lang=DA link="#0563C1" vlink="#954F72" style='tab-interval:65.2pt;word-wrap:break-word'><div class=WordSection1><img id="Billede_x0020_2" src="cid:%udklip_navn%"></span></p><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-fareast-language:DA'</o:p></span></p></div><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>
    )
    svigt_template.htmlbody := beskrivelse html_tekst


    svigt_template.send
    ImageDestroy(udklip)
    P6_planvindue()
    if (vl_notat = "")
        {
            MsgBox, 48, Mail sendt - husk notat på VL, Mail om genåbningen er blevet sendt - husk det faste notat på VL, 3
        }
    else
 {
     MsgBox, 64, Mail sendt, Mail om genåbningen er blevet sendt, 3

 }       
    return
    
l_outlook_svigt: ; tag skærmprint af P6-vindue og indsæt i ny mail til planet
    sys_genvej_start(38)
    FormatTime, dato, , d-MM-y
; GUI svigt
gui, svigt: new
    gui, svigt: +labelsvigt
    Gui svigt: Font, w600
    Gui svigt: Add, Text, x16 y0 w120 h23 +0x200, Vognløbs&nummer
    Gui svigt: Font
    Gui svigt: Add, Edit, vVL x16 y24 w120 h21, %vl%
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Font, w600
    Gui svigt: Add, Text, x161 y0 w130 h25 +0x200, &Lukket? (Afkryds én)
    Gui svigt: Font
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Add, CheckBox, vlukket x160 y24 w39 h23, &Ja
    Gui svigt: Add, Edit, vtid x200 y24 w79 h21, Hjemzone kl.
    Gui svigt: Add, CheckBox, vhelt x160 y48 w120 h23, Ja, og VL &slettet:
    Gui svigt: Add, Edit, vtid_slet x170 y68 h21, Åbningstid garanti
    ; G svigt:ui Add, CheckBox, vhelt2 x160 y72 w120, GV garanti &slettet i variabel tid ; nødvendig?
    Gui svigt: Font
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Add, Edit, vårsag x16 y72 w120 h21
    Gui svigt: Font, w600
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Font, w600
    Gui svigt: Add, Text, x304 y0 w120 h23 +0x200, Garanti eller Var.
    Gui svigt: Font
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Add, Radio, x304 y24 w120 h16, &Garanti
    Gui svigt: Add, Radio, x304 y40 w120 h32, G&arantivognløb i variabel tid
    Gui svigt: Add, Radio, vtype x304 y72 w120 h23, &Variabel
    Gui svigt: Font, w600
    Gui svigt: Add, Text, x16 y48 w120 h23 +0x200, &Årsag
    Gui svigt: Add, Text, x8 y96 h23 +0x200, &Beskrivelse
    Gui svigt: Font
    Gui svigt: Font, s9, Segoe UI
    Gui svigt: Add, Edit, vbeskrivelse x8 y120 w410 h126
    Gui svigt: Add, CheckBox, vgemt_ja x5 y261, Brug &forrige skærmklip
    Gui svigt: Add, Button, x150 y256 w60 h23 ggui_svigt_vis, &Vis
    Gui svigt: Add, Button, x210 y256 w60 h23 ggui_svigt_send, &Send
    Gui svigt: Add, text , x280 y261, Anden &Dato
    Gui svigt: Add, Edit , vny_dato x360 y256 w60, 


    ; FormatTime, tid, , HH:mm
    ; svigt := []
    gemtklip := ClipboardAll
    vl := P6_hent_vl()
    if (vl = 0)
    {
        sys_afslut_genvej()
        return
    }
    clipboard :=
    sleep 500
    SendInput, !{PrintScreen}
    sleep 500
    ; ClipWait, 3,
    klip := ClipboardAll
    ; clipwait 3, 1 ; bedre løsning?
    Gui svigt: Show, w448 h297, Svigt
    ControlFocus, Button1, Svigt
    mod_up()
Return
gui_svigt_vis:
    gui, submit

    if (ny_dato != "")
        {
            FormatTime, dato_tid, YYYYMMDDHH24MISS, d-MM-y
            dato := SubStr(ny_dato, 1 , 2) . "-" SubStr(ny_dato, -1 , 2) "-" SubStr(dato, -1 , 2)
        }
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
    if (type = 1 and lukket = 1 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 1 , % emnefelt,
        ; beskrivelse := "GV lukket kl. " tid ": " . beskrivelse
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 1 and lukket = 1 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 2, % emnefelt,
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - d. " dato
        ; MsgBox, , 3, % emnefelt,
        gui, hide
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        gui, hide
    }
    if (type = 1 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 5, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse

        gui, hide
    }
    if (type = 1 and helt = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - ikke startet op d. " dato
        ; MsgBox, , 5.1, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag !="")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - " dato
        ; MsgBox, , 6, % emnefelt,
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " d. " dato
        ; MsgBox, , 7, % emnefelt,
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 7.1, % emnefelt,
        beskrivelse := "GV slettet i variabel kørsel. Garantitid start: " tid_slet " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 8, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 9, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 3 and årsag != "")
    {
        emnefelt := "Svigt VL " vl ": " årsag " - d. " dato
        ; MsgBox, , 10, % emnefelt,
        gui, hide
    }
    if (type = 3 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        ; MsgBox, , 11, % emnefelt,
        gui, hide
    }

    outlook_template := A_ScriptDir . "\lib\svigt_template.oft"
    svigt_template := outlook.createitemfromtemplate(outlook_template)

    udklip := ImagePutFile(clipboardall, "svigt.png")
    udklip_navn := SubStr(udklip, 3)
    udklip_lok := A_ScriptDir "\" udklip_navn
    
    svigt_template.attachments.add(udklip_lok)
    svigt_template.to := "planet@midttrafik.dk"
    svigt_template.subject := emnefelt
    html_tekst =
    (
    </o:shapelayout></xml><![endif]--></head><body lang=DA link="#0563C1" vlink="#954F72" style='tab-interval:65.2pt;word-wrap:break-word'><div class=WordSection1><img id="Billede_x0020_2" src="cid:%udklip_navn%"></span></p><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-fareast-language:DA'</o:p></span></p></div><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>
    )
    svigt_template.htmlbody := beskrivelse html_tekst
    svigt_template.display
    ImageDestroy(udklip)
    gemtklip :=
    sys_afslut_genvej()
Return
gui_svigt_send:
    gui, submit

    if (ny_dato != "")
        {
            FormatTime, dato_tid, YYYYMMDDHH24MISS, d-MM-y
            dato := SubStr(ny_dato, 1 , 2) . "-" SubStr(ny_dato, -1 , 2) "-" SubStr(dato, -1 , 2)
        }
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
    if (type = 1 and lukket = 1 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 1 , % emnefelt,
        ; beskrivelse := "GV lukket kl. " tid ": " . beskrivelse
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 1 and lukket = 1 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 2, % emnefelt,
        beskrivelse := "GV lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - d. " dato
        ; MsgBox, , 3, % emnefelt,
        gui, hide
    }
    if (type = 1 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        gui, hide
    }
    if (type = 1 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 5, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse

        gui, hide
    }
    if (type = 1 and helt = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - ikke startet op d. " dato
        ; MsgBox, , 5.1, % emnefelt,
        beskrivelse := "Vl slettet. Garantitid start: " tid_slet " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag !="")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - " dato
        ; MsgBox, , 6, % emnefelt,
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 0 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " d. " dato
        ; MsgBox, , 7, % emnefelt,
        gui, hide
    }
    if (type = 2 and lukket = 0 and helt = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": ikke startet op d. " dato
        ; MsgBox, , 7.1, % emnefelt,
        beskrivelse := "GV slettet i variabel kørsel. Garantitid start: " tid_slet " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 1 and årsag != "")
    {
        emnefelt := "Svigt VL " vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
        ; MsgBox, , 8, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 2 and lukket = 1 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " " vl_type " - lukket kl. " tid " d. " dato
        ; MsgBox, , 9, % emnefelt,
        if (tid_slet != "Åbningstid garanti")
            beskrivelse := "Variabel kørsel, lukket kl. " tid ". GV start kl. " tid_slet " — " . beskrivelse
        Else
            beskrivelse := "Variabel kørsel, lukket kl. " tid " — " . beskrivelse
        gui, hide
    }
    if (type = 3 and årsag != "")
    {
        emnefelt := "Svigt VL " vl ": " årsag " - d. " dato
        ; MsgBox, , 10, % emnefelt,
        gui, hide
    }
    if (type = 3 and årsag = "")
    {
        emnefelt := "Svigt VL " vl " d. " dato
        ; MsgBox, , 11, % emnefelt,
        gui, hide
    }

    outlook_template := A_ScriptDir . "\lib\svigt_template.oft"
    svigt_template := outlook.createitemfromtemplate(outlook_template)

    udklip := ImagePutFile(clipboardall, "svigt.png")
    udklip_navn := SubStr(udklip, 3)
    udklip_lok := A_ScriptDir "\" udklip_navn
    
    svigt_template.attachments.add(udklip_lok)
    svigt_template.to := "planet@midttrafik.dk"
    svigt_template.subject := emnefelt
    html_tekst =
    (
    </o:shapelayout></xml><![endif]--></head><body lang=DA link="#0563C1" vlink="#954F72" style='tab-interval:65.2pt;word-wrap:break-word'><div class=WordSection1><img id="Billede_x0020_2" src="cid:%udklip_navn%"></span></p><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-fareast-language:DA'</o:p></span></p></div><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>
    )
    svigt_template.htmlbody := beskrivelse html_tekst
    svigt_template.send
    ImageDestroy(udklip)
    gemtklip :=
    MsgBox, 64, Mail er sendt!, Mailen er afsendt, 2
    sys_afslut_genvej()
Return


svigtEscape:
svigtClose:
Gui, hide
sys_afslut_genvej()
Return

test()
{
    WinGetTitle, tlf, ahk_exe Miralix OfficeClient.exe
    MsgBox, , , % tlf,
}

l_p6_rejsesog:
    sys_genvej_start(48)
    P6_rejsesogvindue()
    sys_afslut_genvej()
return

;; Hotstring
#IfWinActive, PLANET
    :B0:/ank::
        {
            p6_notat_hotstr("st. ankomst_tid initialer")
            return
        }
    :B0:/anko::
        {
            p6_notat_hotstr("st. ankomst_tid, overset initialer")
            return
        }
    :B0:/ankc::
        {
            p6_notat_hotstr("st. ankomst_tid, chf giver kunde besked initialer ")
            return
        }
    :B0:/ankk::
        {
            p6_notat_hotstr("st. ankomst_tid KI initialer")
            return
        }
    :B0:/anks::
        {
            p6_notat_hotstr("st. ankomst_tid SI initialer")
            return
        }
    :B0:/ankf::
        {
            p6_notat_hotstr("st. ca. ankomst_tid jf. FF initialer")
            return
        }
    :B0:/ankfk::
        {
            p6_notat_hotstr("st. ca. ankomst_tid jf. FF. KI initialer")
            return
        }
    :B0:/ankfkf::
        {
            p6_notat_hotstr("st. ca. ankomst_tid jf. FF. KFI initialer")
            return
        }
    :B0:/ankt::
        {
            p6_notat_hotstr("st. ankomst_tid grundet trafik initialer")
            return
        }
    :B0:/anktk::
        {
            p6_notat_hotstr("st. ankomst_tid grundet trafik. KI initialer")
            return
        }
    :B0:/anktc::
        {
            p6_notat_hotstr("st. ankomst_tid grundet trafik. Chf informerer kunde initialer")
            return
        }
    :B0:/ankv::
        {
            p6_notat_hotstr("st. ankomst_tid grundet vejarbejde initialer")
            return
        }
    :B0:/ankvk::
        {
            p6_notat_hotstr("st. ankomst_tid grundet vejarbejde. KI initialer")
            return
        }
    :B0:/ankvkf::
        {
            p6_notat_hotstr("st. ankomst_tid grundet vejarbejde. KFI initialer")
            return
        }
    :B0:/anka::
        {
            p6_notat_hotstr("st. ankomst_tid, problemer m. adresse. initialer")
            return
        }
    :B0:/ankak::
        {
            p6_notat_hotstr("st. ankomst_tid, problemer m. adresse. KI initialer")
            return
        }
    :B0:/ankik::
        {
            p6_notat_hotstr("st. ankomst_tid, ikke kvitteret initialer")
            return
        }
    :B0:/ankik::
        {
            p6_notat_hotstr("st. ankomst_tid, overset initialer")
            return
        }
    :b0:/repl::
        {
            p6_notat_hotstr("st. replaneret repl_tid initialer")
            return
        }
    :b0:/replk::
        {
            p6_notat_hotstr("st. replaneret repl_tid KI ankomst_tid initialer")
            return
        }
    :b0:/replkf::
        {
            p6_notat_hotstr("st. replaneret repl_tid KFI initialer")
            return
        }
    :b0:/tom::
        {
            p6_notat_hotstr("TOM st. initialer")
            return
        }
    :b0:/lad::
        {
            p6_notat_hotstr("ladepause udvidet til pause_tid initialer")
            return
        }
    :b0:/låso::
        {
            p6_notat_hotstr("Låst op initialer")
            return
        }
    :b0:/låsv::
        {
            p6_notat_hotstr("Låst, værksted initialer")
            return
        }
    :b0:/ci::
        {
            p6_notat_hotstr("chf inf initialer")
            return
        }
    :b0:/retf::
        {
            p6_notat_hotstr("st. tid rettet jf. FF initialer")
            return
        }
    :b0:/ryk::
        {
            p6_notat_hotstr("st. rykker initialer")
            return
        }
#IfWinActive, Vognløbsnotering
#IfWinActive, PLANET
    ::/mt::
        {
            FormatTime, tid, YYYYMMDDHH24MISS, HHmm
            initialer := "/mt" A_UserName tid
            SendInput, %initialer%
            return
        }
#IfWinActive

::vllp::Låst, ingen kontakt til chf, privatrejse ikke udråbt
::bsgs::Glemt slettet retur
::rgef::Rejsegaranti, egenbetaling fjernet
::vlaok::Alarm st OK
::svigtudråb::VL ikke startet op, ingen kontakt til chf. VL ryddet og låst.

;; TEST

^+a::databaseview("%A_linefile%\..\db\bruger_ops.tsv")
; Tag alarm, gå til næste VL på liste
p6_tag_alarm()
{
    SendInput, ^l
    SendInput, !{Down}
    return
}
; Samme som p6_tag_alarm, aktiv fra boksen med repl-VL
p6_tag_alarm_vl_box()
{
    SendInput, {enter}
    sleep 50
    SendInput, ^l
    SendInput, !{Down}
    return
}
; Fra notatvindue - indsætter samme notat igen, til når der timeoutes - skal der tages hensyn til nye notater, der kan være skrevet i mellemtiden?

; +^e::FlexFinder_addresse()
; FlexFinder_addresse()
; {
;     ; SendInput, +{tab} {Down} {Tab}
;     If (WinExist("FlexDanmark FlexFinder"))
;     {
;         sleep 200
;         WinActivate, FlexDanmark FlexFinder
;         winwaitactive, FlexDanmark FlexFinder
;         sleep 40
;         SendInput, {Home}
;         sleep 400
;         SendInput, {PgUp}
;         sleep 200
;         WinGetPos, W_X, W_Y, , , FlexDanmark FlexFinder, , ,
;         if(W_X = "1920" or W_X = "-1920")
;         {
;             ; PixelSearch, Px, Py, 0, 0, , 11, 0xF26C5B, 0, Fast
;             sleep 200
;             click %Px% %Py%
;             sleep 999
;             SendInput, {tab 3} {down} {tab}
;             ; ControlClick, x322 y100, FlexDanmark FlexFinder
;             sleep 40
;             return
;         }
;         Else
;         {
;             ; PixelSearch, Px, Py, 1097, 74, 1202, 123, 0x5B6C2, 0, Fast ; Virker ikke i fuld skærm. ControlClick i stedet?
;             ImageSearch, Ix, Iy, 0 , 0, A_ScreenWidth , A_ScreenHeight *100 , /lib/ff.png
;             MsgBox, , , % ix
;             sleep 200
;             click %Px% %Py%
;             sleep 200
;             ControlClick, x322 y100, FlexDanmark FlexFinder
;             sleep 40
;             SendInput, +{tab}{down}{tab}
;             return
;         }
;         ; SendInput, {CtrlUp}{ShiftUp} ; for at undgå at de hænger fast
;     }
;     Else
;         MsgBox, , FlexFinder, Flexfinder ikke åben (skal være den forreste fane)
;     Return
; }
;

; ^z::
; {
; if WinExist("+ ahk_exe Miralix OfficeClient.exe")
;     MsgBox, , , kald,
; Else
;     MsgBox, , , ikke kald
; }

