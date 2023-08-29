#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode, 1 ; matcher så længe et ord er der
#SingleInstance, force
; Define the group: gruppe
GroupAdd, gruppe, PLANET
GroupAdd, gruppe, ahk_class Chrome_WidgetWin_1
GroupAdd, gruppe, ahk_class AccessBar
GroupAdd, gruppe, ahk_class Agent Main GUI
GroupAdd, gruppe, ahk_class Addressbook

;; TODO

; Global læg på ✔️
; ring op til VM
; gemt-klip-funktion ved al brug af clipboard
; Luk om x antal minutter


;; kendte fejl
; P6_initialer sletter ikke, hvis initialerne er eneste ord i notering

; FUNKTIONER

; Gem Clipboard

;; P6

; ***
; P6 alt menu
P6_alt_menu()
{
    keywait Shift ; for ikke at ødelægge shiftgenveje
    SendInput, {Alt}
}

; ***
; Åben planbillede
P6_Planvindue()
{
    P6_alt_menu()
    SendInput, tp
    return
}

; ***
; Åben renset rejsesøg
P6_rejsesogvindue()
{
    P6_alt_menu()
    SendInput rr^t
    Return
}

; ***
; åben alarmvinduet, ny liste alle alarmer, blad til første
P6_alarmer()
{
    P6_alt_menu()
    sendinput ta
    sleep 40
    SendInput, ^{Delete}
    SendInput, !k
    sleep 200
    sleep 200
    SendInput, +{PgDn}
    sleep 40
    SendInput, ^l
    P6_Planvindue()
    sleep 200
    SendInput, !{Down}
    return
}

; ***
; åben alarmvinduet, ny liste alle udråbsalarmer, blad til første
P6_udraabsalarmer()
{
    P6_alt_menu()
    sendinput ta
    sleep 40
    SendInput, ^{Delete}
    SendInput, !u
    sleep 200
    sleep 200
    SendInput, +{PgDn}
    sleep 40
    SendInput, ^l
    P6_Planvindue()
    sleep 200
    SendInput, !{Down}
    return
}

; ***
; gå i rent rejsesøg med karet i telefonfelt
P6_rejsesog_tlf()
{
    P6_rejsesogvindue()
    sleep 300
    SendInput {tab}{tab}^v^r

    Return
}
; ***
;
P6_hent_vl_tlf()
{
    P6_Planvindue()
    SendInput ^{F12}
    sleep 800
    sendinput ^æ
    sleep 200
    SendInput {Enter}{Enter}
    ; Sleep 40
    SendInput !ø
    ; sleep 40
    Clipboard :=
    SendInput {tab}{tab}^c{enter}
    ClipWait, 2, 0
    vl_tlf := Clipboard
    Return vl_tlf
}

; ***
; P6 hent VM tlf
P6_hent_vm_tlf()
{
    P6_vis_KA()()
    sleep 200
    sendinput ^æ
    sleep 200
    SendInput {Enter}
    ; Sleep 40
    SendInput !a
    ; sleep 40
    Clipboard :=
    SendInput {tab}{tab}{tab}{tab}^c{enter}
    ClipWait, 2, 0
    SendInput ^a
    vm_tlf := Clipboard
    Return vm_tlf
}
; ***
;indsæt clipboard i vl-tlf
P6_tlf_vl()
{
    P6_Planvindue()
    sleep 200
    SendInput ^{F12}
    sleep 800
    sendinput ^æ
    sleep 200
    SendInput {Enter}{Enter}
    ; Sleep 40
    SendInput !ø
    ; sleep 40
    SendInput {tab}{tab}^v{enter}
    return
}

;  ***
;indsæt clipboard i vl-tlf dagen efterfølgende
P6_tlf_vl_efter()
{
    WinActivate PLANET version 6 Jylland-Fyn DRIFT
    SendInput, {Tab}
    sleep 200
    SendInput, !{right}^æ
    Sleep 40
    SendInput {Enter}{Enter}
    Sleep 40
    SendInput !ø
    sleep 40
    SendInput {tab}{tab}^v{Enter}

    return
}
; ** Virker ikke hvis noteringen er eneste ord
; Noterer intialer, fjerner dem hvis første ord i notering er initialer

; midlertidig løsning at kun eksakt tidspunkt skæres fra?
P6_initialer()
{
    FormatTime, Time, ,HHmm tt ;definerer format på tid/dato
    initialer = /mt%A_userName%%time%
    initialer_udentid =mt%A_userName%
    P6_Planvindue()
    sleep 40
    sendinput ^n
    sleep 1400
    clipboard :=
    SendInput, ^a^c
    ClipWait, 1, 0
    sleep 40
    notering := Clipboard
    sleep 400
    ; deler notering op i array med ord delt i mellemrum
    ; notering_array := StrSplit(notering, A_Space)
    notering_array := StrSplit(notering, A_space, "/")
    sleep 400
    MsgBox, , Felt 1, % notering_array.1, 
    MsgBox, , Felt 2, % notering_array.2, 
    ; tjekker for initialer uden tid i første ord i notering
    ; falsk positiv, hvis der er skrevet ud i ét, uden mellemrum
    ; hvis ja, fjerner de første 11 bogstaver (= initialer med tid) ? kan det laves smartere?
    if InStr(notering_array[1], initialer_udentid, 0, 1)
    {
        StringTrimLeft, noteringuden, notering, 11
        Clipboard :=
        sleep 200
        Clipboard := noteringuden
        sendinput ^a^v
        sleep 200
        SendInput, !o
        return
    }
    ;indsætter initialer med tid
    Else
        Clipboard :=
    sleep 40
    Clipboard := initialer
    ClipWait, 1, 0
    SendInput, {Left}
    Sendinput ^v
    SendInput, %A_Space%
    sleep 100
    SendInput, !o
}

; ** kan gemtklip-funktion skrives bedre?
;Indsæt initialer med efterf. kommentar, behold tidligere klip
P6_initialer_skriv()
{
    gemtklip := ClipboardAll
    FormatTime, Time, ,HHmm tt ;definerer format på tid/dato
    initialer = /mt%A_userName%%time%
    P6_Planvindue()
    sleep 40
    sendinput ^n
    sleep 40
    Clipboard := initialer
    Sendinput ^v
    Sendinput %A_space%
    Sendinput {home}
    sleep 2000
    Clipboard = %gemtklip%
    gemtklip := ""
    return
}

; ***
;Kørselsaftale på VL til clipboard
P6_k_aftale()
{
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    Sendinput !tp!k
    clipboard := ""
    Sendinput +{F10}c
    ClipWait
    kørselsaftale := clipboard
    return kørselsaftale
}

; ***
;styresystem til clipboard
P6_styresystem()
{
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    Sendinput !tp!k{tab}
    clipboard := ""
    Sendinput +{F10}c
    ClipWait
    styresystem := clipboard
    return styresystem
}

;  ***
;åben tekst m. kørselsaftale udfyldt
P6_tekstTilChf()
{
    ;WinActivate PLANET version 6   Jylland-Fyn DRIFT
    kørselsaftale := P6_k_aftale()
    styresystem := P6_styresystem()
    Sendinput !tt^k
    Sleep 100
    Sendinput !k
    clipboard := kørselsaftale
    sleep 40
    Sendinput +{F10}p{tab}
    sleep 200
    clipboard := styresystem
    Sendinput +{F10}p{Tab}

    return
}

;  ***
; Udfyld kørselsaftale for aktivt planbillede
P6_vis_KA()
{
    P6_alt_menu()
    sleep 40
    SendInput, tk
    sleep 40
    SendInput !{F5}
    return
}

; ***
; Tag skærmprint af aktivt vindue
screenshot_aktivvindue()
{
    SendInput, !{PrintScreen}
    ClipWait, 3, 1
    Return
}


;; Telenor

;træk telenor indgåend
;virker ikke
Telenor()
{
    WinActivate, Telenor KontaktCenter
    ControlClick, x179 y491, Telenor KontaktCenter,, Left,2,
    sleep 100
    ControlClick, x179 y491, Telenor KontaktCenter,, Left,2,
    sleep 100
    SendInput {tab}
    SendInput {tab}
    return
}


;; Trio
; ***
; Sæt kopieret tlf i Trio
Trio_opkald()
{
    ControlClick, Edit2, ahk_class Addressbook
    ; controlsend, Edit2, 50541537, ahk_class Addressbook
    sleep 40
    WinActivate, ahk_class Addressbook
    sleep 40
    SendInput, ^v
    sleep 40
    SendInput, +{enter} ; undgår kobling ved igangværende opkald
    Return
}

; ***
; Sæt kopieret tlf i Trio
; Trio_opkald()
; {
;     ControlClick, Edit2, ahk_class Addressbook
;     sleep 40
;     SendInput, ^v{enter}
;     Return
; }

; ***
; Læg på i Trio
Trio_afslutopkald()
{
    WinActivate, ahk_class AccessBar
    sleep 40
    SendInput, {NumpadSub}

    return
}

; **
; Trio hop til efterbehandling
trio_efterbehandling()
{
    WinActivate, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 8
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; **
; Trio hop til midt uden overløb
trio_udenov()
{
    WinActivate, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 3
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; **
; Trio hop til alarm
trio_alarm()
{
    WinActivate, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 7
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; **
; Trio hop til pause
trio_pause()
{
    WinActivate, ahk_class AccessBar
    sleep 100
    SendInput, {F3}
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; **
; Trio hop til klar
trio_klar()
{
    WinActivate, ahk_class AccessBar
    Sleep 100
    SendInput, {F4}
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; **
; Trio hop til frokost
trio_frokost()
{
    WinActivate, ahk_class Agent Main GUI
    sleep 40
    SendInput, !f
    sleep 40
    SendInput, o
    sleep 40
    SendInput, 9
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

; Trio skift mellem pause og klar

trio_pauseklar()
{
    WinActivate, ahk_class AccessBar
    Sleep 200
    SendInput, {F3}
    sleep 200
    SendInput, {F4}
    WinActivate, PLANET
    ; ControlClick, ComboBox2, ahk_class Agent Main GUI
    ; sleep 400
    ; Controlsend, ComboBox2, ahk_class AccessBar, '
    ; ControlClick, x68 y185, ahk_class AccessBar
    Return
}

;  ***
;Træk tlf fra Trio indkomne kald
Trio_clipboard()
{
    clipboard := ""
    WinActivate, ahk_class AccessBar, , , 
    Sendinput !+k
    ClipWait
    Telefon := Clipboard
    Ciffer_antal := StrLen(Telefon)
    if (Ciffer_antal = 11)
        rentelefon := Substr(Telefon, 4, 8)
    Else
        rentelefon := telefon
    return rentelefon
}

;; Flexfinder

; *
; Kørselsaftale til flexfinder
Flexfinder_opslag()
{
    P6_k_aftale()
    sleep 200
    WinActivate, FlexDanmark FlexFinder
    SendInput, {F5}
    sleep 800
    SendInput, {Tab}{Tab}{Tab}{Tab}{tab}
    sleep 200
    SendInput, ^v
    ClipWait, 2, 0
    KeyWait, Enter, D
    sleep 200
    WinActivate, PLANET
}

; Misc


;; Outlook
; ***
; Åbn ny mail i outlook. Kræver nymail.lnk i samme mappe som script.
Outlook_nymail()
{  
    Run, nymail.lnk, , , 
    Return
}

;
+^p::
{
Gui,Add,Button,vButton1,AUH
Gui,Add,Button,vButton2,RHG
Gui,Add,Button,vButton3,Randers Sygehus
Gui,Add,Button,vButton4,Viborg Sygehus
Gui,Add,Button,vButton5,Horsens Sygehus
Gui,Show
funb1:=Func("fun").Bind("Button1","test1")
funb2:=Func("fun").Bind("Button2","test2")
GuiControl,+g,Button1,%funb1%
GuiControl,+g,Button2,%funb2%
return
Fun(p*){
  MsgBox % p.1 "`n" p.2
}
return

}
GUI_sygehus()
{
 

return
}

;; Testknap
; +^e::
; return


;; HOTKEYS


+Escape::
ExitApp
Return

;; PLANET
#IfWinActive PLANET
    F2::
        P6_initialer()
    Return
#IfWinActive


; skriv initialer og forsæt notering.
#IfWinActive PLANET
    +F2::
        P6_initialer_skriv()
    return

#IfWinActive

;Vis kørselsaftale for aktivt vognløb
#IfWinActive PLANET
    F3::
        P6_vis_KA()
    Return
#IfWinActive

;træk trio-opkald til Vl-tl
; ***
+F3::
    telefon := Trio_clipboard()
    clipboard := telefon
    ClipWait, 1, 0
    WinActivate, PLANET
    P6_tlf_vl()
return

;træk tlf til rejsesøg
; ***
+F4::
    telefon := Trio_clipboard()
    clipboard := telefon
    ClipWait, 1, 0
    WinActivate, PLANET
    P6_rejsesog_tlf()
return

#IfWinActive

; *
;træk tlf fra aktiv planbillede, ring op i Trio
#IfWinActive PLANET
    +F5::
    gemtklip := ClipboardAll
    vl_tlf := P6_hent_vl_tlf()
    clipboard := vl_tlf
    ClipWait, 2, 0
    sleep 200
    Trio_opkald()
    Clipboard = %gemtklip%
    ClipWait, 2, 1
    gemtklip :=
    sleep 800
    WinActivate, PLANET
    P6_Planvindue()
    Return
#IfWinActive

; ***
; træk vm-tlf fra aktivt planbillede, ring op i Trio
#IfWinActive PLANET
    ^+F5::
    gemtklip := ClipboardAll
    vm_tlf := P6_hent_vm_tlf()
    clipboard := vm_tlf
    ClipWait, 2, 0
    sleep 200
    Trio_opkald()
    sleep 500
    Clipboard = %gemtklip%
    ClipWait, 2, 1
    gemtklip :=
    sleep 800
    WinActivate, PLANET
    P6_Planvindue()
    Return
#IfWinActive

;alarmer
#IfWinActive PLANET
    F7::
        P6_alarmer()
    return
#IfWinActive

;udråbsalarmer
#IfWinActive PLANET
    +F7::
        P6_udraabsalarmer()
    return
#IfWinActive

#IfWinActive PLANET
    +^t::
        P6_tekstTilChf()
    return
#IfWinActive

; tag skærmprint af P6-vindue og indsæt i ny mail til planet
#IfWinActive PLANET
+F1::
gemtklip := ClipboardAll
sleep 400
screenshot_aktivvindue()
Outlook_nymail()
sleep 1000
SendInput, pl
sleep 250
SendInput, {Tab}
sleep 40
SendInput, {Tab}{Tab}{Tab}{Enter}{Enter}
sleep 40
SendInput, ^v
SendInput, {Up}{Up}
sleep 2000
Clipboard = %gemtklip%
ClipWait, 2, 1
gemtklip :=
Return
#IfWinActive


;https://www.autohotkey.com/docs/v1/lib/WinActivate.htm


;; Trio-Hotkey
||
#IfWinActive ahk_group gruppe
    ^1::
        trio_klar()
    Return
#IfWinActive

#IfWinActive ahk_group gruppe
    ^0::
        trio_pause()
    Return
#IfWinActive

#IfWinActive ahk_group gruppe
    ^2::
        trio_udenov()
    Return
#IfWinActive


#IfWinActive ahk_group gruppe
    ^3::
        trio_efterbehandling()
        trio_pauseklar()
    Return
#IfWinActive


#IfWinActive ahk_group gruppe
    ^4::
        trio_alarm()
    Return
#IfWinActive

#IfWinActive ahk_group gruppe
    ^5::
        trio_frokost()
    Return
#IfWinActive

; Kald det markerede nummer i trio, global
    !q::
    SendInput, ^c
    sleep 100
    Trio_opkald()
    Return

    ; Minus på numpad afslutter Trioopkald global (Skal der tilbage til P6?)
; #IfWinActive PLANET
+NumpadSub::
Trio_afslutopkald()
sleep 200
WinActivate, PLANET
Return
; #IfWinActive

;; Flexfinder
#IfWinActive PLANET
    ^+f::
        Flexfinder_opslag()
    Return
#IfWinActive

;; GUI


;; HOTSTRINGS

; #IfWinActive PLANET


; Hvorfor virker ifwinactive ikke?
; #IfWinActive PLANET
::vllp::Låst, ingen kontakt til chf, privatrejse ikke udråbt
::bsgs::Glemt slettet retur
::rgef::Rejsegaranti, egenbetaling fjernet
; #IfWinActive
;    Clipboard := "Låst, ingen kontakt til chf, privatrejse ikke udråbt"
;	ClipWait
;    Sendinput ^v

;return

::/mt::
    {
        initialer = /mt%A_userName%%time% %A_space%
        gemtklip := Clipboard
        Clipboard := initialer
        ClipWait, 1, 0
        Sendinput ^v
        sleep 800
        Clipboard := gemtklip
        return
    }





; +r::
;     Reload
;     sleep 2000
; Return