#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode, 1 ; matcher så længe et ord er der

;; TODO

; Global læg på

; gemt-klip-funktion ved al brug af clipboard

; FUNKTIONER

; Gem Clipboard

;P6

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
    gemtklip := ClipboardAll
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
    ClipWait, 1, 0
    vl_tlf := Clipboard
    Clipboard :=
    Clipboard = %gemtklip%
    gemtklip :=
    ;msgbox %vl_tlf%
    Return vl_tlf
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
; ***
; Noterer intialer, fjerner dem hvis første ord i notering er initialer
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
    ; deler notering op i array med ord delt i mellemrum
    notering_array := StrSplit(notering, A_Space)
    ; notering_array := StrSplit(notering, /)
    ; tjekker for initialer uden tid i første ord i notering
    ; falsk positiv, hvis der er skrevet ud i ét, uden mellemrum
    if InStr(notering_array[1], initialer_udentid, 0, 1)
    ; hvis ja, fjerner de første 11 bogstaver (= initialer med tid) ? kan det laves smartere?
    {
        StringTrimLeft, noteringuden, notering, 11
        Clipboard :=
        sleep 200
        Clipboard := noteringuden
        sendinput ^a^v
        sleep 100
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
P6_kørselsaftale()
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
    kørselsaftale := P6_kørselsaftale()
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
P6_udfyldKA()
{
    P6_alt_menu()
    sleep 40
    SendInput, tk
    sleep 40
    SendInput !{F5}
    return
}

;Telefon

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

;  ***
;Træk tlf fra markeret tekst, hvis 10 cifre skær de to første af, return som variabel
Telenor_clipboard()
{
    clipboard := ""
    Sendinput ^c
    ClipWait
    Telefon := Clipboard
    Ciffer_antal := StrLen(Telefon)
    if (Ciffer_antal = 10)
        rentelefon := Substr(Telefon, 3, 8)
    Else
        rentelefon := telefon
    return rentelefon
}
; ***
; Sæt kopieret tlf i Trio
Trio_opkald()
{
    ControlClick, Edit2, ahk_class Addressbook
    sleep 40
    SendInput, ^v{enter}
    Return
}

; *
; Læg på i Trio
Trio_afslutopkald()
{
    WinActivate, ahk_class AccessBar
    sleep 40
    SendInput, {NumpadSub}

    return
}
; Misc

; ***
; Åbn ny mail i outlook. Kræver nymail.lnk i samme mappe som script.
Outlook_nymail()
{  
    Run, nymail.lnk, , , 
    Return
}

; ***
; Tag skærmprint af aktivt vindue
screenshot_aktivvindue()
{
    SendInput, !{PrintScreen}
    Return
}

;; HOTKEYS

+^e::

return

+Escape::
ExitApp
Return


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
        P6_udfyldKA()
    Return
#IfWinActive

;træk markeret tekst til Vl-tlf - hvis 10 cifre skær de to første af
; ***
+F3::
    telefon := Telenor_clipboard()
    clipboard := telefon
    ClipWait, 1, 0
    WinActivate, PLANET
    P6_tlf_vl()
return

;træk tlf til rejsesøg
; ***
+F4::
    telefon := Telenor_clipboard()
    clipboard := telefon
    ClipWait, 1, 0
    WinActivate, PLANET
    P6_rejsesog_tlf()
return

#IfWinActive
; *
;træk tlf fra aktiv planbillede
#IfWinActive PLANET
    +F5::
    vl_tlf := P6_hent_vl_tlf()
    clipboard := vl_tlf
    ClipWait, 2, 0
    Trio_opkald()
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
+F1::
#IfWinActive PLANET
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
ClipWait, 2, 1
SendInput, {Up}{Up}
Return

; Minus på numpad afslutter Trioopkald når P6 aktivt
NumpadSub::
#IfWinActive
Trio_afslutopkald()
WinActivate, PLANET

    Return
#IfWinActive

;https://www.autohotkey.com/docs/v1/lib/WinActivate.htm

;HOTSTRINGS
::vllp::Låst, ingen kontakt til chf, privatrejse ikke udråbt
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
