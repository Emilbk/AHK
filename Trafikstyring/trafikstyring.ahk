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
;   bruger_genvej  telenor_opr     telenor_ahk

;; hotkeydef.
; globale genveje                                           ; Standard-opsætning
Hotkey, % bruger_genvej.4, l_telenor_p6_opslag ; !w
Hotkey, % bruger_genvej.30, l_trio_afslut_opkald ; Numpad -
Hotkey, % bruger_genvej.31, l_trio_afslut_opkaldB ; Numpad -
Hotkey, % bruger_genvej.32, l_trio_til_p6 ; +F4
Hotkey, % bruger_genvej.33, l_quitAHK ; +escape
Hotkey, % bruger_genvej.34, l_p6_aktiver ; +!p

Hotkey, IfWinActive, PLANET
Hotkey, % bruger_genvej.46, l_outlook_svigt ; +F1
Hotkey, % bruger_genvej.5, l_p6_initialer ; F2
Hotkey, % bruger_genvej.6, l_p6_initialer_skriv ; +F2
Hotkey, % bruger_genvej.7, l_p6_vis_k_aftale ; F3
Hotkey, % bruger_genvej.8, l_p6_ret_vl_tlf ; +F3
Hotkey, % bruger_genvej.9, l_p6_vaelg_vl ; ^F3
Hotkey, % bruger_genvej.10, l_p6_søg_vl ; F4
Hotkey, % bruger_genvej.11, l_p6_vl_ring_op ; +F5
Hotkey, % bruger_genvej.12, l_p6_vm_ring_op ; ^+F5
Hotkey, % bruger_genvej.13, l_p6_vl_luk ; #F5
Hotkey, % bruger_genvej.14, l_p6_alarmer ; F7
Hotkey, % bruger_genvej.15, l_p6_udråbsalarmer ; +F7
Hotkey, % bruger_genvej.16, l_p6_ring_til_kunde ; +F8
Hotkey, % bruger_genvej.17, l_p6_udregn_minut ; #t
Hotkey, % bruger_genvej.18, l_p6_sygehus_ring_op ; ^+s
Hotkey, % bruger_genvej.19, l_p6_central_ring_op ; ^+c
Hotkey, % bruger_genvej.20, l_p6_tekst_til_chf ; ^+t
Hotkey, % bruger_genvej.36, l_flexf_fra_p6 ; +^F
Hotkey, IfWinActive

; Trio
Hotkey, IfWinActive, ahk_group gruppe
Hotkey, % bruger_genvej.22, l_trio_pause ; ^0
Hotkey, % bruger_genvej.23, l_trio_klar ; ^1
Hotkey, % bruger_genvej.24, l_trio_udenov ; ^2
Hotkey, % bruger_genvej.25, l_trio_efterbehandling ; ^3
Hotkey, % bruger_genvej.26, l_trio_alarm ; ^4
Hotkey, % bruger_genvej.27, l_trio_frokost ; ^5
Hotkey, % bruger_genvej.28, l_trio_opkald_markeret ; !q
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

; settings

;; FUNKTIONER
;; P6

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
; P6 alt menu
P6_aktiver()
{
    WinActivate, PLANET
}

P6_alt_menu()
{
    keywait Shift ; for ikke at ødelægge shiftgenveje
    keywait ctrl
    keywait alt
    SendInput, {Alt}
    sleep 200
}

; ***
; Åben planbillede
P6_planvindue()
{
    global s

    P6_alt_menu()
    SendInput, tp
    sleep s * 100
    return
}

; ***
; Åben renset rejsesøg
P6_rejsesogvindue()
{
    global s

    P6_alt_menu()
    sleep 300
    SendInput rr^t
    sleep s * 100
    Return
}

;  ***
; Vis kørselsaftale for aktivt planbillede
P6_vis_k()
{
    global s
    P6_alt_menu()
    sleep s * 40
    SendInput, tk
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
}
; Hent VL-nummer
P6_hent_vl()
{
    global s
    clipboard := ""
    SendInput, !l
    sleep 20 ; ikke P6-afhængig
    SendInput, +{F10}c
    ClipWait, 2, 0
    vl := Clipboard
    return vl
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
}

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
}

; ***
; åben alarmvinduet, ny liste alle alarmer, blad til første
P6_alarmer()
{
    global s
    P6_alt_menu()
    sendinput ta
    sleep s * 40
    SendInput, !k
    SendInput, ^{Delete}
    sleep s * 100
    SendInput, {PgUp}
    SendInput, +^{Down}
    sleep s * 200
    SendInput, ^l
    P6_Planvindue()
    sleep s * 200
    SendInput, !{Down}

    return
}

; ***
; åben alarmvinduet, ny liste alle udråbsalarmer, blad til første
P6_udraabsalarmer()
{
    global s
    P6_alt_menu()
    sendinput ta
    sleep s * 40
    SendInput, !u
    SendInput, ^{Delete}
    sleep s * 200
    SendInput, {PgUp}
    SendInput, +^{Down}
    sleep s * 40
    SendInput, ^l
    P6_Planvindue()
    sleep s * 200
    SendInput, !{Down}
    return
}

P6_notat(byref tekst:="")
{
    P6_planvindue()
    SendInput, ^n
    sleep 500
    SendInput, %tekst%
    SendInput, !o

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
    P6_Planvindue()
    SendInput ^{F12}
    sleep s * 1500
    sendinput ^æ
    sleep s * 200
    SendInput {Enter}{Enter}
    sleep s * 40
    SendInput !ø
    sleep s * 40
    Clipboard :=
    SendInput {tab}{tab}^c{enter}
    ClipWait, 2, 0
    vl_tlf := Clipboard
    Return vl_tlf
}
return
; ***
; P6 hent VM tlf
P6_hent_vm_tlf()
{
    global s
    P6_vis_k()()
    sleep * 200
    sendinput ^æ
    sleep * 200
    SendInput {Enter}
    ; sleep * 40
    SendInput !a
    ; sleep * 40
    Clipboard :=
    SendInput {tab}{tab}{tab}{tab}^c{enter}
    ClipWait, 2, 0
    SendInput ^a
    vm_tlf := Clipboard
    Return vm_tlf
}

P6_hent_vl_fra_tlf(ByRef tlf:="")
{

    række := DataBasefind( "%A_linefile%\..\db\VL_tlf.txt", tlf)
    vl := databaseget("%A_linefile%\..\db\VL_tlf.txt", række.1, 2)
    if (række.1 is number) ; hvorfor virker den ikke med true/false?
    {
        vl := StrSplit(vl, "_") ;vl.1 k, vl.2 s
        Return vl
    }
    else
        vl = 0
    Return vl
}

; ***
;indsæt clipboard i vl-tlf
P6_ret_tlf_vl(ByRef telefon:=" ")
{
    P6_Planvindue()
    sleep s * 50
    SendInput ^{F12}
    sleep s * 800
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
    WinActivate PLANET version 6 Jylland-Fyn DRIFT
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
    FormatTime, tid, ,HHmm ;definerer format på tid/dato
    initialer = /mt%A_userName%%tid%
    initialer_udentid =/mt%A_userName%
    P6_Planvindue()
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
    FormatTime, Time, ,HHmm tt ;definerer format på tid/dato
    initialer = /mt%A_userName%%time%
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
    WinActivate PLANET
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
    p6_input_sidste_slut_ops := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1,4)
    KeyWait, Ctrl,
    KeyWait, Shift,
    EnvAdd, nu_plus_5, 5, minutes
    FormatTime, nu_plus_5, %nu_plus_5%, HHmm
    if (p6_input_sidste_slut_ops = "1")
    {
        InputBox, sidste_stop, Sidste stop, Tast tid for sidste stop (4 cifre)
        if (ErrorLevel = "1")
            Return
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
        InputBox, tid_til_hjemzone, Tid til hjemzone, Tid til hjemzone i minutter
        if (ErrorLevel = "1")
            Return
        if (tid_til_hjemzone = "" )
        {
            FormatTime, sidste_stop, %sidste_stop%, HHmm
            return sidste_stop
        }
        EnvAdd, sidste_stop, tid_til_hjemzone + 2, minutes
        FormatTime, sidste_stop, %sidste_stop%, HHmm
        return sidste_stop
    }
    if (p6_input_sidste_slut_ops = "0")
    {
        Input, sidste_stop, T5, {Enter}{escape}
        if (ErrorLevel = "EndKey:Escape")
            Return
        if (ErrorLevel = "Timeout")
        {MsgBox, , Timeout , Det tog for lang tid.
            return 0
        }
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
        Input, tid_til_hjemzone, T5, {enter}{Escape},
        if (ErrorLevel = "EndKey:Escape")
            Return
        if (ErrorLevel = "Timeout")
        {MsgBox, , Timeout , Det tog for lang tid.
            return 0
        }
        if (tid_til_hjemzone = "" )
        {
            FormatTime, sidste_stop, %sidste_stop%, HHmm
            return sidste_stop
        }
        EnvAdd, sidste_stop, tid_til_hjemzone + 2, minutes
        FormatTime, sidste_stop, %sidste_stop%, HHmm
        return sidste_stop
    }
    Input, sidste_stop, T5, {Enter}{escape}
    if (ErrorLevel = "EndKey:Escape")
        Return
    if (ErrorLevel = "Timeout")
    {MsgBox, , Timeout , Det tog for lang tid.
        return 0
    }
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
    Input, tid_til_hjemzone, T5, {enter}{Escape},
    if (ErrorLevel = "EndKey:Escape")
        Return
    if (ErrorLevel = "Timeout")
    {MsgBox, , Timeout , Det tog for lang tid.
        return 0
    }
    if (tid_til_hjemzone = "" )
    {
        FormatTime, sidste_stop, %sidste_stop%, HHmm
        return sidste_stop
    }
    EnvAdd, sidste_stop, tid_til_hjemzone + 2, minutes
    FormatTime, sidste_stop, %sidste_stop%, HHmm
    return sidste_stop
}

; læg minuttal til klokkeslæt eller minuttal til minuttal.
P6_udregn_minut()
{
    resultat := []
    tidA := ; HHmm, starttid. Enten fire cifre for klokkeslæt, mellem 1 og 3 cifre for minuttertal.
    tidB := ; mm, tillægstid. Minuttal
    tidC := ; resultat
    p6_udregn_minut_ops := 1 ; 1 - med inputbox, 0 med input
    if (p6_udregn_minut_ops = 1)
    {
        InputBox, tidA, Udgangspunkt, Skriv tiden`, der skal lægges noget til. `nKlokkeslæt: 4 cifre ud i ét`, minuttal: 3 til 1 ciffer ud i ét. `n `n F. eks: `n Klokken 13:34 skrives 1334 `n 231 minutter skrives 231, , , 240
        if (ErrorLevel != 0)
            return
        if (tida = "")
            tida := "0"
        InputBox, tidB, Tid `, der skal lægges til., Skriv tid`, der skal lægges til. Minuttal ud i ét (- foran`, hvis der skal trækkes fra).,
        if (ErrorLevel != 0)
            return
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
        Input, tida,,{enter}
        if (ErrorLevel = "Match")
            return
        if (tida = "")
            tida := "0"
        Input, tidb,,{enter}
        if (ErrorLevel = "Match")
            return
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
}
; luk vl på variabel tid
P6_vl_luk(ByRef tid:="")
{
    global s

    P6_Planvindue()
    sleep s * 100
    SendInput, ^{F12}
    sleep s * 100
    SendInput, ^æ
    sleep s * 40
    clipboard :=
    SendInput, +{F10}c
    ClipWait, 3, 0
    vl := clipboard
    if (StrLen(vl) <= 3)
    {
        SendInput, Keys{Enter}
        FormatTime, dato, YYYYMMDDHH24MISS, d
        SendInput, %dato%
        SendInput, {tab}
        SendInput, %tid%
        SendInput, {enter}{enter}
        return
    }
    Else
    {
        SendInput, Keys{Enter}{Tab 3}
        SendInput, %tid%
        SendInput, {tab}{tab}
        SendInput, %tid%
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

; ***
; Tag skærmprint af aktivt vindue
screenshot_aktivt_vindue()
{
    SendInput, !{PrintScreen}
    ClipWait, 3, 1
    Return
}

;; Telenor

;; Trio
; ***
; Sæt kopieret tlf i Trio
Trio_opkald(ByRef telefon)
{

    ifWinNotExist, ahk_class AccessBar
    {
        WinActivate, ahk_class Agent Main GUI
        sleep 200
        SendInput, {alt}
        sleep 100
        SendInput, v{Down 5}{enter}
    }
    ControlClick, x360 y17, ahk_class AccessBar
    sleep 800
    WinActivate, ahk_class Addressbook
    ControlClick, Edit2, ahk_class Addressbook
    SendInput, ^a{del}
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
    sleep 100
    SendInput, {F4}
    WinActivate, PLANET
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
    Return
}

; Trio skift mellem pause og klar

trio_pauseklar()
{
    WinActivate, ahk_class AccessBar
    Sleep 200
    SendInput, {F3}
    sleep 400
    SendInput, {F4}
    WinActivate, PLANET

    Return
}

;  ***
;Træk tlf fra Trio indkomne kald
Trio_hent_tlf()
{
    clipboard := ""
    Sendinput !+k
    ClipWait, 3
    Telefon := Clipboard
    rentelefon := Substr(Telefon, 4, 8)
    return rentelefon
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
; Åbn ny mail i outlook. Kræver nymail.lnk i samme mappe som script.
Outlook_nymail()
{
    Run, %A_linefile%\..\lib\nymail.lnk, , ,
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

; *
; færdigskrives
sys_genveje_opslag()
{
    global bruger_genvej
    global genvej_ren := []
    global genvej_navn := databaseget("%A_linefile%\..\db\bruger_ops.tsv", 1, ,1) ; brugerens række i databasen
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
    Gui, Genveje:new
    Gui, Add, Text, , Tekst
    Gui, Add, Text, R1 , % genvej_navn.3 genvej_ren.3
    Gui, Add, Text, R2 , % genvej_navn.3 genvej_ren.3
    Gui, Show, , Genveje

    ; MsgBox,% genvej_navn.4 " - " genvej_ren.4 "`n"  genvej_navn.5 " - " genvej_ren.5

    ; MsgBox, , Genvej, % StrReplace(bruger_genvej.30, "+" , "Shift + ")
    return
}

sys_inputbox_til_fra()
{
    gui, sys:New
    gui, sys:Default
    Gui Font, s9, Segoe UI
    Gui Add, Text, x9 y32 w115 h23 +0x200, P6 - VL Sluttid
    Gui Add, Text, x8 y64 w123 h23 +0x200, P6 - Minutudregner
    Gui Add, DropDownList, vp6_vl_slut x144 y32 w120, Med Inputbox||Uden Inputbox|
    Gui Add, DropDownList, vp6_minut x144 y64 w120, Med Inputbox||Uden Inputbox|

    Gui Show, w307 h332, Window
    Return

    sysGuiEscape:
    sysGuiClose:
    ExitApp

}

;; Misc
; *
; SygehusGUI
; omskrives
l_p6_sygehus_ring_op:
    gui, Sygehus:Default
    Gui,Add,Button,vButton1,&AUH
    Gui,Add,Button,vButton2,RH&G
    Gui,Add,Button,vButton3,&Randers Sygehus
    Gui,Add,Button,vButton4,V&iborg Sygehus
    Gui,Add,Button,vButton5,&Horsens Sygehus
    Gui,Add,Button,vButton6,&Silkeborg Sygehus
    Gui,Show, AutoSize Center , Ring op til sygehus
    knap1:=Func("opkald").Bind("78450000")
    knap2:=Func("opkald").Bind("78430000")
    knap3:=Func("opkald").Bind("78420000")
    knap4:=Func("opkald").Bind("78440000")
    knap5:=Func("opkald").Bind("78425000")
    knap6:=Func("opkald").Bind("78415000")
    GuiControl,+g,Button1,%knap1%
    GuiControl,+g,Button2,%knap2%
    GuiControl,+g,Button3,%knap3%
    GuiControl,+g,Button4,%knap4%
    GuiControl,+g,Button5,%knap5%
    GuiControl,+g,Button6,%knap6%

return
Opkald(p*){
    Gui, Sygehus:Destroy
    telefon := % p.1
    sleep 100
    Trio_opkald(telefon)
    WinActivate, PLANET, , ,
}

SygehusGuiEscape:
    Gui, Destroy
return
SygehusGuiClose:
    gui, Destroy
return

l_p6_central_ring_op:
    gui, Taxa:Default
    Gui,Add,Button,vtaxa1,&Århus Taxa
    Gui,Add,Button,vtaxa2,Århus Taxa Sk&ole
    Gui,Add,Button,vtaxa3,&Dantaxi
    Gui,Add,Button,vtaxa4,Taxa &Midt
    Gui,Add,Button,vtaxa5,&DK Taxi
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

;; Testknap

^+e::
    {
        ControlClick, x360 y17, ahk_class AccessBar
        return
    }
;; HOTKEYS

;; Global
l_quitAHK:
ExitApp
Return

l_p6_aktiver:
    p6_aktiver()
return
;; PLANET
l_p6_hastighed:
    P6_hastighed()
return

l_p6_initialer: ;; Initialer til/fra
    P6_initialer()
Return

l_p6_initialer_skriv: ; skriv initialer og forsæt notering.
    P6_initialer_skriv()
return

l_p6_vis_k_aftale: ;Vis kørselsaftale for aktivt vognløb
    P6_vis_k()
Return

l_p6_ret_vl_tlf: ; +F3 - ret vl-tlf til triopkald
    telefon := Trio_hent_tlf()
    If (WinNotExist, PLANET, , ,)
        MsgBox, , PLANET, P6 er ikke åben.,
    Else
    {
        WinActivate, PLANET
        vl := P6_hent_vl()
        if (telefon = "")
        {
            MsgBox, , Intet indgående telefonnummer, Der er intet indgående telefonnummer, 1
            return
        }
        else
        {
            InputBox, telefon, VL, Skal der bruges et andet telefonnummer end %telefon%?,, 120, 160, X, Y, , Timeout, %telefon%
            if (ErrorLevel = 1 or ErrorLevel = 2)
                return
            MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl%?,
            IfMsgBox, Yes
                P6_ret_tlf_vl(telefon)
            sleep s * 100
            Input, næste, L1 V T4
            if (næste = "n")
            {
                luk_igen:
                    sleep 100
                    MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl% på den efterfølgende dato?,
                    IfMsgBox, Yes
                    {
                        sleep 100
                        P6_tlf_vl_dato_efter(telefon)
                        sleep s * 800
                        Goto, luk_igen
                    }
                    IfMsgBox, no
                    {
                        P6_planvindue()
                    }
                }
            return
        }
        return
    }
    WinActivate, PLANET
    vl := P6_hent_vl()
    if (telefon = "")
    {
        MsgBox, , Intet indgående telefonnummer, Der er intet indgående telefonnummer, 1
        return
    }
    InputBox, telefon, VL, Skal der bruges et andet telefonnummer end %telefon%?,, 120, 160, X, Y, , Timeout, %telefon%
    if (ErrorLevel = 1 or ErrorLevel = 2)
        return
    MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl%?,
    IfMsgBox, Yes
        P6_ret_tlf_vl(telefon)
    sleep s * 100
    Input, næste, L1 V T4
    if (næste = "n")
    {
        p6_vl_ret_tlf_næste:
            sleep 100
            MsgBox, 4, Sikker?, Vil du ændre Vl-tlf til %telefon% på VL %vl% på den efterfølgende dato?,
            IfMsgBox, Yes
            {
                sleep 100
                P6_tlf_vl_dato_efter(telefon)
                sleep s * 800
                Goto, p6_vl_ret_tlf_næste
            }
            IfMsgBox, no
            {
                P6_planvindue()
            }
        }
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

        return
    }
    else
        sleep s * 40
    P6_udfyld_k_og_s(vl)
; MsgBox, , , % vl.2
Return

; ***
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
            return
        }
        if (telefon = "78410222")
        {

            P6_rejsesogvindue()
            sleep s * 40
            SendInput, ^t
            return
        }
        Else
        {
            WinActivate, PLANET
            P6_rejsesog_tlf(telefon)
            return
        }
    }
return

; gå i vl
l_p6_vaelg_vl:
    P6_Planvindue()
    SendInput, !l
return

l_p6_vl_ring_op: ;træk tlf fra aktiv planbillede, ring op i Trio
    P6_planvindue()
    sleep s * 100
    vl_tlf := P6_hent_vl_tlf()
    sleep 200
    Trio_opkald(vl_tlf)
    ; Clipboard = %gemtklip%
    ; gemtklip :=
    sleep 400
    WinActivate, PLANET
    P6_Planvindue()
return

; ***

; ^+F5
l_p6_vm_ring_op: ; træk vm-tlf fra aktivt planbillede, ring op i Trio
    P6_planvindue()
    sleep s * 100
    vm_tlf := P6_hent_vm_tlf()
    sleep 500
    Trio_opkald(vm_tlf)
    sleep 800
    WinActivate, PLANET
Return

; P6 - ring op til kunde markeret i Vl (kræver tlf opsat på kundetilladelse)
l_p6_ring_til_kunde:
    p6_hent_kunde_tlf(telefon)
    sleep s * 200
    if (SubStr(telefon, 1, 3) = "888")
    {
        MsgBox, , Telefon ikke tilknyttet, Kunden har ikke telefon tilknyttet.
        return
    }
    Else
    {
        Trio_opkald(telefon)
        return
    }
return

; #F5
l_p6_vl_luk:
    tid := P6_input_sluttid()
    if tid = 0
        return
    p6_vl_luk(tid)
return

l_p6_udregn_minut:
    tid := P6_udregn_minut()
    tid_tekst := tid.1
    gui, plustid:New,
    gui, plustid:Default
    Gui Font, s9, Segoe UI
    Gui Add, Button, gok x24 y88 w80 h23 +Default, &OK
    Gui Add, Button, gudklip x144 y88 w80 h23, Til &Udklip
    Gui Add, Text, x72 y24 w120 h23 +0x200 +Center, %tid_tekst%

    Gui Show, w260 h125, Resultat
Return

ok:
    {
        gui, cancel
        return
    }
udklip:
    {
        Clipboard := tid.2
        gui, cancel
        return
    }
plustidGuiEscape:
plustidGuiClose:
ExitApp
return

l_p6_alarmer:
    P6_alarmer()
return

l_p6_udråbsalarmer:
    P6_udraabsalarmer()
return
l_p6_tekst_til_chf: ; Send tekst til aktive vognløb
    FormatTime, Time, ,HHmm
    initialer = /mt%A_userName%%time%
    initialer_udentid =/mt%A_userName%
    brugerrække := databasefind("%A_linefile%\..\db\bruger_ops.tsv", A_UserName, ,1)
    bruger := databaseget("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 2)

    KeyWait Alt
    keywait Ctrl
    Input valgt, L1 T5, {esc}, TKF
    if (valgt = "t")
    {
        P6_tekstTilChf() ; tager tekst ("eksempel") som parameter (accepterer variabel)
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
        Gui Add, Text, x16 y86 w120 h23 +0x200, Evt. navn på kunde
        Gui Add, Text, x215 y84 w120 h23 +0x200, Evt. navn på kunde
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
                return
            }
        f_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            tekst := "Jeg har meldt st. " f_stop "`, " . k_navn " `, forgæves og sendt st. " s_stop " `, " k_navn2 . " , i stedet. /" bruger
            ; MsgBox, , , % tekst,
            P6_tekstTilChf("Jeg har meldt st. " f_stop "`, " . k_navn "`, forgæves og sendt st. " s_stop "`, " k_navn2 . ", i stedet. /" bruger)
            sleep 500
            input, tast, M T3, {esc}, p
            if (tast = "p")
            {
                sleep 500
                P6_notat("Ingen kontakt til chf. St. " f_stop " forgæves`, " s_stop " og tekst sendt til chf. " initialer " ")
                gui, cancel
                return
            }
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
        Gui Add, Text, x216 y7 w120 h23 +0x200, Sendt stop
        Gui Add, Edit, vk_navn x14 y106 w120 h21
        Gui Add, Text, x16 y86 w120 h23 +0x200, Evt. navn på kunde
        Gui Add, Text, x215 y84 w120 h23 +0x200, Evt. navn på kunde
        Gui Add, Edit, vk_navn2 x216 y103 w120 h21
        Gui Add, Button, gk_chfok x81 y172 w80 h23 +Default, &OK
        Gui Add, Button, gk_annuller x216 y171 w80 h23, &Annuller

        Gui Show, x812 y22 w381 h220, Send tekst om kvittering til chauffør
        Return

        k_annuller:
        k_chfGuiEscape:
        k_chfGuiClose:
            {
                gui, Cancel
                return
            }
        k_chfok:
            GuiControlGet, f_stop, , ,
            GuiControlGet, s_stop, , ,
            GuiControlGet, k_navn, , ,
            GuiControlGet, k_navn2, , ,
            P6_tekstTilChf("Husk at bede om ny tur ved ankomst. Jeg har kvitteret for " f_stop "`, " . k_navn "`, og sendt st. " s_stop "`, " k_navn2 . "/" bruger)
            sleep 500
            input, tast, M T3, {esc}, ^s
            if (tast = "^s")
            {
                sleep 500
                P6_notat("St. " f_stop " ikke kvitteret ved ankomst`, " s_stop " og tekst sendt til chf. " initialer " ")
                gui, cancel
                return
            }
            Else
                gui, cancel
        return
    }
    return
#IfWinActive ; udelukkende for at resette indentering i auto-formatering

l_outlook_svigt: ; tag skærmprint af P6-vindue og indsæt i ny mail til planet
    gemtklip := ClipboardAll
    sleep 400
    screenshot_aktivt_vindue()
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

;; Trio
l_trio_klar: ;Trio klar
    trio_klar()
Return

l_trio_pause: ;Trio pause
    trio_pause()
Return

l_trio_udenov: ;Trio Midt uden overløb
    trio_udenov()
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
Return

; Telenor accepter indgående kald, søg planet
l_telenor_p6_opslag: ; brug label ist. for hotkey, defineret ovenfor. Bruger.4
    SendInput, % bruger_genvej[3] ; opr telenor-genvej
    sleep 40
    telefon := Trio_hent_tlf()
    sleep 40
    vl := P6_hent_vl_fra_tlf(telefon)
    IfWinNotActive, PLANET
    {
        WinActivate, PLANET
        sleep 500 ; sørger for at vinduet kan nå at Kskifte
    }
    SendInput, {AltUp}
    if (vl != 0) ; giver af og til første VL på listen, når der ikke er ramt et VL. Hvorfor?
    {
        sleep 200
        P6_udfyld_k_og_s(vl)
        Return
    }
    if (telefon = "78410222" OR telefon ="78410224")
    {
        ; MsgBox, ,CPR, CPR, 1
        sleep 200
        P6_rejsesogvindue()
        sleep 200
        SendInput, ^t
        return
    }
    if (telefon = "")
    {
        MsgBox, , , Intet indgående tlf-nr,
        return
    }
    Else
    {
        sleep 200
        P6_rejsesog_tlf(telefon)
        return
    }

l_trio_opkald_markeret: ; Kald det markerede nummer i trio, global. Bruger.12
    clipboard := ""
    SendInput, ^c
    ClipWait, 2, 0
    telefon := clipboard
    sleep 200
    Trio_opkald(telefon)
Return

; Minus på numpad afslutter Trioopkald global (Skal der tilbage til P6?)
l_trio_afslut_opkald:
l_trio_afslut_opkaldB:
    Trio_afslutopkald()
    sleep 200
    WinActivate, PLANET
Return

;; Flexfinder
l_flexf_fra_p6:
    Flexfinder_opslag()
Return

l_flexf_til_p6: ; slår valgte FF-bil op i P6. Bruger.13
    KeyWait, ctrl
    sleep 200
    vl :=Flexfinder_til_p6()
    if (vl = 0)
        return
    Else
    {
        WinActivate PLANET
        sleep s * 200
        P6_udfyld_k_og_s(vl)
        sleep 400 ; skal optimeres
        WinActivate, FlexDanmark FlexFinder, , ,
        Return
    }

;; Outlook
l_outlook_ny_mail: ; opretter ny mail. Bruger.16
    Outlook_nymail()
Return

;; Excel
l_excel_vl_til_P6_A:
l_excel_vl_til_P6_B:
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

^+r:: ; AHK-reload
    SendInput, {CtrlUp}
    Reload
    sleep 2000
Return

^+a::databaseview("%A_linefile%\..\db\bruger_ops.tsv")

