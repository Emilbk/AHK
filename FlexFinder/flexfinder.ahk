﻿#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#SingleInstance, force

; Udfyld variabel aktiv_fil - Vælg mellem:
; gv_0-13, gv_9-36, gv_aktiv_8, tripstol, gv_alle,
; type2, type5, ttj_larve, ttj_hjul,variabel_lift,variabel_type2,
; variabel_ttj,

aktiv_fil := "gv_tripstol"
FileRead, fil, %A_linefile%\..\ff_vl\%aktiv_fil%.txt
; MsgBox, , fil, % fil,
vl := StrSplit(fil, "`n")

; kan der laves en substr-handling på array?

; MsgBox, , Før, % vl.1
; vl := StrSplit(vl, "_")
; for index, element in vl
; vl.element := SubStr("000" . vl.element, -3)
; MsgBox, , Efter, % vl.1
; MsgBox, , , % "Vognløb " index " er " vognløb

; MsgBox, , test, % vl.3

; vl := SubStr("000" . k, -3)

+s::
    ; IfWinExist, FlexDanmark FlexFinder ;insert the window name
    ; WinActivate
    for index, nummer in vl
    {

        MsgBox, , Vognløb, % nummer, 1
        SendInput, ^f
        sleep 300
        SendInput, {del}
        sleep 200
        SendInput, %nummer%
        sleep 200
        ; SendInput, {tab}{tab}{Space}
        ; sleep 200
        PixelSearch, Px, Py, 90, 190, 1062, 621, 0x3296FF, 0, Fast ; oxo0FFFF is the pixel color fould from using the first script, insert yours there
        sleep 200
        click %Px%, %Py%
        sleep 200
        SendInput, ^f
        sleep 200
        SendInput, {del}
        SendInput, {esc}
        sleep 200
    }

MsgBox, , , Vognløb er indtastede,
ExitApp,  
return

^+f::
    {
        aktiv_fil := "gv_tripstol"
        InputBox, aktiv_fil, Vognløbsgruppe, Hvad gruppe skal indlæses?
        FileRead, fil, %A_linefile%\..\ff_vl\%aktiv_fil%.txt
        vl := StrSplit(fil, "`n")
        for index, nummer in vl
        {
            SendInput, ^a{del}
            sleep 100
            SendInput, %nummer%
            KeyWait, Right, D
        }

        ; KeyWait, Right,
    }

return

+z:: ; Control+Z hotkey.
    MouseGetPos, MouseX, MouseY
    PixelGetColor, color, %MouseX%, %MouseY%
    MsgBox The color at the current cursor position is %color%.
return

z::
    IfWinExist, FlexDanmark FlexFinder ;insert the window name
        WinActivate
    PixelSearch, Px, Py, 13, 182, 1203, 230, 0x5B6CF2, 0, Fast
    if ErrorLevel
        MsgBox, That color was not found in the specified region.
    else
        click %Px%, %Py%

return

+Escape::
ExitApp
Return