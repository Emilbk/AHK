﻿#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, %A_linefile%\..\lib\AHKDb\ahkdb.ahk
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
;hjælp GUI

Gui Font, s9, Segoe UI
Gui Color, 0xC0C0C0
Gui Add, StatusBar,, Status Bar
Gui Add, Tab3, x0 y0 w748 h642 0x54010240, Hej|Genvejsoversigt 1|Genvejsoversigt 2|Opsætning|Hjælp|Misc
Gui Tab, Opsætning
Gui Add, Text, x8 y32 w123 h23 +0x200, P6 - Hastighed
Gui Add, Text, x285 y32 h23 +0x200, Tilpas efter P6-langsomhed. 1 = hurtigst. Skal bruge punktum (eks. 1.2).
Gui Add, Text, x9 y64 w115 h23 +0x200, P6 - VL Sluttid
Gui Add, Text, x285 y64 h23 +0x200, Vælg om der skal bruges en popup, der kan skrives i til funktionen Luk Vognløb.
Gui Add, Text, x8 y96 w123 h23 +0x200, P6 - Minutudregner
Gui Add, Text, X285 y96 h23 +0x200, Vælg om der skal bruges en popup, der kan skrives i til funktionen Minutudregner.
Gui Add, edit, vp6_hastighed_ops x144 y32 w120 , %p6_hastighed_ops%
Gui Add, DropDownList, vp6_vl_slut x144 y64 w120 Choose%vl_default%, Med Inputbox|Uden Inputbox|
Gui Add, DropDownList, vp6_minut x144 y96 w120 Choose%min_default%, Med Inputbox|Uden Inputbox|
Gui Add, Button, gsysok, &OK
Gui Tab, Genvejsoversigt
Gui Font
Gui Font, s12 Bold
Gui Add, Text, x0 y0 w748 h642 +0x200, Generelt
Gui Font
Gui Font, s9, Segoe UI
Gui Tab, Genvejsoversigt 1
Gui Font
Gui Font, s14 Bold q4, Segoe UI
Gui Add, Text, x16 y32 w120 h23 +0x200, Generelt
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Text, x16 y56 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x272 y56 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x16 y80 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x272 y80 w227 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y56 w198 h2 +0x10
Gui Tab, Genvejsoversigt 2
Gui Font
Gui Font, s14 Bold q4, Segoe UI
Gui Add, Text, x16 y32 w120 h23 +0x200, Planet
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Text, x8 y56 w198 h2 +0x10
Gui Add, Text, x8 y64 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y88 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y112 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y136 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y160 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y184 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y208 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y232 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y256 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y280 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y304 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y328 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y352 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y376 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y400 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x8 y424 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x248 y64 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y88 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y112 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y136 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y160 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y184 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y208 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y232 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y256 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y280 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y304 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y328 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y352 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y376 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y400 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y424 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y448 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y472 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y496 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y520 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y544 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y568 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x8 y592 w260 h23 +0x200, % genvej_ren.3
Gui Add, Text, x344 y64 w227 h23 +0x200,% genvej_navn.3
Gui Add, Text, x344 y88 w227 h23 +0x200,% genvej_navn.3
Gui Add, Text, x344 y112 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y136 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y160 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y184 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y208 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y232 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y256 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y280 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y304 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y328 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y352 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y376 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y400 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y424 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y472 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y496 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y520 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y544 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y568 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y592 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x344 y448 w227 h23 +0x200, % genvej_navn.3
Gui Add, Text, x248 y448 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y472 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y496 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y520 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y544 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y568 w97 h23 +0x200, % genvej_ren.3
Gui Add, Text, x248 y592 w97 h23 +0x200, % genvej_ren.3
Gui Tab, Hej
Gui Add, Text, x12 y105 w464 h0 +0x10
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
if (p6_vl_slut ="Med Inputbox")
{
    p6_vl_ops = 1
    gui, cancel
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
}
if (p6_vl_slut ="Uden Inputbox")
{
    p6_vl_ops = 0
    gui, cancel
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 42, p6_vl_ops)
}
if (p6_minut ="Med Inputbox")
{
    p6_minut_ops = 1
    gui, cancel
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
}
if (p6_minut ="Uden Inputbox")
{
    p6_minut_ops = 0
    gui, cancel
    databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 44, p6_minut_ops)
}
databasemodifycell("%A_linefile%\..\db\bruger_ops.tsv", brugerrække.1, 41, p6_hastighed_ops)

GuiEscape:
genvejGuiClose:
gui, cancel

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