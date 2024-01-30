#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

vl_ny := []
FileRead, vl_data, db/vl_data.txt
FileRead, paragraf_data, db/paragraf_data.txt
vl_data := StrSplit(vl_data, "`n")
for i,e in vl_data
    {
    vl_ny[i] := StrSplit(e, "`t", "`r")
    }
vl_data := vl_ny
paragraf_ny := []
paragraf_data := StrReplace(paragraf_data, "`r", "")
paragraf_data := StrSplit(paragraf_data, "`n")
for i,e in paragraf_data
    {
    paragraf_ny[i] := StrSplit(e, "`t")
    }


GUIfokus()
{
ControlGetFocus, GUIfokus
return GUIfokus
}

+LButton::
{
if (WinActive("Planet - Indbakke - Planet - Outlook"))
    {
        SendInput, {RButton}
        sleep 50
        SendInput, h
        sleep 50
        SendInput, {enter}
        sleep 50
        SendInput, {Up}
        WinActivate, "Svigt FG8-FV8.xlsx - Excel"
        return
    }
if (WinActive("Svigt FG8-FV8.xlsx - Excel"))
    {
        SendInput, {click2}
        sleep 100
        SendInput, ^v{tab}
        return
    }
else
    {
        SendInput, +{click}
        return
    }
}
^q::
{
if (WinActive("Planet - Indbakke - Planet - Outlook"))
    {
        SendInput, {AppsKey}
        sleep 50
        SendInput, h
        sleep 50
        SendInput, {enter}
        sleep 50
        SendInput, {Up}
        WinActivate, "Svigt FG8-FV8.xlsx - Excel"
        return
    }
if (WinActive("Svigt FG8-FV8.xlsx - Excel"))
    {
        SendInput, {F2}
        SendInput, ^v{tab}
        return
    }
else
    {
        SendInput, +{click}
    }
}
!q::
{
if (WinActive("Svigt FG8-FV8.xlsx - Excel"))
    {
        WinActivate Planet - Indbakke - Planet - Outlook
        ControlFocus, OutlookGrid1, Planet - Indbakke - Planet - Outlook
        sleep 200
        SendInput, {AppsKey}
        sleep 50
        SendInput, h
        sleep 50
        SendInput, {enter}
        sleep 500
        SendInput, {Up}
        WinActivate, Svigt FG8-FV8.xlsx - Excel
        return
    }
if (WinActive("Planet - Indbakke - Planet - Outlook"))
    {
        clipboard :=
        SendInput, ^c
        ClipWait, 1, 
        sleep 50
        WinActivate, Svigt FG8-FV8.xlsx - Excel
        SendInput, {F2}
        SendInput, ^v{tab}  
        return
    }
}


vm_opslag(vl_data)
{
    InputBox, vl
    for i, e in vl_data
        for i2, e2 in e
            if (e2 = vl)
            {
               vm := vl_data[i][2]
            }
return vm
}
paragraf_opslag(paragraf_data)
{
    InputBox, paragraf
    for i, e in paragraf_data
        for i2, e2 in e
            if (e2 = paragraf)
            {
               paragraf := paragraf_data[i][2]
            }
return vm
}
!z::
{
    vm := paragraf_opslag(paragraf_data)   
    MsgBox, , , % vm, 
}
;; GUI

Gui vl_bod: Font, s9, Segoe UI
Gui vl_bod: Add, Edit, x62 y27 w120 h21
Gui vl_bod: Font
Gui vl_bod: Font, Bold
Gui vl_bod: Add, Text, x18 y22 w35 h23 +0x200, VL
Gui vl_bod: Font
Gui vl_bod: Font, s9, Segoe UI
Gui vl_bod: Add, MonthCal, x20 y58 w164 h160
Gui vl_bod: Add, DropDownList, x23 y268 w414, DropDownList||
Gui vl_bod: Font
Gui vl_bod: Font, Bold
Gui vl_bod: Add, Text, x24 y243 w120 h23 +0x200, Paragraf
Gui vl_bod: Font
Gui vl_bod: Font, s9, Segoe UI
Gui vl_bod: Add, Edit, x25 y328 w568 h84
Gui vl_bod: Font
Gui vl_bod: Font, Bold
Gui vl_bod: Add, Text, x25 y300 w260 h23 +0x200, "Kvalitetetsbristen bestod i, at..."
Gui vl_bod: Font
Gui vl_bod: Font, s9, Segoe UI
Gui vl_bod: Add, Text, x285 y10 w120 h23 +0x200, VM
Gui vl_bod: Add, Text, x285 y54 w120 h23 +0x200, Kontaktinfo

Gui vl_bod: Show, w620 h442, Window
Return

GuiEscape:
GuiClose:
    ExitApp

