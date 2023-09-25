#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

+^t::

    Gui, sys:New
    Gui, sys:default
    Gui Font, s9, Segoe UI
    Gui Add, Text, x9 y32 w115 h23 +0x200, P6 - VL Sluttid
    Gui Add, Text, x8 y64 w123 h23 +0x200, P6 - Minutudregner
    Gui Add, DropDownList, vp6_vl_slut x144 y32 w120, Med Inputbox||Uden Inputbox|
    Gui Add, DropDownList, vp6_minut x144 y64 w120, Med Inputbox||Uden Inputbox|
    Gui Add, Button, gsysok, &OK

    Gui Show, w307 h332, Window
Return

sysok:
    GuiControlGet, p6_vl_slut
    GuiControlGet, p6_minut
    MsgBox, , , % p6_vl_slut,

    if (p6_vl_slut ="Med Inputbox")
    {
        p6_vl_ops = "1"
        gui, cancel
    }
    if (p6_vl_slut ="Uden Inputbox")
    {
        p6_vl_ops = "0"
        gui, cancel
    }
    if (p6_minut ="Med Inputbox")
    {
        p6_minut_ops = "1"
        gui, cancel
    }
    if (p6_minut ="Uden Inputbox")
    {
        p6_minut_ops = "0"
        gui, cancel
    }
    MsgBox, , , % p6_minut_ops, 
return

GuiEscape:
GuiClose:
ExitApp