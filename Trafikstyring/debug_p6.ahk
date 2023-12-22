SetTitleMatchMode, 1
; #IfWinActive, PLANET
F11::
{
    VSC := WinExist("ahk_exe Code.exe")
    pl:= WinExist("ahk_exe WIFICA32.exe")
    ControlFocus,, % "ahk_id" VSC
    ControlSend,, {F11}, % "ahk_id" VSC
    WinActivate, pl
    return
}
F10::
{
    VSC := WinExist("ahk_exe Code.exe")
    pl:= WinExist("ahk_exe WIFICA32.exe")
    ControlFocus,, % "ahk_id" VSC
    ControlSend,, {F10}, % "ahk_id" VSC
    WinActivate, pl
    return
}
F5::
{
    VSC := WinExist("ahk_exe Code.exe")
    pl:= WinExist("ahk_exe WIFICA32.exe")
    ControlFocus,, % "ahk_id" VSC
    ControlSend,, {F5}, % "ahk_id" VSC
    WinActivate, pl
    return
}
; #IfWinActive