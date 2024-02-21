#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

hent_vl()
{
    SendInput, !l
    clipboard := ""
    SendInput, {AppsKey}c
    ClipWait, 1
    vl := clipboard
    return vl
}
hent_dato()
{
    clipboard := ""
    SendInput, ^c
    ClipWait, 1
    dato := clipboard
    return dato
}
vl_billede(vl, dato)
{
    SendInput, !tl
    sleep 200
    SendInput, %vl% {tab} %dato% {enter}
    return

}
#IfWinActive, PLANET
^F12::
{
    KeyWait, Ctrl
vl := hent_vl()
SendInput, {tab}
sleep 150
dato := hent_dato()
vl_billede(vl, dato)
return
}


+esc::
{
    ExitApp,
}
#IfWinActive