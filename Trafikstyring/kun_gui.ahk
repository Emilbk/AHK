#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallKeybdHook
#InstallMouseHook
;FileEncoding UTF-8
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetTitleMatchMode, 1 ; matcher så længe et ord er der
#SingleInstance, force

Gui skift: Font, s9, Segoe UI
Gui skift: Font
Gui skift: Font, s12 Bold, Tahoma
Gui skift: Add, Text, x144 y16 w183 h23 +0x200 +Center, Midlertidig ændring
Gui skift: Font
Gui skift: Font, s9, Segoe UI
Gui skift: Add, Text, x24 y48 w408 h157 +Center  wrap, Indtil videre skal AHK nu startes op på en lidt anden (men grundlæggende den samme) måde. `nI stedet for den fil, I har brugt før, skal I bruge .exe-filen med samme navn i samme mappe (det grønne ikon). `n`nKlik på knappen nedenunder, så opretter du automatisk en ny, korrekt genvej på skrivebordet (eller bare åben mappen). `n`nDet er den genvej, der skal bruges fremadrettet.
Gui skift: Add, Button, x80 y264 w80 h23 gvis, &Vis mappe
Gui skift: Add, Button, x208 y264 w194 h23 ggenvej, &Opret genvej på skrivebordet

Gui skift: Show, w473 h298, Ny genvej
Return

skiftEscape:
skiftclose:
    ExitApp

vis:
{
    Run % "explorer.exe /expand, F:\Flextrafik\Driftscentret\Autohotkey (Emil)\AHK\Trafikstyring"
    return
}

genvej:
{
    FileCreateShortcut, F:\Flextrafik\Driftscentret\Autohotkey (Emil)\AHK\Trafikstyring\trafikstyringAHK.exe, %A_Desktop%\TrafikstyringAHK.lnk
    MsgBox, 64 , Fuldført , Genvejen er nu oprettet på dit skrivebord som "TrafikstyringAHK"., 
    ExitApp
}


