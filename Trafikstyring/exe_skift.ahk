#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

Gui Font, s9, Segoe UI
Gui Font
Gui Font, s12 Bold, Tahoma
Gui Add, Text, x144 y16 w183 h23 +0x200 +Center, Midlertidig ændring
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Text, x24 y48 w408 h157 +Center  wrap, Indtil videre skal AHK nu startes op på en lidt anden (men grundlæggende den samme) måde. `nI stedet for den fil, I har brugt før, skal I bruge .exe-filen med samme navn i samme mappe (det grønne ikon). `n`nKlik på knappen nedenunder, så opretter du automatisk en ny, korrekt genvej på skrivebordet (eller bare åben mappen). `n`nDet er den genvej, der skal bruges fremadrettet.
Gui Add, Button, x80 y264 w80 h23 gvis, &Vis mappe
Gui Add, Button, x208 y264 w194 h23 ggenvej, &Opret genvej på skrivebordet

Gui Show, w473 h298, Ny genvej
Return

GuiEscape:
GuiClose:
    ExitApp

vis:
{
    Run % "explorer.exe /expand, F:\Flextrafik\Driftscentret\Autohotkey (Emil)\AHK\Trafikstyring"
    ExitApp
}

genvej:
{
    FileCreateShortcut, F:\Flextrafik\Driftscentret\Autohotkey (Emil)\AHK\Trafikstyring\trafikstyring.exe, %A_Desktop%\TrafikstyringAHK.lnk
    MsgBox, 64 , Fuldført , Genvejen er nu oprettet på dit skrivebord som "TrafikstyringAHK"., 
    ExitApp
}