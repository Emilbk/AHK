#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
FormatTime, dato, , d/MM
FormatTime, tidspunkt, , HH:mm

vl := "31232"
; MsgBox, , , % tidspunkt, 

Gui +OwnDialogs
Gui Font, w600
Gui Add, Text, x16 y0 w120 h23 +0x200, Vognløbs&nummer
Gui Font
Gui Add, Edit, vVL x16 y24 w120 h21, %vl%
Gui Font, s9, Segoe UI
Gui Font, w600
Gui Add, Text, x161 y0 w118 h25 +0x200, &Lukket?
Gui Font
Gui Font, s9, Segoe UI
Gui Add, CheckBox, vlukket x160 y24 w39 h23, &Ja
Gui Add, Edit, vtid x200 y24 w79 h21, Klokken
Gui Add, CheckBox, vhelt x160 y48 w120 h23, Lukket &helt
Gui Font, s9, Segoe UI
Gui Font, w600
Gui Add, Text, x304 y0 w120 h23 +0x200, Garanti eller variabel
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Radio, x304 y24 w120 h16, &Garanti
Gui Add, Radio, x304 y40 w120 h32, G&arantivognløb i variabel tid
Gui Add, Radio, vtype x304 y72 w120 h23, &Variabel
Gui Font, w600
Gui Add, Text, x16 y48 w120 h23 +0x200, &Årsag
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Edit, vårsag x16 y72 w120 h21
Gui Font, w600
Gui Add, Text, x8 y96 h23 +0x200, &Beskrivelse
Gui Font
Gui Font, s9, Segoe UI
Gui Add, Edit, vbeskrivelse x8 y120 w410 h126
Gui Add, Button, gok x176 y256 w80 h23, &OK

Gui Show, w448 h297, Svigt
ControlFocus, Button1, Svigt
SendInput, ^a
^Backspace::Send +^{Left}{Backspace}
Return

ok:
gui, submit
; MsgBox, , , % beskrivelse
; GuiControlGet, tid
; GuiControlGet, årsag
; GuiControlGet, beskrivelse
; GuiControlGet, lukket
; GuiControlGet, helt
; GuiControlGet, vl
beskrivelse := StrReplace(beskrivelse, "`n", " ")
if (lukket = 1 and StrLen(tid) != 4)
    {
    MsgBox, , Klokkeslæt skal være firecifret, Klokkeslæt skal være firecifret (intet kolon).
    Gui Show, w448 h297, Svigt
    return
    }
if (StrLen(tid) = 4)
    tid := SubStr(tid, 1, 2) ":" SubStr(tid, 3, 2)
if (lukket = 1 and tid = "Klokken")
    {
    MsgBox, , Mangler tidspunkt, Husk at udfylde klokkeslæt for lukning af VL.
    Gui Show, w448 h297, Svigt
    return
    }
if (type = 0)
    {
    MsgBox, , Mangler VL-type, Husk at krydse af i typen af VL.
    Gui Show, w448 h297, Svigt
    return
    }
if type = 1
    vl_type := "GV"
if type = 2
    vl_type := "(Variabel tid)"
if type = 3
    
MsgBox, , beskrivelse , % beskrivelse 
; MsgBox, , type , % type 
; MsgBox, , tid , % tid 
; MsgBox, , årsag , % årsag 
; MsgBox, , helt , % helt 
; MsgBox, , vl , % dato
if (type = 1 and lukket = 1 and helt = 0 and årsag != "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
    MsgBox, , 1 , % emnefelt, 
    gui, destroy
    }
if (type = 1 and lukket = 1 and helt = 0 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type " - lukket kl. " tid " d. " dato
    MsgBox, , 2, % emnefelt, 
    gui, destroy
    }
if (type = 1 and lukket = 0 and helt = 0 and årsag != "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": " årsag " - d. " dato
    MsgBox, , 3, % emnefelt, 
    gui, destroy
    }
if (type = 1 and lukket = 0 and helt = 0 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type " d. " dato
    MsgBox, , 4, % emnefelt, 
    gui, destroy
    }
if (type = 1 and helt = 1 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": ikke startet op d. " dato
    MsgBox, , 5, % emnefelt, 
    gui, destroy
    }
if (type = 1 and helt = 1 and årsag != "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": " årsag " - ikke startet op d. " dato
    MsgBox, , 5.1, % emnefelt, 
    gui, destroy
    }
if (type = 2 and lukket = 0 and årsag !="")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": " årsag " - " dato
    MsgBox, , 6, % emnefelt, 
    gui, destroy
    }
if (type = 2 and lukket = 0 and helt = 0 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type " " årsag "d. " dato
    MsgBox, , 7, % emnefelt, 
    gui, destroy
    }
if (type = 2 and lukket = 0 and helt = 1 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": ikke startet op d. " dato
    MsgBox, , 7.1, % emnefelt, 
    gui, destroy
    }
if (type = 2 and lukket = 1 and årsag != "")
    {
    emnefelt := "Svigt VL" vl " " vl_type ": " årsag " - lukket kl. " tid " d. " dato
    MsgBox, , 8, % emnefelt, 
    gui, destroy
    }
if (type = 2 and lukket = 1 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " " vl_type " - lukket kl. " tid " d. " dato
    MsgBox, , 9, % emnefelt, 
    gui, destroy
    }
if (type = 3 and årsag != "")
    {
    emnefelt := "Svigt VL" vl ": " årsag " - d. " dato
    MsgBox, , 10, % emnefelt, 
    gui, destroy
    }
if (type = 3 and årsag = "")
    {
    emnefelt := "Svigt VL" vl " d. " dato
    MsgBox, , 11, % emnefelt, 
    gui, destroy
    }

MsgBox, , , % emnefelt "`n" beskrivelse


Return

GuiEscape:
GuiClose:
    ExitApp
