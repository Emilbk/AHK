Versionsdato = 24/06-2020
Genveje =  Ctrl+Numpad+: Vis version `n Ctrl+Numpad-: Vis genveje `n`n Ctrl+Numpad1: Indsæt "Fak OK" i notat`n Ctrl+Numpad2: Indsæt (IR) i notat `n Ctrl+Numpad3: Ringer til aktive vognløb `n Ctrl+Numpad4: Registrerer svigt `n Ctrl+Numpad5: Tilføjer 3 i bagage på aktive bestilling `n Ctrl+Numpad6: Udruller fast mandag til tirsdag-fredag `n Ctrl+Numpad7: Indsætter "ok bagsæde og faktura" `n Ctrl+Numpad9: Meld tur forgæves `n`n Pause: Stopper scriptet
#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
Pause::Pause
^NumpadAdd::
MsgBox, 0, Midttrafik AHK, Versionsnummer %Version%`n`n Opdateret senest %Versionsdato%
return
^NumpadSub::
MsgBox, 0, Aktive genveje, %Genveje%
return
!Numpad0::
SetTitleMatchMode 2
ifwinexist, Trio Agent
{
winactivate, Trio Agent
Sleep 500
Send {Alt}
Sleep 200
Send {Enter}
Sleep 200
Send o
Sleep 100
Send 8
}
Reload
ExitApp
Return
!Numpad1::
SetTitleMatchMode 2
ifwinexist, Trio Agent
{
winactivate, Trio Agent
Sleep 500
Send {Alt}
Sleep 200
Send {Enter}
Sleep 200
Send o
Sleep 100
Send 3
}
Reload
ExitApp
Return
!Numpad2::
SetTitleMatchMode 2
ifwinexist, Trio Agent
{
winactivate, Trio Agent
Sleep 500
Send {Alt}
Sleep 200
Send {Enter}
Sleep 200
Send o
Sleep 100
Send 4
}
Reload
ExitApp
Return
!Numpad3::
SetTitleMatchMode 2
ifwinexist, Trio Agent
{
winactivate, Trio Agent
Sleep 500
Send {Alt}
Sleep 200
Send {Enter}
Sleep 200
Send o
Sleep 100
Send {up}
Sleep 100
Send {enter}
}
Reload
ExitApp
Return
^Numpad1::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
WinActivate, PLANET version 6
Send {ALT}{RIGHT}{DOWN}
Sleep 500
Send {ENTER}
Sleep 1000
#SingleInstance
SetTimer, ChangeButtonNames, 50
MsgBox, 4, Rejsetype, Vælg rejsetype:
IfMsgBox, YES
{
Sleep 1000
Send ^t
Sleep 100
Send {text}mtflextur
Send {TAB 2}
Send ^v
Send {ENTER}
Sleep 1000
Send {Enter}
Send !n{TAB 2}
Send {text}Flextur ring
Send {Space}
Send ^v
Sleep 100
Send !r
}
else
{
Sleep 1000
Send ^t
Sleep 100
Send {text}mtflexbus
Send {TAB 2}
Send ^v
Send {ENTER}
Sleep 250
Send {Enter}
Send !n{TAB 2}
Send {text}Flexbus ring
Send {Space}
Send ^v
Sleep 100
Send !r
}
ChangeButtonNames:
IfWinNotExist, Rejsetype
return
SetTimer, ChangeButtonNames, On
WinActivate
ControlSetText, Button1, &Flextur
ControlSetText, Button2, &Flexbus
return
}
Reload
ExitApp
Return
^Numpad2::
Send ^æ
sleep 1000
Send !n
sleep 500
Send {tab 2}
Sleep 500
Send {right}{space}
Send {text}(IR)
sleep 500
send ^p
Sleep 500
send !o
Reload
ExitApp
Return
^Numpad3::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
WinActivate, PLANET version 6
Send ^{F12}
Sleep 1000
Send ^æ
Send {Enter}{Enter}
Send !ø
Send {TAB 2}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Send ^a
Sleep 1000
Send ^{TAB}
Sleep 500
Ifwinexist, Trio Agent
{
Winactivate, Trio Agent
Sleep 350
Send {F5}
Sleep 500
Send {TAB 2}
Send ^v
Send {shift down}{Enter down}
Sleep 10
Send {shift up}{enter up}
}
}
Reload
ExitApp
Return
^Numpad4::
SetTitleMatchMode 2
ifwinexist PLANET version 6
{
winactivate PLANET version
send !l
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send {shift down}{F10}
Sleep 10
Send {Shift up}
Send o
}
clip := Clipboard
if (clip <= 6000 or clip >=9999)
{
ifwinnotexist, Driftssvigt
{
run F:\AHK\Driftssvigt.docx
}
ifwinexist, Driftssvigt
Winwaitactive, Driftssvigt
Winactivate, Driftssvigt
Sleep 15000
Send ^a
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Winkill, Driftssvigt
ifwinexist, Outlook
{
winactivate, Outlook
Send ^n
Sleep 500
Send {text}planet@midttrafik.dk
Send {tab 5}
Send ^v
Sleep 1000
Send {shift down}{tab down}
Sleep 10
Send {shift up}{tab up}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
Ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send {F5}
Send {shift down}{F10 down}
Sleep 10
Send {shift up}{F10 up}
Sleep 1000
Send o
Sleep 500
}
Ifwinexist, Ikke-navngivet
{
Winactivate, Ikke-navngivet
Send {text} svigt
Send {space}
Send ^v
Send {space} %A_DD%-%A_MM%-%A_YYYY%
Send {tab}{up 13}{tab}
Send ^v
Send {up 2} %A_Hour%:%A_Min%
Send {up} %A_DD%-%A_MM%-%A_YYYY%
Sleep 500
}
Ifwinexist, PLANET version 6
{
Winactivate, PLANET version 6
Sleep 500
Send {F3}
Sleep 2000
Send ^æ
Sleep 100
Send !a
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Send ^a
}
Ifwinexist, svigt
{
Winactivate, svigt
Send {tab 2}{down}
Send ^v
}
Ifwinexist, PLANET version 6
{
Winactivate, PLANET version 6
Send ^{tab}
Sleep 2500
clipboard:=""
While clipboard
Sleep 10
Send !{PrintScreen}
}
ifwinexist, svigt
{
WinActivate, svigt
Send {down 6}
Send {enter 3}
Send ^v
Sleep 100
Send  ^{Home}
}
}
MsgBox, 0, Svigtark, Svigtark oprettet. Manuel viderebehandling
}
else
{
ifwinnotexist, Driftssvigt
{
run F:\AHK\Driftssvigt.docx
}
ifwinexist, Driftssvigt
Winwaitactive, Driftssvigt
Winactivate, Driftssvigt
Sleep 15000
Send ^a
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Winkill, Driftssvigt
ifwinexist, Outlook
{
winactivate, Outlook
Send ^n
Sleep 500
Send {text}planet@midttrafik.dk
Send {tab 5}
Send ^v
Sleep 1000
Send {shift down}{tab down}
Sleep 10
Send {shift up}{tab up}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
Ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send {F5}
Send {shift down}{F10 down}
Sleep 10
Send {shift up}{F10 up}
Sleep 1000
Send o
Sleep 500
}
Ifwinexist, Ikke-navngivet
{
Winactivate, Ikke-navngivet
Send {text} svigt
Send {space}
Send ^v
Send {space} %A_DD%-%A_MM%-%A_YYYY%
Send {tab}{up 13}{tab}
Send ^v
Send {up 2} %A_Hour%:%A_Min%
Send {up} %A_DD%-%A_MM%-%A_YYYY%
Sleep 500
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send ^{F12}
Sleep 500
Send ^æ
Send {tab 2}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Send ^a
Sleep 500
}
ifwinexist, svigt
{
winactivate, svigt
Send {down 2}
Send ^v
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send {alt}
Sleep 1000
send {tab 2}
Sleep 1000
Send {down}
Send 500
send d
sleep 1500
send ^v
Send {Enter}
Sleep 100
clipboard:=""
While clipboard
Sleep 10
Send !{PrintScreen}
}
ifwinexist, svigt
{
winactivate, svigt
Send {down 5}
send ^v
Send {left}{up}{enter}
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send ^{tab}
Send ^{F10}
Sleep 2500
clipboard:=""
While clipboard
Sleep 10
Send !{PrintScreen}
}
ifwinexist, svigt
{
winactivate, svigt
Send ^v
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send ^{F12}
Sleep 500
Send ^æ
Send {tab 2}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Send ^a
Send {alt}
Sleep 1000
Send {tab 2}
Sleep 1000
Send {down}
Sleep 500
Send v
Sleep 1500
Send {tab 3}
Sleep 500
Send 2359
Send !g
Send ^v
Send {enter}
Sleep 1000
clipboard:=""
While clipboard
Sleep 10
Send !{PrintScreen}
Send ^{F4}
}
ifwinexist, svigt
{
winactivate, svigt
Send ^v
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send ^{F10}
Clipboard:=""
Send !k
Send +{F10}
Sleep 100
Send {down 3}
Sleep 100
Send {enter}
}
clip := Clipboard
if clip =
{
ifwinexist, svigt
{
winactivate, svigt
Send ^{home}
Sleep 500
Send {down 5}{tab 3}
Send ^v
}
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send {F3}
Sleep 1000
Send ^æ
Sleep 500
Send !a
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
Send ^a
}
ifwinexist, svigt
{
winactivate, svigt
Send {down}
Sleep 10
Send ^v
}
}
else
{
ifwinexist, svigt
{
winactivate, svigt
Send ^{home}
Sleep 2000
}
}
}
MsgBox, 0, Svigtark, Svigtark oprettet. Manuel viderebehandling
}
}
Reload
ExitApp
Return
^Numpad5::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
WinActivate, PLANET version 6
Send !o
Sleep 100
Send {shift down}{tab 4}
sleep 10
Send {Shift up}
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send ^c
Sleep 100
}
If clipboard contains Flextur,Flexbus
{
Send ^æ
sleep 1000
Send !ø
Send !a
Send {shift down}{TAB 2}
Sleep 10
Send {shift up}
Send {text}3
}
Else
{
Send ^æ
sleep 1000
Send !ø
Send !a
Send {shift down}{TAB 2}
Sleep 10
Send {shift up}
Send {text}3
}
Send !n
Send {tab 2}
Send {right}
Send {text} + kuffert
Send ^p
Sleep 1000
Send !o
}
Reload
ExitApp
Return
^Numpad6::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Send ^k
Sleep 500
Send !d
Send {text}ti
Send ^p
Sleep 1000
Send !o
Sleep 10000
Send !o
Sleep 1000
Send ^k
Sleep 500
Send !d
Send {text}on
Send ^p
Sleep 1000
Send !o
Sleep 10000
Send !o
Sleep 1000
Send ^k
Sleep 500
Send !d
Send {text}to
Send ^p
Sleep 1000
Send !o
Sleep 10000
Send !o
Sleep 1000
Send ^k
Sleep 500
Send !d
Send {text}fr
Send ^p
Sleep 1000
Send !o
Sleep 10000
Send !o
Sleep 500
}
Reload
ExitApp
Return
^Numpad7::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Sleep 100
Send !n
Send {tab 2}
Sleep 100
Send {right}
Send {text} bagsæde OK
Sleep 100
}
Reload
ExitApp
Return
^Numpad8::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Sleep 100
Send !n
Send {tab 2}
Sleep 100
Send {right}
Send {text} Fak OK
Sleep 100
}
Reload
ExitApp
Return
^Numpad9::
SetTitleMatchMode 2
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
Sleep 250
Send !k
Sleep 250
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send {Appskey}
Sleep 200
Send o
}
Send {Alt}
Send {right 2}
Send {ENTER}
Sleep 250
Send t
Sleep 250
Send {Appskey}
Sleep 100
Send æ
Sleep 100
Send {Alt}
Send {right 2}
Send {ENTER}
Sleep 250
Send p
Sleep 250
Send {tab}
Sleep 250
clipboard:=""
While clipboard
Sleep 10
While !clipboard
{
Send {Appskey}
Sleep 200
Send o
}
Send {Alt}
Send {right 2}
Send {ENTER}
Sleep 250
Send t
Sleep 250
Send {tab}
Send {Appskey}
Sleep 100
Send æ
Sleep 100
Send {tab}
Send {text}Husk at trykke ny ordre ved ankomst Mvh. Flextrafik.
Sleep 250
Send ^S
}
Reload
ExitApp
Return
^Numpad0::
oList:="Bestilling|Tidspunkt|Adresse|Vogntype|Kunde"
List_1:="b1||b2|b3"
List_2:="t1||t2"
List_3:="s1||s2|s3|s4|s5|s6|s7"
List_4:="v1||v2|v3|v4|v5"
List_5:="k1||k2|k3|k4|k5|k6|k7|k8|k9|k10"
Gui, Color, Teal
Gui, Add, Text,, Meld tur forgæves
Gui, Add,DDL,x10 y+5 w200 r10 AltSubmit vList_Selector gChange_List,% oList
Gui, Add,DDL,x10 y+5 w200 r10 vSelected_List
Gui, Add, Text,, Begrundelse:
Gui, Add, Edit, w200 r4 vBegrundelse
Gui, Add, Button, Default w80 gsubmitBtn, OK
Gui, Add, Button, Default w80 gcancelBtn x+5 -Tabstop, Afbryd
Gui, Show, w250 h200, Forgævesvindue
Return
Change_List:
Gui,1:Submit,NoHide
GuiControl,1:,Selected_List,% "|" List_%List_Selector%
return
submitBtn:
ifwinexist, PLANET version 6
{
winactivate, PLANET version 6
}
Gui, Submit, NoHide
Gui, Destroy
Sleep 1000
Send ^f
Sleep 1000
Send !o
Sleep 1000
sendinput, %Begrundelse%
Send {tab}
sendinput, %Selected_List%
send !o
Sleep 100
Reload
ExitApp
Return
cancelBtn:
Reload
ExitApp
Return
GuiClose:
Reload
ExitApp
Return
WinGetClass, Clipboard, A
return
setKeyDelay, 50, 50
setMouseDelay, 50
$~MButton::
Send, ^+{M}
while (getKeyState("MButton", "P"))
{
sleep, 100
}
Send, ^+{M}
returnPAD