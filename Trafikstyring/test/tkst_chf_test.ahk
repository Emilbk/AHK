#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

; sd

ctrls := Chr(19)
tekst := "test"
MsgBox, 4, Send tekst til chauffør?, Send?, 
IfMsgBox, Yes
    MsgBox ja
IfMsgBox, No
    MsgBox nej
