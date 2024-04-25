#singleinstance, force
#noenv
sendmode, input
setbatchlines, -1
setworkingdir, %a_scriptdir%
outlook := ComObjCreate("Outlook.application")

; outlookMail := outlook.MailItem

; outlookMail := outlook.ActiveExplorer.Selection.Item(1)
; mailbody := outlookMail.body
; ; MsgBox, , , % outlookmail.Body
; ; MsgBox, , , % outlookmail.Subject
; if InStr(outlookMail.subject, "Driftsvigt")
; {
;     mailbody := StrSplit(outlookMail.body,"`r`n")
;     MsgBox, , Er FDSvigt, % mailbody[9]
; }
; else
;     MsgBox, , Er ikke FDSvigt, % outlookmail.Body

; genåbnet indenfor 2 tim
#IfWinActive, Svigt FG8-FV8.xlsx - Excel
!g::
{
    KeyWait, alt
    KeyWait, g
    SendInput, {tab 2}
    sleep 100
    SendInput, {AltDown}{Down}{AltUp}
    sleep 200
    SendInput, {down 3}
    sleep 500
    SendInput, {enter}
    sleep 100
    SendInput, {ShiftDown}{tab 4}{ShiftUp}
    sleep 100
    SendInput, fg - vognløb lukket/
    sleep 40
    SendInput, {return}
    return
}
#IfWinActive, Svigt FG8-FV8.xlsx - Excel
    !q::
        {
            winactivate Planet - Svigt til behandling - Planet - Outlook
            if (!WinExist("Planet - Svigt til behandling - Planet - Outlook"))
                {
                    MsgBox, , , Svigtmappe ikke åben
                    return
                }
            sleep 150
            controlfocus, outlookgrid1, Planet - Svigt til behandling - Planet - Outlook
            sleep 150
            sendinput, {appskey}
            ; ControlClick, Outlookgrid1, Planet - Svigt til behandling - Planet - Outlook, , Right, 1
            ; ControlSend, Outlookgrid1, {AppsKey}, Planet - Svigt til behandling - Planet - Outlook
            ; return
            sleep 240
            sendinput, h
            sleep 90
            sendinput, {enter}
            ; ; sleep 500
            ; ; sendinput, {up}
            ; ; sleep 500
            ; ; controlfocus, _WwG1 , Planet - Svigt til behandling - Planet - Outlook
            ; ; sleep 500
            ; ; SendInput, +{down}
            winactivate, Svigt FG8-FV8.xlsx - Excel
            return
        }
#IfWinActive, Svigt FG8-FV8.xlsx - Excel
    !w::
        {
            ; tjek fdsvigt
            mailbody := Fdsvigt(outlook)
            if mailbody[mailbody.MaxIndex()] = "FD"
            {
                for i, e in mailbody
                    {
                        if InStr(mailbody[i], "Beskrivelse af driftsvigt")
                            {
                                Clipboard := substr(mailbody[i], 28)
                                break
                            }
                    }
                
                sleep 150
                SendInput, {tab}
                sendinput, {f2} ^v
                sleep 100
                SendInput, {tab}
                sleep 40
                SendInput, mtebk{tab}
                sleep 40
                SendInput, !{down}
                return
            }
            if mailbody[mailbody.MaxIndex()] = "Ikke FD"
            {
                while (mailbody[1] = "" or mailbody[1] = " ")
                    mailbody.RemoveAt(1)
                Clipboard := mailbody[1]
                sleep 150
                SendInput, {tab}
                sendinput, {f2}
                sleep 40
                sendinput ^v
                sleep 80
                SendInput, {tab}
                sleep 40
                SendInput, mtebk{tab}
                sleep 40
                SendInput, !{down}
                return
            }

            winactivate Planet - Svigt til behandling - Planet - Outlook
            sleep 100
            controlfocus, _WwG1 , Planet - Svigt til behandling - Planet - Outlook
            sleep 300
            clipboard :=
            sendinput, ^c
            clipwait, 1,
            omgang := 0
            while (Clipboard = "")
            {
                if (omgang < 5)
                {
                    winactivate Planet - Svigt til behandling - Planet - Outlook
                    sleep 100
                    controlfocus, _WwG1 , Planet - Svigt til behandling - Planet - Outlook
                    sleep 300
                    clipboard :=
                    sendinput, ^c
                    clipwait, 1,
                    omgang += 1
                }
                Else
                {
                    MsgBox, , , Fejl,
                    return
                }
            }
            sleep 50
            winactivate, Svigt FG8-FV8.xlsx - Excel
            return
        }
#IfWinActive

Fdsvigt(outlook)
{
    outlookMail := outlook.ActiveExplorer.Selection.Item(1)
    mailbody := outlookMail.body
    if InStr(outlookMail.subject, "Driftsvigt")
    {
        mailbody := StrSplit(outlookMail.body,"`r`n")
        ; mailbody[9] := SubStr(mailbody[9], 28)
        mailbody.Push("FD")

        ; MsgBox, , Er FDSvigt, % mailbody
    }
    Else
    {
        mailbody := StrSplit(outlookMail.body,"`r`n")
        test := SubStr(mailbody[1], 1, 1)
        mailbody.Push("Ikke FD")
        ; MsgBox, , Er ikke FDSvigt, % mailbody[1]
    }
    return mailbody
}
