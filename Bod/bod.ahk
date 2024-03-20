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

#IfWinActive, Svigt FG8-FV8.xlsx - Excel
    !q::
        {
            winactivate Planet - Svigt til behandling - Planet - Outlook
            sleep 100
            controlfocus, outlookgrid1, Planet - Svigt til behandling - Planet - Outlook
            sleep 100
            sendinput, {appskey}
            ; ControlClick, Outlookgrid1, Planet - Svigt til behandling - Planet - Outlook, , Right, 1
            ; ControlSend, Outlookgrid1, {AppsKey}, Planet - Svigt til behandling - Planet - Outlook
            ; return
            sleep 100
            sendinput, h
            sleep 50
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
                Clipboard := mailbody[9]
                sleep 150
                sendinput, {f2} ^v
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
                sendinput, {f2}^v
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
        mailbody[9] := SubStr(mailbody[9], 28)
        test := SubStr(mailbody[9], 1, 1)
        while (test = "")
        {
            test := SubStr(mailbody[9], 1, 1)
            mailbody[9] := SubStr(mailbody[9], 2)
        }
        mailbody.Push("FD")

        ; MsgBox, , Er FDSvigt, % mailbody
    }
    Else
    {
        mailbody := StrSplit(outlookMail.body,"`r`n")
        test := SubStr(mailbody[1], 1, 1)
        while (test = "")
        {
            test := SubStr(mailbody[1], 1, 1)
            mailbody[1] := SubStr(mailbody[1], 2)
        }
        mailbody.Push("Ikke FD")
        ; MsgBox, , Er ikke FDSvigt, % mailbody[1]
    }
    return mailbody
}
