#singleinstance, force
#noenv
sendmode, input
setbatchlines, -1
setworkingdir, %a_scriptdir%

#IfWinActive, Svigt FG8-FV8.xlsx - Excel
!q::
    {
            winactivate Planet - Svigt til behandling - Planet - Outlook
            sleep 100
            controlfocus, outlookgrid1, Planet - Svigt til behandling - Planet - Outlook
            sleep 500
            sendinput, {appskey}
            sleep 50
            sendinput, h
            sleep 50
            sendinput, {enter}
            sleep 500
            sendinput, {up}
            sleep 500
            controlfocus, _WwG1 , Planet - Svigt til behandling - Planet - Outlook
            sleep 1000
            SendInput, +{down}
            ; winactivate, Svigt FG8-FV8.xlsx - Excel
            return
        }
#IfWinActive, Svigt FG8-FV8.xlsx - Excel
!w::
    {
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
        sleep 150
        sendinput, {tab}{f2}
        sendinput, ^v{tab}
        sleep 40
        SendInput, mtebk{tab}
        sleep 40
        SendInput, !{down}
        return
    }

