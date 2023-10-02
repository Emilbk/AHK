^t::
    loop
    {
        faste_dage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
        uge_dage := ["faste mandage", "faste tirsdage", "faste onsdage", "faste torsdage", "faste fredage", "faste lørdage", "faste søndage"]

        Clipboard :=
        sleep 20
        SendInput, !l{tab}
        sleep 200
        SendInput, !{right}
        sleep 400
        SendInput, ^c
        clipwait 3
        if (InStr(clipboard, "eksistere"))
            continue
        if (ErrorLevel = 1)
            SendInput, ^c
            clipwait 3
        tid_ind := Clipboard
        if (InStr(tid_ind, "senare"))
        {
            SendInput, {enter}
            for index, element in faste_dage
            {
                SendInput, !l
                sleep 200
                sendinput {tab}
                SendInput, %element% {enter}
                sleep 200
                clipboard :=
                SendInput, {tab}
                SendInput, ^c
                clipwait 3
                if (clipboard = element)
                {
                    ; MsgBox, , , match
                    tid_ind := uge_dage[index]
                    ramt_dag := index
                    ; MsgBox, , overført, % tid_ind
                    Break
                }
                else
                {
                    ; MsgBox, , ,% clipboard "." element,
                    SendInput, {enter}
                    continue
                }

            }

        }

        dato := SubStr(tid_ind, 4, 2) SubStr(tid_ind, 1, 2)
        OutputDebug, % dato
        dato:= A_YYYY dato
        FormatTime, dato, %dato%, dddd dd/MM
        if (StrLen(clipboard) = 2)
        {
            StringLower, clipboard, clipboard, ; dage med ø skal være lowercase. Hvorfor?
            for index, element in faste_dage
            {
                if (clipboard = element)
                    {
                    ramt_dag := index
                    ; MsgBox, , , % ramt_dag,
                    }
            }
            dato := uge_dage[ramt_dag]
            ; MsgBox, , næste dag , % uge_dag[ramt_dag]
        }
        ; MsgBox, , , % dato,
        sleep 100
        MsgBox, 4, Sikker?, Vil du sætte %telefon% på VL %vl% på %dato%?,
        IfMsgBox, Yes
            {
                if (ramt_dag = 7)
                    break
            sleep 200
            continue
            }
        IfMsgBox, no
            {
                break
            }
        ; MsgBox, Lop
        ; P6_tlf_vl_dato_efter(telefon)
        ; sleep s * 800
        Continue
    
    }
sleep 100
MsgBox, , , Færdig, 
return

; IfMsgBox, no
; {
; MsgBox brek

; }

return

+^t::
    {
        faste_dage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
        uge_dage := ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]
        ramt_dag := 2

        MsgBox, , , % uge_dage[ramt_dag]
        ; for index, element in faste_dage
        ;     {
        ;         MsgBox, , , % uge_dage[index]
        ;     }

        ; MsgBox, , ,% element
    }
; 04-10-2023

^e::
faste_dage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
uge_dage := ["faste mandage", "faste tirsdage", "faste onsdage", "faste torsdage", "faste fredage", "faste lørdage", "faste søndage"]

clipboard = sø

for index, element in faste_dage
    if (element = clipboard)
        MsgBox, , , Ja, 
    Else
        MsgBox, , , nej