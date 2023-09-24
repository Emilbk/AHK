#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%



+^t::
{
tid := P6_regn_tid()
Gui Font, s9, Segoe UI
Gui Add, Button, gok x24 y88 w80 h23, &OK
Gui Add, Button, gudklip x144 y88 w80 h23, Til &Udklip
Gui Add, Text, x72 y24 w120 h23 +0x200, %tid%

Gui Show, w260 h125, Window
Return

ok:
{
    gui, cancel
    return
}
udklip:
{
    ; Clipboard := tid.2
    ; MsgBox, , , % tid.1,
    MsgBox, , , % tid[1], 
    MsgBox, , , % tid[2], 
    
    gui, cancel
    return
}

GuiEscape:
GuiClose:
    ExitApp

}
P6_regn_tid()
{
    resultat := []
    tidA :=      ; HHmm, starttid. Enten fire cifre for klokkeslæt, mellem 1 og 3 cifre for minuttertal.
    tidB :=      ; mm, tillægstid. Minuttal
    tidC :=      ; resultat
    p6_regn_tid_ops := 1 ; 1 - med inputbox, 0 med input
    if (p6_regn_tid_ops = 1)
        {
        InputBox, tidA, Udgangspunkt, Skriv tiden`, der skal lægges noget til. `nKlokkeslæt: 4 cifre ud i ét`, minuttal: 3 til 1 ciffer ud i ét. `n `n F. eks: `n Klokken 13:34 skrives 1334 `n 231 minutter skrives 231, , , 240
        if (ErrorLevel != 0)
            return
        if (tida = "")
            tida := "0"
        InputBox, tidB, Tid `, der skal lægges til., Skriv tid`, der skal lægges til. Minuttal ud i ét (- foran`, hvis der skal trækkes fra).,
        if (ErrorLevel != 0)
            return
        if (tidb = "")
            tidb := "0"
        if (tidb + tida < 0)
            {
                MsgBox, , Lad vær', , 
                return 
            }
        if (StrLen(tida) <= "3")
            {
                tid_nul := A_YYYY A_MM A_DD "00" "00"
                EnvAdd, tid_nul, tida , minutes
                EnvAdd, tid_nul, tidb, minutes
                FormatTime, tidC, %tid_nul%, HHmm
                FormatTime, tid_time, %tid_nul%, H
                FormatTime, tid_min, %tid_nul%, m
                FormatTime, result_mid, %tid_nul%, HHmm
                if (tid_time = "0" and tid_min = "1")
                    {
                    resultat.1 := tid_min " minut."
                    resultat.2 := result_mid
                    return resultat
                    }
                if (tid_time = "0" and tid_min >= "1")
                    {
                    resultat.1 :=  tid_min " minutter."
                    resultat.2 := result_mid
                    return resultat
                    }
                if (tid_time = "1" and tid_min = "0")
                    {
                    resultat.1 := tid_time " time."
                    resultat.2 := result_mid
                    return resultat
                    }
                if (tid_time > "1" and tid_min = "0")
                    {
                    resultat.1 := tid_time " timer."
                    resultat.2 := result_mid
                    return resultat
                    }
                if (tid_time = "1" and tid_min = "1")
                    {
                    resultat.1 := tid_time " time og " tid_min " minut."
                    resultat.2 := result_mid
                    return resultat
                    }            
                if (tid_time >= "1" and tid_min = "1")
                    {
                    resultat.1 := tid_time " timer og " tid_min " minut."
                    resultat.2 := result_mid
                    return resultat
                    }
                if (tid_time >= "1" and tid_min >= "1")
                    {
                    resultat.1 :=  tid_time " timer og " tid_min " minutter."
                    resultat.2 := result_mid
                    return resultat
                    }
                }
        if (StrLen(tida) = "4")
            {
            tidA := A_YYYY A_MM A_DD tida "00"
            EnvAdd, tidA, tidB, minutes
            FormatTime, tid_time, %tidA%, HH
            FormatTime, tid_min, %tidA%, mm
            if (tid_time != "00")
                {
                resultat :=  tid_time ":" tid_min "."
                return resultat
                }
            }
        return
        }
    if (p6_regn_tid_ops = 0)
        {
            Input, tida,,{enter}
            if (ErrorLevel = "Match")
                return
            if (tida = "")
                tida := "0"
            Input, tidb,,{enter}
            if (ErrorLevel = "Match")
                return
            if (tidb = "")
                tidb := "0"
            if (tidb + tida < 0)
                {
                    MsgBox, , Lad vær', , 
                    return 
                }
            if (StrLen(tida) <= "3")
                {
                    tid_nul := A_YYYY A_MM A_DD "00" "00"
                    EnvAdd, tid_nul, tida , minutes
                    EnvAdd, tid_nul, tidb, minutes
                    FormatTime, tidC, %tid_nul%, HHmm
                    FormatTime, tid_time, %tid_nul%, H
                    FormatTime, tid_min, %tid_nul%, m
                    if (tid_time = "0" and tid_min = "1")
                        {
                        MsgBox, , , % tid_min " minut."
                        return
                        }
                    if (tid_time = "0" and tid_min >= "1")
                        {
                        MsgBox, , , % tid_min " minutter."
                        return
                        }
                    if (tid_time = "1" and tid_min = "0")
                        {
                        MsgBox, , , % tid_time " time."
                        return
                        }
                    if (tid_time > "1" and tid_min = "0")
                        {
                        MsgBox, , , % tid_time " timer."
                        return
                        }
                    if (tid_time = "1" and tid_min = "1")
                        {
                        MsgBox, , , % tid_time " time og " tid_min " minut."
                        return
                        }            
                    if (tid_time >= "1" and tid_min = "1")
                        {
                        MsgBox, , , % tid_time " timer og " tid_min " minut."
                        return
                        }
                    if (tid_time >= "1" and tid_min >= "1")
                        {
                        MsgBox, , , % tid_time " timer og " tid_min " minutter."
                        return
                        }
                    }
            if (StrLen(tida) = "4")
                {
                tidA := A_YYYY A_MM A_DD tida "00"
                ; MsgBox, , , % tidA
                EnvAdd, tidA, tidB, minutes
                FormatTime, tid_time, %tidA%, HH
                FormatTime, tid_min, %tidA%, mm
                if (tid_time != "00")
                    {
                    MsgBox, , , % tid_time ":" tid_min "."
                    return
                    }
                }
            
            return
            }
} 
^+r::
    SendInput, {CtrlUp}
    Reload
    sleep 2000
Return