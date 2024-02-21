#noenv
#singleinstance, force
sendmode, input
setbatchlines, -1
setworkingdir, %a_scriptdir%

gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x62 y27 w120 h21 number vvl gvl_slaa_op
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x18 y22 w35 h23 +0x200, vl:
gui vl_bod: add, text, x285 y10 w120 h23 +0x200, vm:
gui vl_bod: add, text, x285 y54 w120 h23 +0x200, kontaktinfo:
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, text, x325 y10 w120 h23 vvm +0x200, % vm
gui vl_bod: add, monthcal, x20 y58 w164 h160 vdato
gui vl_bod: add, dropdownlist, x23 y244 w414 vFG, 
gui vl_bod: add, dropdownlist, x23 y270 w414 vFV,  
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x24 y223 w120 h23 +0x200, &Paragraf
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x25 y328 w568 h84
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x25 y300 w260 h23 +0x200, "kvalitetetsbristen bestod i, at..."
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, button, default x288 y415, &ok
vl_ny := []
fileread, vl_data, db/vl_data.txt
fileread, paragraf_data, db/paragraf_data.txt
vl_data := strsplit(vl_data, "`n")
for i,e in vl_data
    {
    vl_ny[i] := strsplit(e, "`t", "`r")
    }
vl_data := vl_ny
paragraf_ny := []
paragraf_data := strreplace(paragraf_data, "`r", "")
paragraf_data := strsplit(paragraf_data, "`n")
for i,e in paragraf_data
    {
    paragraf_ny[i] := strsplit(e, "`t")
    }
paragraf_data := paragraf_ny

;; paragraf_drop_down
; msgbox, , , % substr(paragraf_data[1][1], 1,2)
paragraf_drop_down_fg := "-|"
paragraf_drop_down_fv := "-|"
for i,e in paragraf_data
    if (substr(e[1], 1 ,2) = "fg")
    {
    paragraf_drop_down_fg .= paragraf_data[i][1] "|"

    }
for i,e in paragraf_data
    if (substr(e[1], 1 ,2) = "fv")
    {
    paragraf_drop_down_fv .= paragraf_data[i][1] "|"

    }


guicontrol, vl_bod: , combobox1 , %paragraf_drop_down_fg%
guicontrol, vl_bod: , combobox2 , %paragraf_drop_down_fv%
guicontrol, vl_bod: choose, combobox1, 1
guicontrol, vl_bod: choose, combobox2, 1

vl_slaa_op:
{
    guicontrolget, vl, , edit1, 
    for i,e in vl_data
        {
            if (e[1] = vl)
                {
                    vm := e[2]
                    guicontrol, vl_bod: , vm , % vm
                    return
                }
            else
                {

                    guicontrol, vl_bod: , vm , ikke gyldigt vl
                }
        }
return
}

guifokus()
{
controlgetfocus, guifokus
return guifokus
}

; +lbutton::
; {
; if (winactive("Planet - Indbakke - Planet - Outlook"))
;     {
;         sendinput, {rbutton}
;         sleep 50
;         sendinput, h
;         sleep 50
;         sendinput, {enter}
;         sleep 200
;         sendinput, {up}
;         ; winactivate, Svigt FG8-FV8.xlsx - Excel
;         return
; ;     }
; if (winactive("Svigt FG8-FV8.xlsx - Excel"))
;     {
;         sendinput, {click2}
;         sleep 100
;         sendinput, ^v{tab}
;         return
;     }
; else
;     {
;         sendinput, +{click}
;         sleep 150
; if (winactive("Planet - Indbakke - Planet - Outlook"))
;     {
;         sendinput, {rbutton}
;         sleep 50
;         sendinput, h
;         sleep 50
;         sendinput, {enter}
;         sleep 200
;         sendinput, {up}
;         ; winactivate, Svigt FG8-FV8.xlsx - Excel
;         return
;     }
;         return
;     }
; }
^q::
{
if (winactive("Planet - Indbakke - Planet - Outlook"))
    {
        sendinput, {appskey}
        sleep 50
        sendinput, h
        sleep 50
        sendinput, {enter}
        sleep 50
        sendinput, {up}
        winactivate, Svigt FG8-FV8.xlsx - Excel
        return
    }
if (winactive("Svigt FG8-FV8.xlsx - Excel"))
    {
        sendinput, {f2}
        sendinput, ^v{tab}
        return
    }
else
    {
        sendinput, +{click}
        return
    }
return
}
!q::
{
KeyWait, Alt
if (winactive("Svigt FG8-FV8.xlsx - Excel"))
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
        sleep 1500
        sendinput, {up}
        sleep 1500
        controlfocus, _WwG1 , Planet - Svigt til behandling - Planet - Outlook
        sleep 1000
        SendInput, +{down}
        ; winactivate, Svigt FG8-FV8.xlsx - Excel
        return
    }
if (winactive("Planet - Indbakke - Planet - Outlook"))
    {
        clipboard :=
        sendinput, ^c
        clipwait, 1, 
        sleep 50
        winactivate, Svigt FG8-FV8.xlsx - Excel
        sendinput, {f2}
        sendinput, ^v{tab}  
        sleep 40
    SendInput, ^d{tab}
    sleep 40
    SendInput, !{down}
        return
    }
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
        sendinput, {f2}
        sendinput, ^v{tab}  
        sleep 40
    SendInput, ^d{tab}
    sleep 40
    SendInput, !{down}
        return
    }
vm_opslag(vl_data)
{
    inputbox, vl
    for i, e in vl_data
        for i2, e2 in e
            if (e2 = vl)
            {
               vm := vl_data[i][2]
            }
return vm
}
paragraf_opslag(paragraf_data)
{
    inputbox, paragraf
    for i, e in paragraf_data
        for i2, e2 in e
            if (e2 = paragraf)
            {
               paragraf := paragraf_data[i][2]
            }
return paragraf
; }
; !z::
; {
;     vm := paragraf_opslag(paragraf_data)   
;     msgbox, , , % vm, 
;     return
; }
; ;; gui



; !z::
; {
; gui vl_bod: show, w620 h442, window
;     return

}
vl_bodguiescape:
vl_bodguiclose:
   gui vl_bod: hide 
    return

vl_bodbuttonok:
gui vl_bod: hide
GuiControl, , edit1, 
valgt := "fv"
GuiControlGet, paragraf, , % valgt,
FormatTime, dato, %dato%, dd/MM-yyyy
MsgBox, , , VL er %vl%, VM er %vm%, dato er %dato%, FV er %fv%, FG er %FG%