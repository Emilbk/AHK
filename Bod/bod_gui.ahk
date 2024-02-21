#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
;; GUI

gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x18 y22 w35 h23 +0x200, vl:
gui vl_bod: add, text, x285 y10 w120 h23 +0x200, vm:
gui vl_bod: add, text, x285 y54 w120 h23 +0x200, kontaktinfo:
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x62 y27 w120 h21 number vvl gvl_slaa_op
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, text, x380 y10 w200 h23 vvm +0x200, % vm
gui vl_bod: add, text, x380 y54 w200 h23 vemail +0x200, % email
gui vl_bod: add, dateTime, vdato x20 y58 w164 h60, 
; gui vl_bod: add, monthcal, x20 y58 w164 h160 vdato
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, text, x24 y170 w120 h23 +0x200, &Søg Paragraf
gui vl_bod: add, edit, x24 y200 w120 h23 +0x200 vparagraf_søg gparagraf_slaa_op, 
gui vl_bod: add, text, x24 y223 w120 h23 +0x200, &Paragraf
gui vl_bod: add, dropdownlist, x23 y244 w414 vFG, 
gui vl_bod: add, dropdownlist, x23 y270 w414 vFV,
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x25 y300 w260 h23 +0x200, &kvalitetsbristen bestod i, at...
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x25 y328 w568 h84 vbrist
gui vl_bod: add, button, default x288 y415, &ok

fileread, paragraf_data, db/paragraf_data.txt
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
    if (substr(e[2], 1 ,2) = "fg")
    {
    paragraf_drop_down_fg .= paragraf_data[i][2] "|"

    }
for i,e in paragraf_data
    if (substr(e[2], 1 ,2) = "fv")
    {
    paragraf_drop_down_fv .= paragraf_data[i][2] "|"

    }


;; STAMDATA

stamopl_sti := "C:\Users\ebk\Stamoplysninger FV8 og FG8.xlsx"

stamopl:= ComObjCreate("Excel.application")
; stamopl.Workbooks.Open(stamopl_sti,, readonly := false)
; stamopl_workbook := 
stamopl_workbook := stamopl.workbooks.open(stamopl_sti,, readonly := true)
stamopl.visible := 0
stamopl_ark := stamopl.sheets("ark1") ; after opening workbook its better to define sheet 
stamopl_kolonne_a := stamopl_ark.range("A:A") 
stamopl_kolonne_b := stamopl_ark.range("B:B") 
r_a_sidste := stamopl_kolonne_a.end(-4121).row
r_b_sidste := stamopl_kolonne_b.end(-4121).row
vm_stam := []    
kontakt_stam := []
loop, %r_a_sidste%
    {
         vm_stam.push(stamopl_ark.range("A" A_index).value)
         kontakt_stam.push(stamopl_ark.range("L" A_index).value)
    }
vm_stam.RemoveAt(1)
kontakt_stam.RemoveAt(1)
stamopl.quit()

stamopl_sti := "C:\Users\ebk\Svigt FG8-FV8.xlsx"

stamopl:= ComObjCreate("Excel.application")
; stamopl.Workbooks.Open(stamopl_sti,, readonly := false)
; stamopl_workbook := 
stamopl_workbook := stamopl.workbooks.open(stamopl_sti,, readonly := true)
stamopl.visible := 0
stamopl_ark := stamopl.sheets(4) ; after opening workbook its better to define sheet 
stamopl_kolonne_a := stamopl_ark.range("A:A") 
stamopl_kolonne_b := stamopl_ark.range("B:B") 
r_a_sidste := stamopl_kolonne_a.end(-4121).row
r_b_sidste := stamopl_kolonne_b.end(-4121).row

vl_svigt := []
vm_svigt := []
email_svigt := []
fundet := []
loop, %r_a_sidste%
    {
         vl_svigt.push(stamopl_ark.range("A" A_index).value)
         vm_svigt.push(stamopl_ark.range("B" A_index).value)
    }
stamopl.quit()
for i,e in vl_svigt
    {
        if e is number
            {
                vl_svigt[i] := Format("{:d}", e)
                
            }
    
    }
stamdata := []
for i,e in vl_svigt
    {
        stamdata[i] := [vl_svigt[i], vm_svigt[i]]
    }    

for i, e in vm_svigt
    {
        for i2, e2 in vm_stam
            if (e = e2)
                {
                stamdata[i].Push(kontakt_stam[i2])
                Break 1
                }
    }


; for i,e in stamdata
;     {
;         MsgBox, , , % "Vl " stamdata[i][1] " tilhører " stamdata[i][2] ", som har email " stamdata[i][3]
;     }
; ; stamopl_ark :½= stamopl_workbook.worksheets("Ark1")
; test := stamopl.worksheets(stamopl_ark).columns(1)
; stamopl.workbooks()
; vm := stamopl_ark.range("A:A").end("xldown")

guicontrol, vl_bod: , combobox1 , %paragraf_drop_down_fg%
guicontrol, vl_bod: , combobox2 , %paragraf_drop_down_fv%
guicontrol, vl_bod: choose, combobox1, 1
guicontrol, vl_bod: choose, combobox2, 1


gui vl_bod: show, w620 h442, window

+esc::
{
    stamopl.quit()
    ExitApp
}
; oWorkbook := ComObjCreate("Excel.Application")
; oWorkbook.Workbooks.open(FilePath,, readonly := true)
; oWorkbook.Visible := 0 
; clientsname := oWorkbook.Worksheets("test doc").Range("A3").Value
; StringRight, clientsname, clientsname, 5
; clientsphone := oWorkbook.Worksheets("test doc").Range("B3").Value
; clientsstate := oWorkbook.Worksheets("test doc").Range("C3").Value
; clientsfax := oWorkbook.Worksheets("test doc").Range("D3").Value


;; GUI-funktion

paragraf_slaa_op:
{
    guicontrolget, vl, , edit2, 
    if (vl = "")
        guicontrolget, vl, , edit2, 
    for i,e in paragraf_data
        {
            if (InStr(e[2], fg))
                {
                    guicontrol, vl_bod: , fg , % paragraf_data[i][2]
                    break
                    return
                }
            if (InStr(e[2], fv))
                {
                    guicontrol, vl_bod: , fv , % paragraf_data[i][2]
                    break
                    return
                }
            else
                {

                }
        }
return
}

vl_slaa_op:
{
    guicontrolget, vl, , edit1, 
    if (vl = "")
        guicontrolget, vl, , edit1, 
    for i,e in stamdata
        {
            if (e[1] = vl)
                {
                    vm := e[2]
                    email := e[3]
                    guicontrol, vl_bod: , vm , % vm
                    guicontrol, vl_bod: , email , % email
                    return
                }
            else
                {

                    guicontrol, vl_bod: , vm , ikke gyldigt vl
                    guicontrol, vl_bod: , email ,  
                }
        }
return
}
;; GUI-label
vl_bodguiescape:
vl_bodguiclose:
    stamopl.quit()
    ExitApp

vl_bodbuttonok:
gui Submit, nohide
if (fg != "-" and fv != "-")
    {
        MsgBox, 16, Både FG og FV valgt, Der skal kun vælges fra ét udbud.
        return
    }
    for i,e in paragraf_data
        {
            if (e[2] = fg) or (e[2] = fv) 
                {
                    paragraf := paragraf_data[i][3]
                    bod := paragraf_data[i][1]
                    break
                }
        }


FormatTime, dato, %dato%, dd-MM-yyyy

test = 
(
Til
%vm%
Bod for kvalitetsbrist
 
Midttrafik har den %dato% registreret en kvalitetsbrist på vognløb %vl%, der medfører en bod på kr. %bod%,- jf. FG8, side 52, § 31, stk. 3, litra 

%paragraf%
 
Kvalitetsbristen bestod i, at %brist%
 
Beløbet vil blive modregnet i vognmandsafregningen.
Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.
)

MsgBox, , , %test%


guicontrol, vl_bod: , vl , 
return


