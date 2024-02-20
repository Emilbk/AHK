#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

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
xl.quit()

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


for i,e in stamdata
    {
        MsgBox, , , % "Vl " stamdata[i][1] " tilhører " stamdata[i][2] ", som har email " stamdata[i][3]
    }
; stamopl_ark :½= stamopl_workbook.worksheets("Ark1")
; test := stamopl.worksheets(stamopl_ark).columns(1)
; stamopl.workbooks()
; vm := stamopl_ark.range("A:A").end("xldown")
MsgBox, , , % r_a_sidste

+esc::
{
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
Xl.Quit()