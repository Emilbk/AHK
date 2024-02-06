#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, %A_linefile%\..\DSVParser\DSVParser.ahk

FileRead, input, Genudbud FG8 - FlexGaranti.txt
FormatTime, tid, YYYYMMDDHH24MISS, dd/MM-yy
output_table := "Opdateret `t" tid "`n"
input_table := TSVParser.ToArray(input)
input_table2 := []
for i,e in input_table
    for i2,e2 in e
    {
        input_table[i][i2] := input_table2[i][i2]
    }
input_table2 := input_table
for i,e in input_table
    for i2,e2 in input_table[i]
    {
        if (i2 = 37 and e2 = "") ; fjern dem uden tlf
            input_table2.RemoveAt(i)
        if (i2 = 46 and e2 = "") ; fjern dem uden vl-nummer
            input_table2.RemoveAt(i)
    }
input_table := input_table2
for i,e in input_table
    {
        output_table := output_table . input_table[i][37] "`t"
        output_table := output_table . input_table[i][46] "`n"

    }
;     for i2,e2 in input_table[i]
;         {
;             if (i2 = 46)
;                 vl := vl . input_table[i][i2] "`n"
;         }
; for i,e in input_table
;     for i2,e2 in input_table[i]
;         {
;             if (i2 = 37)
;                 tlf := tlf . "`t" input_table[i][i2] "`n"
;         }
FileDelete, FG8.txt
FileAppend, %output_table%, FG8.txt

MsgBox, , , ,