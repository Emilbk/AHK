#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, %A_linefile%\..\DSVParser\DSVParser.ahk

FileRead, input, Genudbud FG8 - FlexGaranti.txt
input_table := TSVParser.ToArray(input)
for i,e in input_table
    for i2,e2 in input_table[i]
    {
        if (i2 = 37 and e2 = "") ; fjern dem uden tlf
            input_table.RemoveAt(i)
        if (i2 = 46 and e2 = "") ; fjern dem uden vl-nummer
            input_table.RemoveAt(i)
    }
for i,e in input_table
    for i2,e2 in input_table[i]
        {
            if (i2 = 37)
                output_table := output_table . input_table[i][i2] "`n"
        }
FileDelete, test.txt
FileAppend, %output_table%, test.txt

MsgBox, , , ,