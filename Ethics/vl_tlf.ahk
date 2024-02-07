#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, %A_linefile%\..\DSVParser\DSVParser.ahk

FileRead, input, Genudbud FG8 - FlexGaranti.txt
FormatTime, tid, YYYYMMDDHH24MISS, dd/MM-yy
opdateret := "Opdateret `t" tid "`n"

input_table := TSVParser.ToArray(input)
input_table2 := []

for i,e in input_table
    {
        if (input_table[i][37] != "" and input_table[i][46] != "") ; fjern dem uden tlf
            {
            input_table2.Push(input_table[i])
            }
    }
input_table := input_table2
for i,e in input_table
    {
        if InStr(input_table[i][46], "(")
            {
            RegExMatch(input_table[i][46], "\([^)]*\)", test)
            test := SubStr(test, 2 , 7)
            if InStr(test, "_9)")
                {
                    test := SubStr(test, 1, 6)
                }
            input_table[i][46] := test
            }
    }
for i,e in input_table
    {
        output_table := output_table . input_table[i][37] "`t"
        output_table := output_table . input_table[i][46] "`n"

    }

FileDelete, FG8.txt
FileAppend, %output_table%, FG8.txt

MsgBox, , , ,