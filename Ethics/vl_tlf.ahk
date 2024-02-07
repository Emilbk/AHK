#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, %A_linefile%\..\DSVParser\DSVParser.ahk

FileRead, FV8_input, FV8 - FlexVariabel.txt
FileRead, FV8_VG_input, FV8 - FlexVariabel_VG.txt
FileRead, FG8_input, Genudbud FG8 - FlexGaranti.txt
FormatTime, tid, YYYYMMDDHH24MISS, dd/MM-yy
opdateret := "Opdateret `t" tid "`n"

FG8_output := tekst_til_array(FG8_input)
FV8_output := tekst_til_array(FV8_input)
FV8_VG_output:= tekst_til_array(FV8_VG_input)

tekst_til_array(input)
{
    input:= TSVParser.ToArray(input)
    input2 := []
for i,e in input
    for i2,e2 in e
    if (i = 1)
    {
        if (e2 = "Vognløbsnummer")
            vl := i2
        if (e2 = "Telefonnummer til chauffør" )
            tlf := i2
    }
    else Break 1
for i,e in input
    {
        if (input[i][tlf] != "" and input[i][vl] != "") 
            {
            input2.Push(input[i])
            }
    }

input := input2
for i,e in input
    {
        if InStr(input[i][vl], "(")
            {
            RegExMatch(input[i][vl], "\([^)]*\)", test)
            test := SubStr(test, 2 , 7)
            if InStr(test, "_9)")
                {
                    test := SubStr(test, 1, 6)
                }
            input[i][vl] := test
            }
    }

for i,e in Input
    {
        ; tjek for dobbelt VL
    }

for i,e in input
    {
        output := output . input[i][tlf] "`t"
        output := output . input[i][vl] "`n"

    }

return output
}



FileDelete, FG8_resultat.txt
FileDelete, FV8_resultat.txt
FileDelete, FV8_VG_resultat.txt
FileAppend, %FG8_output%, FG8_resultat.txt
FileAppend, %FV8_output%, FV8_resultat.txt
FileAppend, %FV8_VG_output%, FV8_VG_resultat.txt

FileDelete, samlet.txt
FileAppend, Opdateret d. %tid% `n, samlet.txt
FileAppend, % FG8_output, samlet.txt
FileAppend, % FV8_output, samlet.txt
FileAppend, % FV8_VG_output, samlet.txt

MsgBox, , , Færdig