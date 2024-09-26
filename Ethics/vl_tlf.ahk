#Include dsvparser-ahk2.ahk

; inputs

inputFG8csv := FileRead("Genudbud FG8 - FlexGaranti.txt")
inputFV8csv := FileRead("FV8 - FlexVariabel.txt")
inputFV8VGcsv := FileRead("FV8 - FlexVariabel_VG.txt")

inputFG8Array := TSVParser.ToArray(inputFG8csv)
inputFV8Array := TSVParser.ToArray(inputFV8csv)
inputFV8VGArray := TSVParser.ToArray(inputFV8VGcsv)

rensFG8Array := arrayRens(inputFG8Array)
rensFV8Array := arrayRens(inputFV8Array)
rensFV8VGArray := arrayRens(inputFV8VGArray)

outputArray := []

outputArray.Push(rensFG8Array*)
outputArray.Push(rensFV8Array*)
outputArray.Push(rensFV8VGArray*)

outputTSV := TSVParser.FromArray(outputArray)
if FileExist("vl_tlf_output.txt")
FileDelete("vl_tlf_output.txt")
FileAppend(outputTSV, "vl_tlf_output.txt")

arrayRens(p_array_input)
{
    array_output := []
    for index, element in p_array_input
    {
        if (element[1] != "" and element[2] != "" and InStr(element[2], "_"))
        {
            parantes_start_pos := InStr(element[2], "(")
            parantes_slut_pos := InStr(element[2], ")")
            underscore_pos := InStr(element[2], "_")
            if (InStr(element[2], "("))
            {
                k_aftale := SubStr(element[2], parantes_start_pos + 1, underscore_pos - parantes_start_pos - 1)
                sys := SubStr(element[2], underscore_pos + 1, parantes_slut_pos - underscore_pos - 1)
            }
            else
            {
                element_split := StrSplit(element[2], "_")
                ; position := InStr(element[2], "_")
                k_aftale := element_split[1]
                sys := element_split[2]
                ; if (InStr(sys, "9)"))
                ; sys := "09"
            }
            element[2] := k_aftale . "_" sys
            array_output.Push(element)
        }
    }

    return array_output
}


return