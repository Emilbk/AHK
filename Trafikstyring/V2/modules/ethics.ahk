; testing
; Excel
; TODO #96 omskriv til generel funktion
root_path_array := StrSplit(A_ScriptDir, "Trafikstyring V2")
root_path := root_path_array[1] . "\Trafikstyring V2"
excel := ComObject("excel.application")
excel.visible := 0

gv_excel_fil := root_path . "\Trafikstyring\V2\lib\Ethics\Genudbud FG8 - FlexGaranti.xlsx"
variabel_excel_fil := root_path . "\Trafikstyring\V2\lib\Ethics\FV8 - FlexVariabel.xlsx"
vl_data := Map("vognløbsnummer", [], "kørselsaftale", [], "telefonnummer", [])

gv_telefon_række := "AK"
gv_vl_række := "AT"
gv_worksheet := 1
variabel_telefon_række := "AU"
variabel_vl_række := "BD"
variabel_worksheet := 1


; excelrække [vl_række, telefonrække]
excel_række(p_excelob, p_excelfil, p_excelark, p_excelrække)
{
    ; MsgBox(type(excelrække))
    if (Type(p_excelrække) != "array")
        throw ValueError("Excelrække skal være array")


    excel_array := map()
    xl := p_excelob.workbooks.open(p_excelfil, , readonly := true)
    xl_ark := xl.sheets(p_excelark)
    usedrange := xl_ark.cells(xl_ark.Rows.Count, 1).end(-4162).row
    excel_map := map()
    for i, e in p_excelrække
    {
        excel_map.Set(e, [])

    }


    for i, række in p_excelrække
    {
        loop usedrange
            excel_map[række].push(xl_ark.range(række A_Index).value)

    }
    return excel_map
}

excel_rens_gv(excel_map, vognløbsrække, telefonrække, vl_data)
{
    for i, e in excel_map[vognløbsrække]
        if (substr(e, 1, 1) = "3" and excel_map[telefonrække][i] != "" and i > 1)
        {
            vl := StrSplit(e, " ")
            if (vl.MaxIndex = 1)
            {
                vl_data["vognløbsnummer"].push(vl[1])
                return
            }
            vl_data["vognløbsnummer"].push(vl[1])
            vl_data["kørselsaftale"].push(substr(vl[2], 2, 7))
            vl_data["telefonnummer"].Push(excel_map[telefonrække][i])
        }

}
gv_map := excel_række(excel, gv_excel_fil, gv_worksheet, [gv_vl_række, gv_telefon_række])
variabel_map := excel_række(excel, variabel_excel_fil, variabel_worksheet, [variabel_vl_række, variabel_telefon_række])
excel_rens_gv(gv_map, gv_vl_række, gv_telefon_række, vl_data)
excel_rens_gv(variabel_map, variabel_vl_række, variabel_telefon_række, vl_data)


; gv.Delete("vognløb_ind")

return