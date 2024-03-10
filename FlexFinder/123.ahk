FVWorkBookSti := A_ScriptDir "\lib\Ethics\Genudbud FG8 - FlexGaranti.xlsx"
xl := ComObject("Excel.Application")
FVWorkBook := xl.WorkBooks.Open(FVWorkBookSti, , readonly := True)
sarr :=  xl.Intersect(xl.columns("A:AT"), xl.ActiveSheet.UsedRange).value
itemlist := []
loop sarr.maxindex(1)
    {
        if a_index > 1
            if SubStr(sarr[a_index, 46], 1, 1) =  3 and sarr[a_index, 6] = "approved"
                itemlist.push([sarr[a_index, 1], sarr[a_index, 6], sarr[a_index, 46]])
    }
	; (a_index > 1) && itemlist.push([sarr[a_index,1],sarr[a_index,3],sarr[a_index,8]]) ;or push a map(code, group, qty)


    MsgBox sarr[2, 4]

; for x,y in itemlist									; to show content
; 	msgbox itemlist[x][1] " " itemlist[x][2] " " itemlist[x][3]