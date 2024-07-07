#Requires AutoHotkey v2.0

;; Vognløbsdata

vognløb_indhent_data(vognløbsnummer, dato)
{
    gv_data_fil := "../lib/gv_data.txt"

    gv_data := FileRead(gv_data_fil)
    gv_data :=
    return
}

vognløb_indhent_data("31200", "07-07-2024")


; p6_svigt_tjek_ugedag(k_aftale, dato)
; {
;         gv_svigt := []
;         FileRead, gv_svigt_ind, db\gv_svigt.txt
;         gv_svigt_ind := StrReplace(gv_svigt_ind, "`r", "")
;         gv_svigt_ind := StrSplit(gv_svigt_ind, "`n")
;         for i,e in gv_svigt_ind
;             {
;                 gv_svigt[i] := StrSplit(gv_svigt_ind[i], "`t")
;             }
; for i, e in gv_svigt
;     {
;         ; MsgBox, , , %gv_svigt%, Timeout]
;         if (gv_svigt[i][1] = k_aftale)
;             {
;             vl := gv_svigt[i][2]
;             break
;             }
;     } 