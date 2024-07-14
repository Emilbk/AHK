global global_garanti_data := 0

; Indhenter data på garantivogne, køres ved opstartet. Global variabel defineret som global_garanti_data
; => array
indhent_garanti_data()
{
    gv_garantidage_fil := "lib/gv_garantidage.tsv"
    gv_garantidage_ind := FileRead(gv_garantidage_fil)
    gv_garantidage_ind := StrReplace(gv_garantidage_ind, "`r", "")
    gv_garantidage_ind := StrSplit(gv_garantidage_ind, "`n")
    gv_garantidage := []

    for i, e in gv_garantidage_ind
    {
        gv_garantidage.Push(StrSplit(gv_garantidage_ind[i], "`t"))
    }
    gv_garantidage_ind := unset

    return gv_garantidage
}
; sdfklsdf
global_garanti_data := indhent_garanti_data()