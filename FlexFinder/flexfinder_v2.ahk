/************************************************************************
 * @description 
 * @file flexfinder_v2.ahk
 * @author 
 * @date 2024/03/09
 * @version 0.0.0
 ***********************************************************************/

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent


;; Include
#Include "JXON.ahk"


; goo := Gui()
; goo.btn_exit := goo.AddButton('xm ym w100', 'Exit Script')
; goo.btn_exit.OnEvent('Click', (*) => ExitApp())
; goo.Show('AutoSize')
GruppeArray := ["GV", "GV Lift", "GV Lav", "GV TTJ Hjul", "GV TTJ Larve", "GV Type 7", "GV Type 8", "GV Type 9", "FV Alle", "FV Lav", "FV Lift", "FV TTJ Hjul", "FV TTJ Larve", "FV Type 7"]
indlæst_data := Map()
;; XL

;; Menubar
FFgui := Gui(, "FlexFinder-Grupper",)
FilMenu := Menu()
FilMenu.Add("&Indlæs data fra Excel`tCtrl+I", (*) => (indlæst_data["VLdata"] := indlæs_data_excel()))
FilMenu.Add()
; FilMenu.Add("Exit", gui_exit)
FFMenuBar := MenuBar()
FFMenuBar.Add("&Fil", Filmenu)
FFgui.MenuBar := FFMenuBar


FFgui.AddGroupBox("W300 H200", "Grupper")
ListboxGruppe := FFgui.AddListbox("Choose1 R10 T32 YP20 XP5 Section Multi", GruppeArray)
KnapOpretGruppe := FFgui.AddButton("XP", "&Opret gruppe")
TekstGruppe := FFgui.Addtext("YS W150", "Tekst om grupper`nBlasdlkasjdlkjasdl asdlkja sldkasjd lkasjd`nASdlkjasd lkasjdl kjasdlkj asdlkj asldjk")
; KnapOpretGruppe.OnEvent("Click", (*) => indlæs_data_excel())
KnapOpretGruppe.OnEvent("Click", (*) => OpretGruppe())
FFgui.show("Autosize")


OpretGruppe()
{
    ValgtGruppeTekst := ListboxGruppe.Text
    ValgtGruppeRække := ListboxGruppe.Value
    if ValgtGruppeTekst is String
        Knapvalg := MsgBox("Du har valgt `"" ValgtGruppeTekst "`".`n`nOpret gruppen i FlexFinder?", "Valg af FF-Gruppe", "YesNo")
    if ValgtGruppeTekst is Array
    {
        for i, e in ValgtGruppeTekst
        {
            ValgtGruppeTekstString .= ValgtGruppeTekst[i] "`n"
        }
        ValgtGruppeTekstString := SubStr(ValgtGruppeTekstString, 1, -1)
        Knapvalg := MsgBox("Du har valgt: `n`n" ValgtGruppeTekstString "`n`nOpret grupperne i FlexFinder?", "Valg af FF-Gruppe", "YesNo")
    }
    if Knapvalg = "Yes"
    {
        MsgBox("OK")
    }

}
indlæs_data_excel(*)
{
    vldata := Map("GVData", 0, "FVData", 0)
    ; gvdata := []
    vldata["GVData"] := (indlæs_data_gv())
    vldata["FVData"] := (indlæs_data_fv())
    ; vldata["GVdata"] := gvdata
    return vldata
}

indlæs_data_fv()
{
    FVData := [[]]

    FVWorkBookSti := "Z:\delt\dokumenter\AHK\FV8 - FlexVariabel.xlsx"
    xl := ComObject("Excel.Application")
    FVWorkBook := xl.WorkBooks.Open(FVWorkBookSti, , readonly := True)
    FVActiveSheet := FVWorkBook.ActiveSheet
    FVUsedRange := FVActiveSheet.UsedRange
    FVAntalRækker := FVUsedRange.Rows.Count
    FVAntalKolonner := FVUsedRange.Columns.Count
    FVAktivRække := 0

    ; Indlæs FV-Data
    loop FVAntalRækker
    {
        i := A_Index
        FirmaRange := "A" i
        VLrange := "BD" i
        ChfTlfRange := "AU" i
        VLVogntypeRange := "W" i
        VL := FVActiveSheet.Range(VLrange).Value
        VLTjek := SubStr(VL, 1, 1)
        ChfTlf := FVActiveSheet.Range(ChfTlfRange).Value
        if (InStr(VlTjek, "3") and vl != "" and ChfTlf != "")
        {
            FVData[FVData.length].Push(i)
            FVData[FVData.length].Push(FVActiveSheet.Range(FirmaRange).value)
            FVData[FVData.length].Push(FVActiveSheet.Range(VLrange).value)
            FVData[FVData.length].Push(FVActiveSheet.Range(ChfTlfRange).value)
            FVData[FVData.length].Push(FVActiveSheet.Range(VlVognTypeRange).value)
            FVData.Push([])
        }
        ; FVData.push([])
        ; FVData[i].Push(FVActiveSheet.Range(Trafikselskab).Value)
    }
    FVData.RemoveAt(FVData.length)
    for i, e in FVData
    {
        if (InStr(FVData[i][3], "(")) ; Tjek for VG el. Drift
        {
            FVData[i].push("Drift")
            FVData[i][3] := StrSplit(e[3], ["(", ")"])
            FVData[i][3][1] := SubStr(FVData[i][3][1], 1, -1)
            FVData[i][3].RemoveAt(3)
        }
        else
            FVData[i].Push("VG")
    }
    JsonArray := Jxon_Dump(FVData, indent := 0)
    FileDelete("FVData.txt")
    FileAppend(JsonArray, "FVData.txt")
    MsgBox "Data er indlæst", "Done!"
    return FVdata
}
indlæs_data_gv()
{
    GVData := [[]]

    GVWorkBookSti := "Z:\delt\dokumenter\AHK\Genudbud FG8 - FlexGaranti.xlsx"
    xl := ComObject("Excel.Application")
    GVWorkBook := xl.WorkBooks.Open(GVWorkBookSti, , readonly := True)
    GVActiveSheet := GVWorkBook.ActiveSheet
    GVUsedRange := GVActiveSheet.UsedRange
    GVAntalRækker := GVUsedRange.Rows.Count
    GVAntalKolonner := GVUsedRange.Columns.Count
    GVAktivRække := 0

    ; Indlæs GV-Data
    loop GVAntalRækker
    {
        i := A_Index
        FirmaRange := "A" i
        TrafikselskabRange := "V" i
        VLrange := "AT" i
        ChfTlfRange := "AK" i
        VLVogntypeRange := "X" i
        Trafikselskab := GVActiveSheet.Range(TrafikselskabRange).Value
        VL := GVActiveSheet.Range(VLrange).Value
        ChfTlf := GVActiveSheet.Range(ChfTlfRange).Value
        if (InStr(Trafikselskab, "Midttrafik") and vl != "" and ChfTlf != "")
        {
            GVData[GVData.length].Push(i)
            GVData[GVData.length].Push(GVActiveSheet.Range(FirmaRange).value)
            GVData[GVData.length].Push(GVActiveSheet.Range(VLrange).value)
            GVData[GVData.length].Push(GVActiveSheet.Range(ChfTlfRange).value)
            GVData[GVData.length].Push(GVActiveSheet.Range(VlVognTypeRange).value)
            GVData.Push([])
        }
        ; GVData.push([])
        ; GVData[i].Push(GVActiveSheet.Range(Trafikselskab).Value)
    }
    GVData.RemoveAt(GVData.length)
    for i, e in GVData
    {
        GVData[i][3] := StrSplit(e[3], ["(", ")"])
        GVData[i][3][1] := SubStr(GVData[i][3][1], 1, -1)
        GVData[i][3][3] := SubStr(GVData[i][3][3], 2,)

    }
    JsonArray := Jxon_Dump(GVData, indent := 0)
    FileDelete("GVData.txt")
    FileAppend(JsonArray, "GVData.txt")
    MsgBox "Data er indlæst", "Done!"
    return GVdata
}
; vaelg_fil := Gui(, "GUI",)
; ; vaelg_vil_FilMenu := menu("&Indlæs data`tCtrl+i")
; ; vaelg_vil_FilMenu.add()
; vaelg_fil_valg := ["123", "232", "234234"]
; vaelg_combo := vaelg_fil.AddComboBox('+Vvaelg_liste, Ccc1212 x200 y200', vaelg_fil_valg)
; vaelg_knap := vaelg_fil.Add("Button", "x200", "&En Knap")
; vaelg_knap_vis := vaelg_fil.Add("Button", "y300", "&Vis")
; vaelg_knap.OnEvent("Click", gui_knap)
; vaelg_knap_vis.OnEvent("Click", gui_knap_vis)
; vaelg_fil.Menubar := GuiMenu
; vaelg_fil.Show("w800 h900 Center")
; ; sldfjklsdjf
; gui_knap(ctrlobj, info)
; {

;     valgt := vaelg_combo.value
;     if (valgt = 0)
;         vaelg_fil_valg.Push(vaelg_combo.Text)
;     vaelg_combo.delete()
;     vaelg_combo.add(vaelg_fil_valg)
;     vaelg_combo.choose(vaelg_fil_valg.length)
;     ; valgt := vaelg_combo.value
;     return
;     ; MsgBox(vaelg_fil_valg[valgt])
; }
; gui_knap_vis(a, e)
; {
;     valgt := vaelg_combo.value
;     if valgt = 0
;     {
;         MsgBox("Du skal vælge noget", "", "icon!")
;         return
;     }
;     MsgBox(vaelg_fil_valg[valgt])
;     return
; }

; indlæs_data()
; {
;     MsgBox("Data er indlæst", "Vognløbsdata", "iconi")
; }

; gui_exit(a, x, y)
; {
;     MsgBox("1" a " 3" x)
;     ExitApp()
; }
