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
GruppeArrayString := ["Alle GV", "GV Lift", "GV Type 2", "GV TTJ Hjul", "GV TTJ Larve", "GV Type 7", "GV Type 8", "GV Type 9", "FV Alle", "FV Lav", "FV Lift", "FV TTJ Hjul", "FV TTJ Larve", "FV Type 7"]
GruppeArray := [[[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [[], []], [GruppeArrayString]]
indlaest_data := Map()
;; Indlæs fra Json

if !FileExist("lib\Ethics\Genudbud FG8 - FlexGaranti.xlsx") or !FileExist("lib\Ethics\FV8 - FlexVariabel.xlsx")
{
    MsgBox("Ingen Excel-data fundet!`n`nLæs Readme.", "Fejl", "Iconx")
    ExitApp()
}
if !FileExist("lib\GVData.txt") or !FileExist("lib\FVData.txt")
{
    MsgBox("Ingen tidligere data fundet - Indlæser nye.`n`nDet kan tage lidt tid.", "Ingen data", "icon!")
    IndlaesDataAlleOnevent()
}
GvData := FileRead("GVdata.txt")
FvData := FileRead("FVData.txt")
GVdata := Jxon_Load(&GVdata)
FVData := Jxon_Load(&FVdata)
indlaest_data["GVData"] := GvData
indlaest_data["FVData"] := FvData
indlaest_data["Opdateret GV"] := FormatTime(FileGetTime("lib\Ethics\Genudbud FG8 - FlexGaranti.xlsx"), "dd/MM/yyyy kl. HH:mm")
indlaest_data["Opdateret FV"] := FormatTime(FileGetTime("lib\Ethics\FV8 - FlexVariabel.xlsx"), "dd/MM/yyyy kl. HH:mm")
test := FunkIndelGrupper(GruppeArrayString, indlaest_data)
;; Menubar
FFgui := Gui(, "FlexFinder-Grupper",)
FilMenu := Menu()
FilMenu.Add("&Indlæs alle data fra Excel`tCtrl+I", (*) => (indlaest_data["VLdata"] := IndlaesDataAlleOnevent(), TekstGruppe.text := "Tekst om grupper`nData senest opdateret: " FormatTime(FileGetTime("GVData.txt"), "dd/MM/yyyy kl. HH:mm:ss")))
FilMenu.Add("&Indlæs GV fra Excel", (*) => (indlaest_data["GVdata"] := IndlaesDataFGOnevent(), TekstGruppe.text := "Tekst om grupper`nData senest opdateret: " FormatTime(FileGetTime("GVData.txt"), "dd/MM/yyyy kl. HH:mm:ss")))
FilMenu.Add("&Indlæs FV fra Excel", (*) => (indlaest_data["FVdata"] := IndlaesDataFVOnevent(), TekstGruppe.text := "Tekst om grupper`nData senest opdateret: " FormatTime(FileGetTime("GVData.txt"), "dd/MM/yyyy kl. HH:mm:ss")))
FilMenu.Add()
FilMenu.Add("Afslut", (*) => ExitApp())
FFMenuBar := MenuBar()
FFMenuBar.Add("&Fil", Filmenu)
FFMenuBar.Add("&?", Filmenu, "Right")
FFgui.MenuBar := FFMenuBar

FFgui.AddGroupBox("W300 H200", "Grupper")
ListboxGruppe := FFgui.AddListbox("Choose1 R10 T32 YP20 XP5 Section Multi", GruppeArrayString)
KnapOpretGruppe := FFgui.AddButton("XP", "&Opret gruppe")
KnapVisGruppe := FFgui.AddButton("YP", "&Vis gruppe")
TekstGruppe := FFgui.Addtext("YS W130", "Data senest opdateret:`nGV - " indlaest_data["Opdateret GV"] "`nFV - " indlaest_data["Opdateret FV"])
KnapOpretGruppe.OnEvent("Click", (*) => FunkOpretGruppe())
KnapVisGruppe.OnEvent("Click", FunkVisGruppe.Bind(GruppeArray, indlaest_data))
KnapVisTest := FFgui.AddButton("XP", "&Test")
KnapVisTest.OnEvent("Click", (*) => TekstGruppe.Text := ("Tekst om gruppret: " indlaest_data["Opdateret"]))
FFgui.show("Autosize")

FunkIndelGrupper(gruppearay, indlaest_data)
{

    ; Alle GV
    for i, e in indlaest_data["GVData"]
    {
        GruppeArray[1][1].push(e[3][1])
        GruppeArray[1][2].push(e[3][2])
    }
    ; Lift [2]
    for i, e in indlaest_data["GVData"]
        for i2, e2 in ["Type 5", "Type 6", "Type 7", "Type 8", "Type 9"]
            if (instr(e[3][3], e2))
            {
                GruppeArray[2][1].push(e[3][1])
                GruppeArray[2][2].push(e[3][2])
            }
    ; Lav [3]
    for i, e in indlaest_data["GVData"]
        if (e[3][3] = "Type 2")
        {
            GruppeArray[3][1].push(e[3][1])
            GruppeArray[3][2].push(e[3][1])
        }
    return GruppeArray
}

FunkVisGruppe(gruppearray, indlaest_data, *)
{
    Valgt := ListboxGruppe.Value
    vl := ""
    k := ""
    if (Valgt[1] = 1) ; Alle GV
    {
        for i, e in gruppearray[1][1]
            vl .= e ", "
        vl := SubStr(vl, 1, -2)
        vl .= "."
        for i, e in gruppearray[1][2]
            k .= e ", "
        k := SubStr(k, 1, -2)
        k .= "."
    }
    MsgBox("FlexFinder-gruppen er:" gruppearray[15][1][Valgt[1]] "`nVL er:`n" Vl "`nKørselsaftalerne er :`n" k)
}
FunkOpretGruppe()
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
IndlaesDataAlleOnevent(*)
{
    vldata := Map("GVData", 0, "FVData", 0)
    vldata["GVData"] := (FunkIndlaesDataFG())
    vldata["FVData"] := (FunkIndlaesDataFV())
    return vldata
}
IndlaesDataFGOnevent(*)
{
    vldata := Map("GVData", 0, "FVData", 0)
    vldata["GVData"] := (FunkIndlaesDataFV())
    return vldata
}
IndlaesDataFVOnevent(*)
{
    vldata := Map("GVData", 0, "FVData", 0)
    vldata["FVData"] := (FunkIndlaesDataFV())
    return vldata
}


FunkIndlaesDataFV()
{
    FVData := [[]]

    FVWorkBookSti := A_ScriptDir "\lib\Ethics\FV8 - FlexVariabel.xlsx"
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
    datafil := A_ScriptDir "\lib\FVData.txt"
    if FileExist(datafil)
        FileDelete(datafil)
    FileAppend(JsonArray, datafil)
    MsgBox "FV-data er indlæst", "Done!"
    return FVdata
}
FunkIndlaesDataFG()
{
    GVData := [[]]

    GVWorkBookSti := A_scriptdir "\lib\Ethics\Genudbud FG8 - FlexGaranti.xlsx"
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
    datafil := A_ScriptDir "\lib\GVData.txt"
    if FileExist(datafil)
        FileDelete(datafil)
    FileAppend(JsonArray, datafil)
    MsgBox "GV-data er indlæst", "Done!"
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
;     MsgBox("Data er indlaest", "Vognløbsdata", "iconi")
; }

; gui_exit(a, x, y)
; {
;     MsgBox("1" a " 3" x)
;     ExitApp()
; }
