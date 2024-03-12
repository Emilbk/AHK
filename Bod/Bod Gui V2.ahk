#Requires Autohotkey v2
;AutoGUI creator: Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter creator: github.com/mmikeww/AHK-v2-script-converter
;EasyAutoGUI-AHKv2 github.com/samfisherirl/Easy-Auto-GUI-for-AHK-v2
#SingleInstance Force
;; XL
Excel := ComObject("Excel.Application")
ExcelDBStamdata := A_ScriptDir "\db\Stamoplysninger FV8 og FG8.xlsx"
ExcelDBSvigt := A_ScriptDir "\db\Svigt FG8-FV8.xlsx"
; Excel.Visible := true
ExcelSvigtWorkbok := Excel.Workbooks.Open(ExcelDBSvigt, , readonly := True)
ExcelSvigtWorksheet := ExcelSvigtWorkbok.worksheets.item("Vognløbsdata")
ExcelSvigtWorksheet.Select
; MsgBox Excel.ActiveSheet.Name
ExcelSvigtInputDrift := Excel.Intersect(Excel.Columns("A:B"), excel.Activesheet.UsedRange).value

StamData := []
loop ExcelSvigtInputDrift.MaxIndex(1)
{
    StamData.push([ExcelSvigtInputDrift[A_index, 1], ExcelSvigtInputDrift[A_index, 2]])
}
for i, e in StamData
{
    if i > 1
    {
        e[1] := Format("{:.0i}", e[1])
        e[1] := Format("{:s}", e[1])
    }
    ; RouOnd(e[1])
    ; e[1] := String(e[1])
}
; Excel.Visible := true
ExcelSvigtWorkbok := Excel.Workbooks.Open(ExcelDBStamdata, , readonly := True)
; ExcelSvigtWorksheet := ExcelSvigtWorkbok.worksheets.item("Vognløbsdata")
; ExcelSvigtWorksheet.Select
; MsgBox Excel.ActiveSheet.Name
ExcelStamdataInput := Excel.Intersect(Excel.Columns("A:P"), excel.Activesheet.UsedRange).value
StamdataVm := []
loop ExcelStamdataInput.MaxIndex(1)
{
    StamdataVm.push([ExcelStamdataInput[A_index, 1], ExcelStamdataInput[A_index, 12]])
}
for i_vm, e_vm in StamdataVm
for i_svigt, e_svigt in StamData
    if e_svigt[2] = e_vm[1]
        StamData[i_svigt].Push(e_vm[2])
        ; MsgBox  e[1] "og" e[2]

VM := ""
Email := ""
VmData := ["test"]
test := ""
ParagrafDataListboxFG := []
ParagrafDataListboxFV := []
ParagrafDataFG := []
ParagrafDataFV := []
ParagrafDataInd := FileRead("db\paragraf_data.txt")
ParagrafDataInd := StrReplace(ParagrafDataInd, "`n", "")
ParagrafDataArray := StrSplit(ParagrafDataInd, "`r")
ParagrafDataArray.RemoveAt(ParagrafDataArray.length)
for i, e in ParagrafDataArray
{
    ParagrafDataArray[i] := StrSplit(e, "`t")
}
for i, e in ParagrafDataArray
{
    if InStr(e[2], "FG", 1)
    {
        ParagrafDataListboxFG.Push(e[2])
        ParagrafDataFG.Push(e)
        if ParagrafDataFG[i].Length < 4
            ParagrafDataFG.Push()
    }
    if InStr(e[2], "FV", 1)
    {
        ParagrafDataListboxFV.Push(e[2])
        ParagrafDataFV.Push(e)
    }
}

myGui := Gui()
myGui.VmData := ""
myGui.SetFont("s12 Bold", "Palatino Linotype")
myGui.Add("Text", "x16 y8 w55 h23 +0x200", "&VL")
VmInfo := myGui.Add("Text", "x192 y8 w209 h92", "VM:`n " VM "`n " Email "Kontaktinfo:")
myGui.SetFont("S10 Norm", "Palatino Linotype")
VLSoeg := myGui.Add("Edit", "Number x16 y32 w120 h25 VVLResultat")
DatoVaelg := myGui.Add("DateTime", "x16 y64 w122 h24 VDatoResultat")
ParagrafSoegFG := myGui.Add("Edit", "x16 y132 w120 h26 vParagrafFGSoeg", "Søg FG")
ParagrafSoegFV := myGui.Add("Edit", "x156 y132 w120 h26 vParagrafFVSoeg", "Søg FV")
myGui.Add("GroupBox", "x8 y108 w394 h137", "Vælg &paragraf")
ParagrafFG := myGui.Add("DropDownList", "x16 y164 w351 vParagrafFGResultat", ParagrafDataListboxFG)
ParagrafFV := myGui.Add("DropDownList", "x16 y196 w351 vParagrafFVResultat", ParagrafDataListboxFV)
ParagrafTekst := myGui.Add("Text", "x16 y248 w384 h51 VParagrafTekst", "")
myGui.Add("Text", "x16 y302 w120 h23 +0x200", "Bod:")
BodVaelg := myGui.Add("Edit", "x16 y326 w120 h21 VBod", "1000")
myGui.Add("Text", "x16 y360 w221 h23 +0x200", "&Kvalitetsbristen bestod i, at...")
Kvalitetsbrist := myGui.Add("Edit", "x16 y384 w373 h99 VKvalitetsbrist", "Kvalitetsbrist")
ButtonOK := myGui.Add("Button", "x173 y496 w95 h27", "&OK")
VLSoeg.OnEvent("LoseFocus", (*) => (mygui.VmData := FunkVLSoeg()))
DatoVaelg.OnEvent("Change", OnEventHandler)
ParagrafSoegFG.OnEvent("Change", FunkparagrafSoeg)
ParagrafSoegFV.OnEvent("Change", FunkparagrafSoeg)
ParagrafFG.OnEvent("Change", FunkParagrafVaelg)
ParagrafFV.OnEvent("Change", FunkParagrafVaelg)
; KvalitetsBrist.OnEvent()
BodVaelg.OnEvent("Change", OnEventHandler)
ButtonOK.OnEvent("Click", (*) => FunkKnapOK(mygui.Vmdata))
myGui.OnEvent('Close', (*) => ExitApp())
myGui.Title := "Ny Optimeret Bodsudskriver"

VLSoeg.Focus()
MyGui.Show("W442 H544")


FunkParagrafVaelg(AktivControl, *)
{
    if AktivControl = ParagrafFG
    {
        ParagrafTekst.Text := ParagrafDataFG[ParagrafFG.Value][3]
        BodVaelg.Value := ParagrafDataFG[ParagrafFG.Value][1]
        KvalitetsBrist.Value := ParagrafDataFG[ParagrafFG.Value][4]
        ParagrafFV.Choose(0)
    }
    if AktivControl = ParagrafFV
    {
        ParagrafTekst.Text := ParagrafDataFV[ParagrafFV.Value][3]
        BodVaelg.Value := ParagrafDataFV[ParagrafFV.Value][1]
        KvalitetsBrist.Value := ParagrafDataFV[ParagrafFV.Value][4]
        ParagrafFG.Choose(0)
    }
    return
}
FunkParagrafSoeg(AktivControl, *)
{
    if AktivControl.Name = "ParagrafFGSoeg"
    {
        ParagrafFV.Choose(0)
        ParagrafSoegFV.Value := ""
        if ParagrafSoegFG.Value = ""
        {
            ParagrafFG.Choose(0)
            return
        }
        for i, e in ParagrafDataListboxFG
        {
            if InStr(e, ParagrafSoegFG.Value, , 4)
            {
                ParagrafFG.Choose(i)
                ParagrafFV.Choose(0)
                ParagrafTekst.Text := ParagrafDataFG[i][3]

            }
        }
    }
    if AktivControl.Name = "ParagrafFVSoeg"
    {
        ParagrafFG.Choose(0)
        ParagrafSoegFG.Value := ""
        if ParagrafSoegFV.Value = ""
        {
            ParagrafFV.Choose(0)
            return
        }
        for i, e in ParagrafDataListboxFV
        {
            if InStr(e, ParagrafSoegFV.Value)
            {
                ParagrafFV.Choose(i)
                ParagrafFG.Choose(0)
                ParagrafTekst.Text := ParagrafDataFV[i][3]
            }
        }
    }

    Return
}
FunkVLSoeg(*)
{
    fundet := 0
    if VLSoeg.Value != ""
        {
            for i,e in StamData
                {
                    if VLSoeg.Value = StamData[i][1]
                        {
                            VM := Stamdata[i][2]
                            Email := Stamdata[i][3]
                            VmInfo.Text := "VM:`n" VM "`nKontaktinfo:`n" Email
                            fundet := 1
                        }
                }
        }
    if fundet = 0
        {
        VmInfo.Value := "Ikke et gyldigt vognløb."
        return 0
        }
    ud := [VM, Email]
    return ud
}

OnEventHandler(*)
{
    ToolTip("Click! This is a sample action.`n"
        . "Active GUI element values include:`n"
        . "Edit1 => " VLSoeg.Value "`n"
        . "DateTime1 => " DatoVaelg.Value "`n"
        . "DropDownList1 => " ParagrafFG.Text "`n"
        . "DropDownList2 => " ParagrafFV.Text "`n"
        . "Edit2 => " KvalitetsBrist.Value "`n"
        . "Edit3 => " BodVaelg.Value "`n"
        . "Edit4 => " ParagrafSoegFG.Value "`n"
        . "ButtonOK => " ButtonOK.Text "`n", 77, 277)
    SetTimer () => ToolTip(), -3000 ; tooltip timer
}
FunkKnapOK(VD, *)
{
    valg := ""
    vl := ""
    bod := ""
    Kvalitetsbrist := ""
    paragraf := ""
    if FormatTime(DatoVaelg.Value, "ddMM") = FormatTime(A_Now, "ddMM")
        valg := MsgBox("Sikker på dags dato?", "Korrekt Dato?", "YN Icon!")
    if valg = "Yes"
        MsgBox "OK"
    GuiSubmit := mygui.Submit()
    msgbox FormatTime(Guisubmit.Datoresultat, "dd.MM.yyyy")
    Vl := GuiSubmit.VLResultat
    Bod := Guisubmit.Bod
    Kvalitetsbrist := GuiSubmit.Kvalitetsbrist
    if GuiSubmit.ParagrafFGResultat != ""
        Paragraf := GuiSubmit.ParagrafFGResultat
    if GuiSubmit.ParagrafFVResultat != ""
        Paragraf := GuiSubmit.ParagrafFVResultat
    for i,e in GuiSubmit.OwnProps()
        if e = ""
            MsgBox(GuiSubmit.i " er tom.")
    msgbox("Bod er: " bod "`nVl er: " vl "`nKvalitetsbrist er: " Kvalitetsbrist "`nVM er: " VD[1] "`nKontaktinfo er: " VD[2] "`nParagraf er: " paragraf)
    VLSoeg.Focus()
    return 
}
#HotIf WinActive("Ny")
!g::
{
    ParagrafSoegFG.Focus
    return
}
!f::
{
    ParagrafSoegFV.Focus
    return
}