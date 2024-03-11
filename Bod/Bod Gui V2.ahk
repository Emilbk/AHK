#Requires Autohotkey v2
;AutoGUI creator: Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter creator: github.com/mmikeww/AHK-v2-script-converter
;EasyAutoGUI-AHKv2 github.com/samfisherirl/Easy-Auto-GUI-for-AHK-v2
#SingleInstance Force
if A_LineFile = A_ScriptFullPath && !A_IsCompiled
    VM := ""
Email := ""
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
        ParagrafDataFG.Push(e[2])
    }
    if InStr(e[2], "FV", 1)
    {
        ParagrafDataFV.Push(e[2])
    }
}

myGui := Gui()
myGui.SetFont("s12 Bold", "Palatino Linotype")
myGui.Add("Text", "x16 y8 w55 h23 +0x200", "&VL")
VLSoeg := myGui.Add("Edit", "x16 y32 w120 h25")
myGui.Add("Text", "x192 y8 w209 h92", "VM:`n " VM "`n " Email "Kontaktinfo:")
myGui.SetFont("S10", "Palatino Linotype")
DatoVaelg := myGui.Add("DateTime", "x16 y64 w122 h24")
ParagrafSoegFG := myGui.Add("Edit", "x16 y132 w120 h26 vFG", "Søg FG")
ParagrafSoegFV := myGui.Add("Edit", "x156 y132 w120 h26 vFV", "Søg FV")
myGui.Add("GroupBox", "x8 y108 w394 h137", "Vælg &paragraf")
FGParagraf := myGui.Add("DropDownList", "x16 y164 w351", ParagrafDataFG)
FVParagraf := myGui.Add("DropDownList", "x16 y196 w351", ParagrafDataFV)
myGui.Add("Text", "x16 y248 w384 h51 +0x200", "Paragraf er")
myGui.Add("Text", "x16 y302 w120 h23 +0x200", "Bod")
BodVaelg := myGui.Add("Edit", "x16 y326 w120 h21", "1000")
myGui.Add("Text", "x16 y360 w221 h23 +0x200", "&Kvalitetsbristen bestod i, at...")
KvalitetsBrist := myGui.Add("Edit", "x16 y384 w373 h99", "Kvalitetsbrist")
ButtonOK := myGui.Add("Button", "x173 y496 w95 h27", "&OK")
VLSoeg.OnEvent("LoseFocus", FunkVLSoeg)
DatoVaelg.OnEvent("Change", OnEventHandler)
ParagrafSoegFG.OnEvent("Change", FunkparagrafSoeg)
ParagrafSoegFV.OnEvent("Change", FunkparagrafSoeg)
FGParagraf.OnEvent("Change", FunkParagrafVaelg)
FVParagraf.OnEvent("Change", FunkParagrafVaelg)
; KvalitetsBrist.OnEvent()
BodVaelg.OnEvent("Change", OnEventHandler)
ButtonOK.OnEvent("Click", OnEventHandler)
myGui.OnEvent('Close', (*) => ExitApp())
myGui.Title := "Ny Optimeret Bodsudskriver"

MyGui.Show("W442 H544")

FunkParagrafVaelg(*)
{
    ; MsgBox(ParagrafSoeg.Value)
}
FunkParagrafSoeg(AktivControl, *)
{
    if AktivControl.Name = "FG"
    {
        FVParagraf.Choose(0)
        ParagrafSoegFV.Value := ""
        if ParagrafSoegFG.Value = ""
        {
            FGParagraf.Choose(0)
            return
        }
        for i, e in ParagrafDataFG
        {
            if InStr(e, ParagrafSoegFG.Value, , 4)
            {
                FGParagraf.Choose(i)
                FVParagraf.Choose(0)
            }
        }
    }
    if AktivControl.Name = "FV"
    {
        FGParagraf.Choose(0)
        ParagrafSoegFG.Value := ""
    if ParagrafSoegFV.Value = ""
    {
        FVParagraf.Choose(0)
        return
    }
    for i, e in ParagrafDataFV
    {
        if InStr(e, ParagrafSoegFV.Value)
        {
            FVParagraf.Choose(i)
            FGParagraf.Choose(0)
        }
    }
    }
    penge := 2000
    BodVaelg.Value := penge " ,-"

    Return
}
FunkVLSoeg(*)
{
    if VLSoeg.Value != ""
        MsgBox(VLSoeg.Value)
    return
}

OnEventHandler(*)
{
    ToolTip("Click! This is a sample action.`n"
        . "Active GUI element values include:`n"
        . "Edit1 => " VLSoeg.Value "`n"
        . "DateTime1 => " DatoVaelg.Value "`n"
        . "DropDownList1 => " FGParagraf.Text "`n"
        . "DropDownList2 => " FVParagraf.Text "`n"
        . "Edit2 => " KvalitetsBrist.Value "`n"
        . "Edit3 => " BodVaelg.Value "`n"
        . "Edit4 => " ParagrafSoegFG.Value "`n"
        . "ButtonOK => " ButtonOK.Text "`n", 77, 277)
    SetTimer () => ToolTip(), -3000 ; tooltip timer
}