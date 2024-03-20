#Requires Autohotkey v2
;AutoGUI creator: Alguimist autohotkey.com/boards/viewtopic.php?f=64&t=89901
;AHKv2converter creator: github.com/mmikeww/AHK-v2-script-converter
;EasyAutoGUI-AHKv2 github.com/samfisherirl/Easy-Auto-GUI-for-AHK-v2
#SingleInstance Force
;; XL
Excel := ComObject("Excel.Application")
ExcelDBStamdata := "F:\Flextrafik\Fælles\Udbud\FG8 - FV8\Stamoplysninger FV8 og FG8.xlsx"
ExcelDBSvigt := "F:\Flextrafik\Fælles\Udbud\Svigt\Svigt FG8-FV8.xlsx"
; Outlook
Outlook := ComObject("Outlook.Application")
signatur := A_ScriptDir "\lib\signatur_logo.png"
bodtemplate := outlook.createitem(0)

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
; DatoVaelg.OnEvent("Change", OnEventHandler)
ParagrafSoegFG.OnEvent("Change", FunkparagrafSoeg)
ParagrafSoegFV.OnEvent("Change", FunkparagrafSoeg)
ParagrafFG.OnEvent("Change", FunkParagrafVaelg)
ParagrafFV.OnEvent("Change", FunkParagrafVaelg)
; KvalitetsBrist.OnEvent()
; BodVaelg.OnEvent("Change", OnEventHandler)
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

FunkKnapOK(VD, *)
{
    valg := ""
    vl := ""
    bod := ""
    brist := ""
    paragraf := ""
    bod := ""
    dato := ""
    if FormatTime(DatoVaelg.Value, "ddMM") = FormatTime(A_Now, "ddMM")
        valg := MsgBox("Sikker på dags dato?", "Korrekt Dato?", "YN Icon!")
    if valg = "Yes"
        MsgBox "OK"
    GuiSubmit := mygui.Submit("Nohide")
    dato := formattime(Guisubmit.Datoresultat, "dd.MM.yyyy")
    Vl := GuiSubmit.VLResultat
    Bod := Guisubmit.Bod
    Kvalitetsbrist := GuiSubmit.Kvalitetsbrist
    if GuiSubmit.ParagrafFGResultat != ""
        Paragraf := GuiSubmit.ParagrafFGResultat
    if GuiSubmit.ParagrafFVResultat != ""
        Paragraf := GuiSubmit.ParagrafFVResultat
    msgbox("Bod er: " bod "`nVl er: " vl "`nKvalitetsbrist er: " Kvalitetsbrist "`nVM er: " VD[1] "`nKontaktinfo er: " VD[2] "`nParagraf er: " paragraf)
    VLSoeg.Focus()
    
    html_test_med_billede := 
(
    "<html xmlns:v=`"urn:schemas-microsoft-com:vml`" xmlns:o=`"urn:schemas-microsoft-com:office:office`" xmlns:w=`"urn:schemas-microsoft-com:office:word`" xmlns:m=`"http://schemas.microsoft.com/office/2004/12/omml`" xmlns=`"http://www.w3.org/TR/REC-html40`"><head><meta name=Generator content=`"Microsoft Word 15 (filtered medium)`"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
    o\:* {behavior:url(#default#VML);}
    w\:* {behavior:url(#default#VML);}
    .shape {behavior:url(#default#VML);}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
        {font-family:`"Cambria Math`";
        panose-1:2 4 5 3 5 4 6 3 2 4;}
    @font-face
        {font-family:Calibri;
        panose-1:2 15 5 2 2 2 4 3 2 4;}
    @font-face
        {font-family:Verdana;
        panose-1:2 11 6 4 3 5 4 4 2 4;}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
        {margin:0cm;
        font-size:11.0pt;
        font-family:`"Calibri`",sans-serif;
        mso-fareast-language:EN-US;}
    a:link, span.MsoHyperlink
        {mso-style-priority:99;
        color:#0563C1;
        text-decoration:underline;}
    p.Default, li.Default, div.Default
        {mso-style-name:Default;
        margin:0cm;
        text-autospace:none;
        font-size:12.0pt;
        font-family:`"Verdana`",sans-serif;
        color:black;}
    span.EmailStyle20
        {mso-style-type:personal-reply;
        font-family:`"Calibri`",sans-serif;
        color:windowtext;}
    .MsoChpDefault
        {mso-style-type:export-only;
        font-size:10.0pt;
        mso-ligatures:none;}
    @page WordSection1
        {size:612.0pt 792.0pt;
        margin:3.0cm 2.0cm 3.0cm 2.0cm;}
    div.WordSection1
        {page:WordSection1;}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext=`"edit`" spidmax=`"1026`" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext=`"edit`">
    <o:idmap v:ext=`"edit`" data=`"1`" />
    </o:shapelayout></xml><![endif]--></head><body lang=DA link=`"#0563C1`" vlink=`"#954F72`" style='word-wrap:break-word'><div class=WordSection1><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Til</span><span style='font-family:`"Verdana`",sans-serif'><o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif;mso-fareast-language:DA'>" VM "<o:p></o:p></span></p><p class=MsoNormal><b><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Bod for kvalitetsbrist<o:p></o:p></span></b></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Midttrafik har d. " dato " registreret en kvalitetsbrist på  vognløb <b>" VL ",</b> der medfører en bod på kr. " bod ",- jf. FG8, side 52, § 31, stk. 3, litra<o:p></o:p></span></p><p class=Default><o:p>&nbsp;</o:p></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>" paragraf "<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Kvalitetsbristen bestod i, at " brist "<o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'><o:p>&nbsp;</o:p></span></p><div><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Beløbet vil blive modregnet i vognmandsafregningen.<o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif'>Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif;mso-ligatures:standardcontextual;mso-fareast-language:DA'>Venlig hilsen<br><br>Flextrafiks Driftsafdeling<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-bottom-alt:auto'><span style='font-size:8.0pt;font-family:`"Verdana`",sans-serif;mso-ligatures:standardcontextual;mso-fareast-language:DA'><br>Flextrafik - Trafikstyring<br>&nbsp;<br>70 11 22 10<br><u><span style='color:blue'><a href=`"mailto:planet@Midttrafik.dk`">planet@Midttrafik.dk</a></span></u><br><br><span style='color:#9B1C3C'>Sender du fortrolige eller følsomme personoplysninger til Midttrafik, skal det ske via en sikker mailforbindelse. Se Midttrafiks <a href=`"https://www.midttrafik.dk/kundeservice/privatlivspolitik`"><span style='color:blue'>privatlivspolitik</span></a>.</span><o:p></o:p></span></p><p class=MsoNormal><a href=`"http://www.midttrafik.dk/`"><span style='font-family:`"Verdana`",sans-serif;color:blue;mso-fareast-language:DA;text-decoration:none'><img border=0 width=141 height=131 style='width:1.4687in;height:1.3645in' src=`"cid:signatur_logo.png`"></span></a><span style='font-size:10.0pt;font-family:`"Verdana`",sans-serif;mso-ligatures:standardcontextual'><o:p></o:p></span></p></div><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>



)"


bodtemplate := outlook.createitem(0)

bodtemplate.SentOnBehalfOfName := "planet@midttrafik.dk"
bodtemplate.To := email
bodtemplate.CC := "oekonomi@midttrafik.dk"
bodtemplate.subject := "Bod for kvalitetsbrist - vognløb " vl " d. " dato
bodtemplate.attachments.add(signatur)
bodtemplate.htmlbody := html_test_med_billede


bodtemplate.display
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