#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
;; GUI

gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x18 y22 w35 h23 +0x200, vl:
gui vl_bod: add, text, x285 y10 w120 h23 +0x200, vm:
gui vl_bod: add, text, x285 y54 w120 h23 +0x200, kontaktinfo:
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x62 y27 w120 h21 number vvl gvl_slaa_op
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, text, x380 y10 w200 h23 vvm +0x200, % vm
gui vl_bod: add, text, x380 y54 w200 h23 vemail +0x200, % email
gui vl_bod: add, dateTime, vdato x20 y58 w164 h60, 
; gui vl_bod: add, monthcal, x20 y58 w164 h160 vdato
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, text, x24 y170 w120 h23 +0x200, &Søg Paragraf
gui vl_bod: add, edit, x24 y200 w120 h23 +0x200 vparagraf_søg gparagraf_slaa_op, 
gui vl_bod: add, text, x24 y223 w120 h23 +0x200, &Paragraf
gui vl_bod: add, dropdownlist, x23 y244 w414 vFG, 
gui vl_bod: add, dropdownlist, x23 y270 w414 vFV,
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x25 y300 w260 h23 +0x200, &kvalitetsbristen bestod i, at...
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x25 y328 w568 h84 vbrist
gui vl_bod: add, button, default x288 y415, &ok

fileread, paragraf_data, db/paragraf_data.txt
paragraf_ny := []
paragraf_data := strreplace(paragraf_data, "`r", "")
paragraf_data := strsplit(paragraf_data, "`n")
for i,e in paragraf_data
    {
    paragraf_ny[i] := strsplit(e, "`t")
    }
paragraf_data := paragraf_ny

;; paragraf_drop_down
; msgbox, , , % substr(paragraf_data[1][1], 1,2)
paragraf_drop_down_fg := "-|"
paragraf_drop_down_fv := "-|"
for i,e in paragraf_data
    if (substr(e[2], 1 ,2) = "fg")
    {
    paragraf_drop_down_fg .= paragraf_data[i][2] "|"

    }
for i,e in paragraf_data
    if (substr(e[2], 1 ,2) = "fv")
    {
    paragraf_drop_down_fv .= paragraf_data[i][2] "|"

    }

;; OUTLOOK
outlook_template := "C:\Users\ebk\Bod for kvalitetsbrist.oft"

outlook := ComObjCreate("Outlook.application")

bodtemplate := outlook.createitem(0)
; bodtemplate := outlook.createitemfromtemplate(outlook_template)


;; STAMDATA

stamopl_sti := "C:\Users\ebk\Stamoplysninger FV8 og FG8.xlsx"

stamopl:= ComObjCreate("Excel.application")
; stamopl.Workbooks.Open(stamopl_sti,, readonly := false)
; stamopl_workbook := 
stamopl_workbook := stamopl.workbooks.open(stamopl_sti,, readonly := true)
stamopl.visible := 0
stamopl_ark := stamopl.sheets("ark1") ; after opening workbook its better to define sheet 
stamopl_kolonne_a := stamopl_ark.range("A:A") 
stamopl_kolonne_b := stamopl_ark.range("B:B") 
r_a_sidste := stamopl_kolonne_a.end(-4121).row
r_b_sidste := stamopl_kolonne_b.end(-4121).row
vm_stam := []    
kontakt_stam := []
loop, %r_a_sidste%
    {
         vm_stam.push(stamopl_ark.range("A" A_index).value)
         kontakt_stam.push(stamopl_ark.range("L" A_index).value)
    }
vm_stam.RemoveAt(1)
kontakt_stam.RemoveAt(1)
stamopl.quit()

stamopl_sti := "C:\Users\ebk\Svigt FG8-FV8.xlsx"

stamopl:= ComObjCreate("Excel.application")
; stamopl.Workbooks.Open(stamopl_sti,, readonly := false)
; stamopl_workbook := 
stamopl_workbook := stamopl.workbooks.open(stamopl_sti,, readonly := true)
stamopl.visible := 0
stamopl_ark := stamopl.sheets(4) ; after opening workbook its better to define sheet 
stamopl_kolonne_a := stamopl_ark.range("A:A") 
stamopl_kolonne_b := stamopl_ark.range("B:B") 
r_a_sidste := stamopl_kolonne_a.end(-4121).row
r_b_sidste := stamopl_kolonne_b.end(-4121).row

vl_svigt := []
vm_svigt := []
email_svigt := []
fundet := []
loop, %r_a_sidste%
    {
         vl_svigt.push(stamopl_ark.range("A" A_index).value)
         vm_svigt.push(stamopl_ark.range("B" A_index).value)
    }
stamopl.quit()
for i,e in vl_svigt
    {
        if e is number
            {
                vl_svigt[i] := Format("{:d}", e)
                
            }
    
    }
stamdata := []
for i,e in vl_svigt
    {
        stamdata[i] := [vl_svigt[i], vm_svigt[i]]
    }    

for i, e in vm_svigt
    {
        for i2, e2 in vm_stam
            if (e = e2)
                {
                stamdata[i].Push(kontakt_stam[i2])
                Break 1
                }
    }


; for i,e in stamdata
;     {
;         MsgBox, , , % "Vl " stamdata[i][1] " tilhører " stamdata[i][2] ", som har email " stamdata[i][3]
;     }
; ; stamopl_ark :½= stamopl_workbook.worksheets("Ark1")
; test := stamopl.worksheets(stamopl_ark).columns(1)
; stamopl.workbooks()
; vm := stamopl_ark.range("A:A").end("xldown")

guicontrol, vl_bod: , combobox1 , %paragraf_drop_down_fg%
guicontrol, vl_bod: , combobox2 , %paragraf_drop_down_fv%
guicontrol, vl_bod: choose, combobox1, 1
guicontrol, vl_bod: choose, combobox2, 1


gui vl_bod: show, w620 h442, window

+esc::
{
    stamopl.quit()
    ExitApp
}
; oWorkbook := ComObjCreate("Excel.Application")
; oWorkbook.Workbooks.open(FilePath,, readonly := true)
; oWorkbook.Visible := 0 
; clientsname := oWorkbook.Worksheets("test doc").Range("A3").Value
; StringRight, clientsname, clientsname, 5
; clientsphone := oWorkbook.Worksheets("test doc").Range("B3").Value
; clientsstate := oWorkbook.Worksheets("test doc").Range("C3").Value
; clientsfax := oWorkbook.Worksheets("test doc").Range("D3").Value


;; GUI-funktion

paragraf_slaa_op:
{
    guicontrolget, vl, , edit2, 
    if (vl = "")
        guicontrolget, vl, , edit2, 
    for i,e in paragraf_data
        {
            if (InStr(e[2], fg))
                {
                    guicontrol, vl_bod: , fg , % paragraf_data[i][2]
                    break
                    return
                }
            if (InStr(e[2], fv))
                {
                    guicontrol, vl_bod: , fv , % paragraf_data[i][2]
                    break
                    return
                }
            else
                {

                }
        }
return
}

vl_slaa_op:
{
    guicontrolget, vl, , edit1, 
    if (vl = "")
        guicontrolget, vl, , edit1, 
    for i,e in stamdata
        {
            if (e[1] = vl)
                {
                    vm := e[2]
                    email := e[3]
                    guicontrol, vl_bod: , vm , % vm
                    guicontrol, vl_bod: , email , % email
                    return
                }
            else
                {

                    guicontrol, vl_bod: , vm , ikke gyldigt vl
                    guicontrol, vl_bod: , email ,  
                }
        }
return
}
;; GUI-label
vl_bodguiescape:
vl_bodguiclose:
    {
    MsgBox, , , quit
    stamopl.quit()
    ExitApp
    }
vl_bodbuttonok:
gui Submit, nohide
if (fg != "-" and fv != "-")
    {
        MsgBox, 16, Både FG og FV valgt, Der skal kun vælges fra ét udbud.
        return
    }
    for i,e in paragraf_data
        {
            if (e[2] = fg) or (e[2] = fv) 
                {
                    paragraf := paragraf_data[i][3]
                    bod := paragraf_data[i][1]
                    break
                }
        }


FormatTime, dato, %dato%, dd.MM.yyyy

test = 
(
Til
%vm%
Bod for kvalitetsbrist
 
Midttrafik har den %dato% registreret en kvalitetsbrist på vognløb %vl%, der medfører en bod på kr. %bod%,- jf. FG8, side 52, § 31, stk. 3, litra 

%paragraf%
 
Kvalitetsbristen bestod i, at %brist%
 
Beløbet vil blive modregnet i vognmandsafregningen.
Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.

Venlig hilsen

Flextrafiks Driftcenter
• 70112210
 
planet@midttrafik.dk

Sender du fortrolige eller følsomme personoplysninger til Midttrafik, skal det ske via en sikker mailforbindelse. Se Midttrafiks privatlivspolitik.
 

)

; MsgBox, , , %test%


html_test =
(
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
    o\:* {behavior:url(#default#VML);}
    w\:* {behavior:url(#default#VML);}
    .shape {behavior:url(#default#VML);}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
        {font-family:"Cambria Math";
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
        font-family:"Calibri",sans-serif;
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
        font-family:"Verdana",sans-serif;
        color:black;}
    span.EmailStyle20
        {mso-style-type:personal-reply;
        font-family:"Calibri",sans-serif;
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
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]--></head><body lang=DA link="#0563C1" vlink="#954F72" style='word-wrap:break-word'><div class=WordSection1><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Til</span><span style='font-family:"Verdana",sans-serif'><o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-fareast-language:DA'>%VM%<o:p></o:p></span></p><p class=MsoNormal><b><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Bod for kvalitetsbrist<o:p></o:p></span></b></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Midttrafik har d. %dato% registreret en kvalitetsbrist på  vognløb <b>%VL%,</b> der medfører en bod på kr. %bod%,- jf. FG8, side 52,   31, stk. 3, litra<o:p></o:p></span></p><p class=Default><o:p>&nbsp;</o:p></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>%paragraf%<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Kvalitetsbristen bestod i, at %brist%<o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p><div><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Beløbet vil blive modregnet i vognmandsafregningen.<o:p></o:p></span></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'>Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-ligatures:standardcontextual;mso-fareast-language:DA'>Venlig hilsen<br><br>Flextrafiks Driftsafdeling<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-bottom-alt:auto'><span style='font-size:8.0pt;font-family:"Verdana",sans-serif;mso-ligatures:standardcontextual;mso-fareast-language:DA'><br>Flextrafik - Trafikstyring<br>&nbsp;<br>70 11 22 10<br><u><span style='color:blue'><a href="mailto:planet@Midttrafik.dk">planet@Midttrafik.dk</a></span></u><br><br><span style='color:#9B1C3C'>Sender du fortrolige eller f lsomme personoplysninger til Midttrafik, skal det ske via en sikker mailforbindelse. Se Midttrafiks <a href="https://www.midttrafik.dk/kundeservice/privatlivspolitik"><span style='color:blue'>privatlivspolitik</span></a>.</span><o:p></o:p></span></p><p class=MsoNormal><a href="http://www.midttrafik.dk/"><span style='font-family:"Verdana",sans-serif;color:blue;mso-fareast-language:DA;text-decoration:none'><img border=0 width=141 height=131 style='width:1.4687in;height:1.3645in' id="Billede_x0020_1" src="cid:image001.png@01DA0104.C6BC5820"></span></a><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-ligatures:standardcontextual'><o:p></o:p></span></p></div><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>



)




bodtemplate.SentOnBehalfOfName := "planet@midttrafik.dk"
bodtemplate.To := email
bodtemplate.CC := "oekonomi@midttrafik.dk"
bodtemplate.subject := "Bod for kvalitetsbrist - vognløb " vl " d. " dato
bodtemplate.htmlbody := html_test


bodtemplate.display

guicontrol, vl_bod: , vl , 
return