#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


outlook := ComObjCreate("Outlook.application")

;; Fileinstall

FileInstall, lib/bod_template.oft, temp.oft
;; GUI

brugsanvisning =
(
Vogn og vognløb indskrives. Sendes der bod på flere vognløb af gangen deles de op med "," (5023, 5052, osv.)

Størrelsen på boden udregnes automatisk, med udgangspunkt i 1000 kr. pr. vognløb. Bodsstørrelsen kan ændres manuelt.

)   
bod :=
bod_kr := 1000


Gui vl_bod: Font, s9, Segoe UI

gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x18 y17 w35 h23 +0x200, Vogn:
gui vl_bod: add, text, x18 y42 w35 h23 +0x200, VL:
gui vl_bod: add, text, x245 y15 w100 h29 +0x200, Brugsanvisning
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x62 y22 w120 h21 number vvogn 
gui vl_bod: add, edit, x62 y47 w120 h21 vvl gvl_tael
gui vl_bod: add, text, x245 y40 w300 , %Brugsanvisning%
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
; gui vl_bod: add, text, x18 y151 h23 +0x200, Eventuelt andre datoer:
gui vl_bod: add, text, x62 y151 h23 +0x200, Bod:
gui vl_bod: add, dateTime, vdato x18 y82 w164 h60, 
; gui vl_bod: add, edit, x62 y173 w120 h21 vny_dato gny_dato 
gui vl_bod: add, edit, x62 y173 w120 h21 number vbod, %bod%
; gui vl_bod: add, text, x368 y176 h23 +0x200 vbod_antal gbod_antal, Én bod for hvert vognløb, hver dag:
; gui vl_bod: add, Checkbox, x348 y176 w10 h21 
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: font
gui vl_bod: font, bold
gui vl_bod: add, text, x25 y200 w260 h23 +0x200, &Kvalitetsbristen bestod i, at...
gui vl_bod: font
gui vl_bod: font, s9, segoe ui
gui vl_bod: add, edit, x25 y228 w568 h84 vbrist, der i den faste chaufførs fravær ikke indsættes den faste afløser til udførsel af kørslen.
gui vl_bod: add, button, default x288 y315 gbod_ok, &OK
gui vl_bod: show, w620 h342, Ny Århus-taxa Bod

vl_tael:
{
    gui submit, nohide
    if (not InStr(vl, ","))
        {
         vl_antal := 1   
         bod := vl_antal * bod_kr
         GuiControl, , bod, %bod%
         return
        }
    vl_antal := 0
    vl_array := StrSplit(vl, ",")

    for i,e in vl_array
        {
            if (e != "")
                {
                    vl_antal += 1
                    bod := vl_antal * bod_kr
                    GuiControl, , bod, %bod%

                }
        }

    return

}
ny_dato:
{
    gui submit, NoHide
    if (not InStr(ny_dato, ","))
        {
         dag_antal := 1   
         return
        }
    dag_antal := 0
    ny_dato_array := StrSplit(ny_dato, ",")
    for i,e in ny_dato_array
        {
            if (e != "")
                {
                    dag_antal += 1
                    bod := dag_antal * (vl_antal * 1000)
                    GuiControl, , bod, %bod%
                }
        }
return
}

; bod_ok:
; {
;     MsgBox, , , %vl_antal% - %dag_antal%
;     return
; }

;; OUTLOOK

outlook_template := "C:\Users\ebk\Bod for kvalitetsbrist.oft"

outlook := ComObjCreate("Outlook.application")
; bodtemplate := outlook.createitemfromtemplate(outlook_template)


;; STAMDATA

+esc::
{
    ; stamopl.quit()
    ExitApp
}
return
outlook_ryd:
guicontrol, vl_bod: , vl , 
guicontrol, vl_bod: , combobox1 , 
guicontrol, vl_bod: , combobox2 , 
guicontrol, vl_bod: , edit3 ,

;; GUI-label
vl_bodguiescape:
vl_bodguiclose:
    {
    FileDelete, temp.oft
    ExitApp
    }

bod_ok:
gui Submit, nohide

for i,e in vl_array
    {
    vl_array[i] := RegExReplace(vl_array[i], "\D")
    if (if A_Index = 1)
    {
        vl := vl_array[i]
    }
    else if (A_Index = vl_array.MaxIndex())
        {
            vl .= " og " vl_array[i]
        }
    else
        {
            vl .= ", " vl_array[i]
        }
    }
FormatTime, dato, %dato%, dd.MM.yyyy
vm := "Foreningen Taxa f.m.b.a."
paragraf := "d. Manglende overholdelse af pkt. 19 – kørslens udførelse, se dog litra g, i og k"


html_tekst =
(
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head>
<meta http-equiv="Content-Type" content="text/html; charset=Windows-1252"><meta name="Generator" content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
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
@font-face
	{font-family:Aptos;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;
	mso-ligatures:standardcontextual;
	mso-fareast-language:EN-US;}
span.EmailStyle18
	{mso-style-type:personal-compose;
	font-family:"Verdana",sans-serif;
	color:windowtext;}
p.Default, li.Default, div.Default
	{mso-style-name:Default;
	margin:0cm;
	text-autospace:none;
	font-size:12.0pt;
	font-family:"Verdana",sans-serif;
	color:black;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;
	mso-ligatures:none;
	mso-fareast-language:EN-US;}
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
</o:shapelayout></xml><![endif]--></head><body lang="DA" link="#0563C1" vlink="#954F72" style="word-wrap:break-word"><div class="WordSection1"><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Til<br></span><span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;mso-fareast-language:DA">%vm%</span><o:p></o:p></p><p class="MsoNormal"><b><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Bod for kvalitetsbrist</span></b><o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">&nbsp;</span><o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Midttrafik har den %dato% registreret en kvalitetsbrist på vogn <b>%vogn%</b>, vognløb <b>%vl%</b>, der medfører en bod på kr. %bod%,- jf. FS11, side 25, §26, stk. 1.7., litra</span><o:p></o:p></p><p class="Default">&nbsp;<o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">d. Manglende overholdelse af pkt. 19 — kørslens udførelse, se dog litra g, i og k.</span><o:p></o:p></p><p class="MsoNormal">&nbsp;<o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Kvalitetsbristen bestod i, at %brist%</span><o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">&nbsp;</span><o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Beløbet vil blive modregnet i vognmandsafregningen.</span><o:p></o:p></p><p class="MsoNormal"><span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.</span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>
)

; bodtemplate := outlook.createitem(0)
outlook_template := A_ScriptDir . "\lib\bod_template.oft"
bodtemplate := outlook.createitemfromtemplate(outlook_template)


bodtemplate.SentOnBehalfOfName := "specialkoersel@midttrafik.dk"
bodtemplate.To := "ouv@aarhus-taxa.dk"
bodtemplate.CC := "oekonomi@midttrafik.dk"
bodtemplate.subject := "Bod for kvalitetsbrist - vogn " vogn " d. " dato
; bodtemplate.attachments.add(signatur)
bodtemplate.htmlbody := html_tekst


bodtemplate.display

guicontrol, vl_bod: , vogn , 
guicontrol, vl_bod: , vl , 
guicontrol, vl_bod: , edit3 , 
guicontrol, vl_bod: choose , combobox1 , 1
guicontrol, vl_bod: choose , combobox2 , 1
return

