#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


#Include, %A_linefile%\..\..\Trafikstyring\lib\ImagePut-master\ImagePut (for v1).ahk
outlook_template := A_ScriptDir . "\lib\svigt_template.oft"

outlook := ComObjCreate("Outlook.application")

bodtemplate := outlook.createitemfromtemplate(outlook_template)

; billede := ImagePutWindow(ClipboardAll)
fil := ImagePutFile(clipboardall, "test.png")
fil_navn := SubStr(fil, 3)
fil_lok := A_ScriptDir "\" fil_navn
; MsgBox, , , %A_ScriptDir% %fil_lok%, 
html_test =
(
    </o:shapelayout></xml><![endif]--></head><body lang=DA link="#0563C1" vlink="#954F72" style='tab-interval:65.2pt;word-wrap:break-word'><div class=WordSection1><p class=MsoNormal>Billede start</p><p class=MsoNormal><span style='mso-ligatures:none'><img id="Billede_x0020_2" src="cid:%fil_navn%"></span></p><div><p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-size:10.0pt;font-family:"Verdana",sans-serif;mso-fareast-language:DA'>Billede slut<o:p></o:p></span></p></div><p class=MsoNormal><span style='font-size:10.0pt;font-family:"Verdana",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>
)
bodtemplate.To := "planet@midttrafik.dk"
bodtemplate.subject := "Bod for kvalitetsbrist - " vl " den " dato
bodtemplate.attachments.add(fil_lok)
bodtemplate.htmlbody := html_test


bodtemplate.display

ImageDestroy(fil)