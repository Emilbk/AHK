#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


outlook_template := "C:\Users\ebk\Bod for kvalitetsbrist.oft"

outlook := ComObjCreate("Outlook.application")

bodtemplate := outlook.createitemfromtemplate(outlook_template)

test = 
(
Til
Reliable ApS
Bod for kvalitetsbrist
 
Midttrafik har den %A_now% 12.02.24 registreret en kvalitetsbrist på vognløb 31261, der medfører en bod på kr. 2.000,- jf. FG8, side 52, § 31, stk. 3, litra
 
17) Hvis garantivognen eller reservemateriel ikke er til rådighed i den aftalte garanti- og rådighedsperiode jf. § 20, stk. 3. Se dog § 25 stk. 2.
 
Kvalitetsbristen bestod i, at der ikke indsættes reservevogn på vognløbet, efter det lukkes ved vognløbets start.
 
Beløbet vil blive modregnet i vognmandsafregningen.
Eventuel indsigelse skal foretages skriftligt inden 5 arbejdsdage.
)



bodtemplate.To := "email@email.com"
bodtemplate.subject := "Bod for kvalitetsbrist - " vl " den " dato
bodtemplate.body := test


bodtemplate.display