
  gui, vmsvigt: New, , vmsvigt
  gui, vmsvigt: add, text, , er vognmanden kontaktet?
  gui, vmsvigt: add, button, default, Kontaktet
  gui, vmsvigt: add, text, ,sdfljsdf
  gui, vmsvigt: add, text, ,sdfljsdf


Menu, SvigtFilmenu, add, Vogngruppe`tCtrl+t, p6_vgsvigt  
Menu, SvigtSkærmprintMenu, add, FlexFinder-skærmprint`tCtrl+f, flexfinderskærmprint  
Menu, SvigtSkærmprintMenu, add, Tilføj nuværende skærmprint`tCtrl+n, nuværendeSkærmprint  
Menu, SvigtSkærmprintMenu, add, Vis tilføjede skærmprint`tCtrl+v, SvigtSkærmprintOversigt  
Menu, SvigtOmmenu, add, Hjælp`tF1, svigtHjælp  
Menu, SvigtVGMenu, add, Opret Vogngruppesvigt`tCtrl+t, svigtVogngruppe  
Menu, SvigtMenu, add, &Vogngruppe, :SvigtVGMenu,
Menu, SvigtMenu, add, S&kærmprint, :SvigtSkærmprintMenu,
Menu, SvigtMenu, add, Hjælp, :SvigtOmMenu, +right

gui, svigt: new
gui, svigt: +labelsvigt
Gui svigt: Add, Text, x16 y0 w120 h23 +0x200, Vognløbs&nummer
Gui svigt: Font
Gui svigt: Add, Text, x16 y53 h35 w100 vgarantitid, Garantiperiode: %garanti_tid%
Gui svigt: Font, w600
Gui svigt: Add, Text, x16 y73 w120 , Garanti eller Var.
Gui svigt: Font
Gui svigt: Add, Edit, vVL x16 y24 w120 h21, %vl%
Gui svigt: Font, s9, Segoe UI
; Gui svigt: Font, w600
Gui, svigt:Add, GroupBox,  x150 y0 w140 h130 ,Hvis GV &lukket:
; Gui svigt: Add, Text, x161 y0 w130 h25 +0x200, Hvis GV &lukket:
Gui svigt: Font
Gui svigt: Font, s9, Segoe UI
Gui svigt: Add, Radio, våbningstidradio x160 y25 , &Åbningstid udskudt
Gui svigt: Add, Edit, disabled våbningstidedit x180 y40 w79 h21, Vl start kl.
Gui svigt: Add, Radio, vlukket x160 y65 , Lukket &midt på VL
Gui svigt: Add, Edit, disabled vtid x180 y80 w79 h21, Hjemzone kl.
Gui svigt: Add, Radio, vsvigtVlSlettetRadio x160 y105 , VL S&lettet
; Gui svigt: Add, CheckBox, vlukket x160 y19 w39 h23, &Ja
; Gui svigt: Add, CheckBox, vhelt x160 y38 w115 h23, &Ja, og VL slettet
; Gui svigt: Add, CheckBox, vvmKontakt x160 y60 w120 h23, V&M kontaktet ca. kl.
; Gui svigt: Add, Edit, vvmKontaktTid x200 y80 w79 h21, %tidforsvigt%
Gui svigt: Font
Gui svigt: Font, s9, Segoe UI
; Gui svigt: Font, w600
Gui svigt: Add, Text, x16 y95 w120 h23 +0x200, &Årsag
Gui svigt: Font
Gui svigt: Add, Edit, vårsag x16 y120 w120 h21
Gui svigt: Font, s9, Segoe UI
Gui svigt: Font, w600
Gui svigt: Font
Gui svigt: Font, s9, Segoe UI
Gui, svigt:Add, GroupBox,  x294 y0 w140 h130 ,Type VL:
Gui svigt: Add, Radio, x304 y24 w120 h16, &Garanti
Gui svigt: Add, Radio, x304 y40 w120 h32, G&arantivognløb i variabel tid
Gui svigt: Add, Radio, x304 y72 w120 h23, Va&riabel
Gui svigt: Add, Radio, disabled vtype x304 y92 w120 h32, V&ogngruppe
Gui svigt: Add, GroupBox,  x150 y130 w283 h48 ,Kontakt til Vognmand:
Gui svigt: Add, Radio, vSvigtVmKontaktradio x160 y147 h23, Kontaktet
Gui svigt: Add, Radio, vSvigtIngenVmKontaktRadio x240 y147 h23, Ingen kontakt
Gui svigt: Add, Edit, disabled vsvigtVmKontaktEdit x340 y147 w50, ca. kl.
Gui svigt: Add, Text, x16 y157 h23 +0x200, &Beskrivelse
Gui svigt: Font
Gui svigt: Font, s9, Segoe UI
Gui svigt: Add, Edit, vbeskrivelse x16 y180 w410 h106
Gui svigt: Add, CheckBox, vgemt_ja x16 y294, Brug &forrige skærmklip
; Gui svigt: Add, CheckBox, vgemt_j x5 y374, Tilføj & skærmklip
Gui svigt: Add, Button, x160 y309 w60 h23 vvis +default, &Vis
Gui svigt: Add, Button, x240 y309 w60 h23 vsend , &Send
; Gui svigt: Add, Button , vvogngruppesvigt x320 y354 w80, Op&ret vogngruppesvigt
Gui svigt: menu, svigtMenu
; Gui svigt: Add, text , x280 y261, Anden &Dato
; Gui svigt: Add, Edit , vny_dato x360 y256 w60,



+^z::
{

   Gui svigt: Show, w448 h350, Svigt
  return


}
Skærmprint:
Flexfinderskærmprint:
nuværendeSkærmprint:
svigtHjælp:
svigtVogngruppe:
SvigtSkærmprintOversigt:
MsgBox, , , Hjælp 