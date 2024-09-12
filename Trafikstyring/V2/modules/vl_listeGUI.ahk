#Requires AutoHotkey v2.0
textWidth := "300"
listboxWidth := "300"
testvlliste := ["31200", "31204", "31430"]

; GUI
ListeGUI := Gui(, "VL-Liste")

ListeGUI.SetFont("Bold")
ListeGUI.Add("Text","Section w" textWidth "" , "Replaneringer")
ListeGUI.SetFont("Norm")
ListboxReplaner := ListeGUI.Add("ListBox", "XP w" listboxWidth " h449" , testvlliste)
KnapReplanerSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth / 2 , "Ryd")
ListeGUI.SetFont("Bold")
ListeGUI.Add("Text", "YS XS+" 1 * listboxWidth + 10 "" , "Wakeup")
ListeGUI.SetFont("Norm")
ListboxWakeup := ListeGUI.Add("ListBox", "YP+20 w" listboxWidth " h449 XP" , testvlliste)
KnapWakeupSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth / 2 , "Ryd")
ListeGUI.SetFont("Bold")
ListeGUI.Add("Text", "YS XS+" 2 * listboxWidth + 20 "" , "Privatr.")
ListeGUI.SetFont("Norm")
ListboxPrivatrejser := ListeGUI.Add("ListBox", "YP+20 w" listboxWidth " h449 XP" , testvlliste)
KnapPrivatrejserSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth / 2 , "Ryd")
ListeGUI.SetFont("Bold")
ListeGUI.Add("Text", "YS XS+" 3 * listboxWidth + 30 "" , "Huskeliste")
ListeGUI.SetFont("Norm")
ListboxHuskeliste := ListeGUI.Add("ListBox", "YP+20 w" listboxWidth " h449  XP" , testvlliste)
KnapHuskelisteSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth / 2 , "Ryd")
ListeGUI.SetFont("Bold")
ListeGUI.Add("Text", "YS XS+" 4 * listboxWidth + 40 "" , "Kvitt.")
ListeGUI.SetFont("Norm")
ListboxKvitteringer := ListeGUI.Add("ListBox", "YP+20 w" listboxWidth " h449 XP" , testvlliste)
KnapKvitteringerSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth / 2 , "Ryd")
ListeGUI.SetFont("Bold")
ListeGUI.Add("Text", "YS XS+" 5 * listboxWidth + 50 "" , "Låst.")
ListeGUI.SetFont("Norm")
ListboxLåst := ListeGUI.Add("ListBox", "YP+20 w" listboxWidth " h449 XP" , testvlliste)
KnapLåstSlet := ListeGUI.Add("Button", "Y+10 x+-" listboxWidth * 2/3, "Ryd")
; ListboxReplaner := ListeGUI.Add("ListBox", "x430 wp h449" ,["31200", "31202"])
; ListboxReplaner := ListeGUI.Add("ListBox", "x636 w200 h449" ,["31200", "31202"])

ListeGUI.show("w1872 h574")


; Funktioner