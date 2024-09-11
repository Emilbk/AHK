#Requires AutoHotkey v2.0

testvlliste := ["31200", "31204", "31430"]

; GUI
ListeGUI := Gui(, "VL-Liste")

ListeGUI.SetFont("Bold")
ListeGUI.Add("Text","Section w200" , "Replaneringer")
ListeGUI.SetFont("Norm")
ListboxReplaner := ListeGUI.Add("ListBox", "XP w300 h449" , testvlliste)
ListeGUI.Add("Text", "YS XS+310" , "Wakeup")
ListboxWakeup := ListeGUI.Add("ListBox", "YP+20 w300 h449 XP" , testvlliste)
ListeGUI.Add("Text", "YS XS+620" , "Privatr.")
ListboxPrivatrejser := ListeGUI.Add("ListBox", "YP+20 w300 h449 XP" , testvlliste)
ListeGUI.Add("Text", "YS XS+930" , "Huskeliste")
ListboxHuskeliste := ListeGUI.Add("ListBox", "YP+20 w300 h449  XP" , testvlliste)
ListeGUI.Add("Text", "YS XS+1240" , "Kvitt.")
ListboxKvitteringer := ListeGUI.Add("ListBox", "YP+20 w300 h449 XP" , testvlliste)
ListeGUI.Add("Text", "YS XS+1550" , "Låst.")
ListboxLåst := ListeGUI.Add("ListBox", "YP+20 w300 h449 XP" , testvlliste)
; ListboxReplaner := ListeGUI.Add("ListBox", "x430 wp h449" ,["31200", "31202"])
; ListboxReplaner := ListeGUI.Add("ListBox", "x636 w200 h449" ,["31200", "31202"])

ListeGUI.show("w1872 h574")


; Funktioner