#NoEnv
#SingleInstance, Force
#Include, %A_linefile%\..\..\lib\JSON.ahk
FileRead, vl_liste_array_json, vl_tekst.txt
if (vl_liste_array_json = "")
    vl_liste_array := []
else if (vl_liste_array_json != "")
    vl_liste_array := json.load(vl_liste_array_json)
; vl_liste_array := json.load(vl_liste_array_json)

;; GUI

Gui repl: Font, s9, Segoe UI
Gui repl: Add, Text, x8 y0 w120 h23 +0x200, Replaneret
Gui repl: Add, ListBox, x8 y24 w170 h349, ListBox
Gui repl: Add, ListBox, x184 y24 w170 h349, 
Gui repl: Add, ListBox, x360 y24 w170 h349, ListBox
Gui repl: Add, ListBox, x536 y24 w170 h349, ListBox
Gui repl: Add, Text, x184 y0 w120 h23 +0x200, WakeUp
Gui repl: Add, Text, x360 y0 w120 h23 +0x200, Privat&rejse
Gui repl: Add, Text, x536 y0 w120 h23 +0x200, &Listet
Gui repl: Add, Button, x40 y360 w80 h23, Ryd
Gui repl: Add, Button, x224 y360 w80 h23, Ryd
Gui repl: Add, Button, x400 y360 w80 h23, Ryd
Gui repl: Add, Button, x584 y360 w80 h23, Ryd
Gui repl: Add, Button, x304 y408 w131 h23, Vis &note
Gui repl: Add, Button, x304 y440 w131 h23, &Opslag
Gui repl: Add, Button, x304 y472 w131 h23, Opslag og &slet
Gui repl: Add, Button, x304 y504 w131 h23, Sle&t
Gui repl: Add, Button, x304 y536 w131 h23, Slet alt

;; GUI-label



;; GUI-funktioner
^/::
    {
        vl_vis_gui()
        return
    }
    vl_vis_gui()
    {
        vl_gui_repl_liste := vl_dan_liste()
        GuiControl, repl: , ListBox1, %vl_gui_repl_liste%
        Gui repl: Show, w718 h574, Window
        Return

    }
    ; vl_liste_array := [vl31234, "vl31235", "vl31236"]
    ; vl_liste_array[1] := [31234, "replaneret", tidspunkt, notat]
    vl_liste_replaneret_vl()
    {
        vl_replaner_vl_til_liste()
    }
^s::
    {
        vl_slet_fra_liste()
        return
    }
    vl_slet_fra_liste()
    {
        global vl_liste_array
        global vl_liste_array_json

        InputBox, valgt_vl
        for i,e in vl_liste_array
            for i2, e2 in vl_liste_array[i]
            {
                if (i2 = 1 and e2 = valgt_vl)
                    vl_liste_array.RemoveAt(i)
                break
            }
        vl_liste_array_json := json.dump(vl_liste_array)
        FileDelete, vl_tekst.txt
        FileAppend, % vl_liste_array_json, vl_tekst.txt

        return
    }

    vl_replaner_hent_vl()
    {
        replaneret_vl := []
        InputBox, vl, VL
        InputBox, note, Note

        FormatTime, vl_replaner_tidspunkt_vis, YYYYMMDDHH24MISS, HH:mm
        FormatTime, vl_replaner_tidspunkt_intern, YYYYMMDDHH24MISS, HHmmss

        replaneret_vl[1] := vl
        replaneret_vl[2] := ", repl. kl "
        replaneret_vl[3] := vl_replaner_tidspunkt_vis
        replaneret_vl[4] := vl_replaner_tidspunkt_intern
        replaneret_vl[5] := note
        replaneret_vl[6] :=
        replaneret_vl[7] := "|"

        return replaneret_vl
    }

    vl_replaner_vl_til_liste()
    {
        global vl_liste_array

        vl_replaner_listet_vl := vl_replaner_hent_vl()
        vl_liste_array.Push(vl_replaner_listet_vl)

        return
    }

    vl_dan_liste()
    {
        global vl_liste_array
        vl_liste_repl_str := "|"

        for i,e in vl_liste_array
            for i2, e2 in vl_liste_array[i]
            {
                if (i2 = 4) or if (i2 = 5 and e2 = "")
                    {}
                else if (i2 = 5 and e2 != 0)
                    {
                        if (vl_liste_array[i][6] = "")
                            vl_liste_array[i].InsertAt(6, " (N)")
                    }
                    ; if (i2 = 5 and e2 != 0)
                    ; if (i2 = 5 and e2 = 0)
                    ; vl_liste_array.InsertAt(6, "")
                else if (i2 = 5 and e2 = 0)
                    {

                    }
                else
                    vl_liste_repl_str := vl_liste_repl_str . e2
            }
        vl_liste_array_json := JSON.Dump(vl_liste_array)
        vl_liste_arra_json_read := json.load(vl_liste_array_json)
        FileDelete, vl_tekst.txt
        FileAppend, % vl_liste_array_json, vl_tekst.txt

        return vl_liste_repl_str
    }
^e::
    {
        vl_liste_replaneret_vl()
        return
    }
+^e::
    {
        vl_dan_liste()
        Return
    }
    sd := asd

