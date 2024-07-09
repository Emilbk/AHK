#Requires AutoHotkey v2.0
FileEncoding "UTF-8"

;; Vognløbsdata

; garantidata-array
; 1 vognløbsnummer
; 2 kørselsaftale
; 3 garantiperiode hverdage
; 4 garantiperiode weekend/helligdag
; 5-11 garantidage
; 12 ferieuger
; 13 jul
; 14 nytår

; gv_garantidata := indhent_garanti_data()
; datotid := 202405141453
; MsgBox FormatTime(202405141453, "dddd")
; ; datotid := "202412311753"
; ; datotid := "202407071753"
; ; OBS fix data 3100_47 i gv_garantidage.tsv
; gv_data := vognløb_indhent_data(gv_garantidata, "3100_47", datotid)
; resultat := vognløb_bearbejd_data(gv_data)
; MsgBox resultat["Garanti"]
vognløb_indhent_data(data, kørselsaftale, datotid)
{
    data_ud := map()
    dato_dag := FormatTime(datotid, "dddd")
    dato_time := FormatTime(datotid, "HH")
    dato_uge := SubStr(formatTime(datotid, "yweek"), 5, 2)
    data_ud["Garanti_hverdag"] := 0
    data_ud["Garanti_weekend"] := 0
    data_ud["Garanti_periode_idag"] := 0
    data_ud["aktiv_jul"] := 0
    data_ud["aktiv_nytår"] := 0
    data_ud["ferie_i_nuværende_uge"] := 0
    data_ud["Aktiv_garanti_dag"] := 0
    data_ud["indenfor_garanti_timer_hverdag"] := 0
    data_ud["indenfor_garanti_timer_weekend"] := 0
    data_ud["indenfor_garanti_timer_idag"] := 1
    data_ud["er_garanti_vogn"] := 0
    data_ud["dato_er_hverdag"] := 0

    for i, e in ["mandag", "tirsdag", "onsdag", "torsdag", "fredag"]
        if (dato_dag == e)
            data_ud["dato_er_hverdag"] := 1


    for index_yderst, element_yderst in data
    {
        if (index_yderst == 1)
            continue
        if (kørselsaftale == data[index_yderst][2])
        {
            data_ud["er_garanti_vogn"] := 1
            data_ud["Garanti_hverdag"] := data[index_yderst][3]
            data_ud["Garanti_weekend"] := data[index_yderst][4]
            if (dato_time >= SubStr(data_ud["Garanti_hverdag"], 1, 2) and dato_time <= SubStr(data_ud["Garanti_hverdag"], 9, 2) and data_ud["dato_er_hverdag"])
            {
                data_ud["indenfor_garanti_timer_idag"] := 1
            }
            if (data[index_yderst][4] != "")
                if (dato_time >= SubStr(data_ud["Garanti_weekend"], 1, 2) and dato_time <= SubStr(data_ud["Garanti_weekend"], 9, 2) and !data_ud["dato_er_hverdag"])
                {
                    data_ud["indenfor_garanti_timer_idag"] := 1
                    data_ud["Garanti_periode_idag"] := data[index_yderst][3]
                }
            for index_ferie_tjek, element_ferie_tjek in data[index_yderst]
            {
                ; tjek for ferie
                if (InStr(data[index_yderst][12], dato_uge))
                    data_ud["ferie_i_nuværende_uge"] := 1
                ; tjek for jul
                if (((SubStr(datotid, 5, 4) == 1225 or SubStr(datotid, 5, 4) == 1226)) and data[index_yderst][13] == "Ja")
                    data_ud["aktiv_jul"] := 1
                ; tjek for nytår
                if (((SubStr(datotid, 5, 4) == 1231 or SubStr(datotid, 5, 4) == 0101)) and data[index_yderst][14] == "Ja")
                    data_ud["aktiv_nytår"] := 1

                ; tjek for weekend eller hverdag
                for index_dag_tjek, element_dag_tjek in data[index_yderst]
                {
                    if (data[1][index_dag_tjek] == dato_dag)
                    {
                        if (element_dag_tjek == "Ja")
                        {
                            data_ud["Aktiv_garanti_dag"] := 1
                            break 3
                        }
                        if (element_dag_tjek == "Nej")
                        {
                            data_ud["Aktiv_garanti_dag"] := 0
                            break 3
                        }

                    }
                }
            }
        }

    }
    return data_ud

}

vognløb_bearbejd_data(vognløbs_data)
{
    vognløbs_data["Garanti"] := 0
    if (!vognløbs_data["er_garanti_vogn"])
    {
        vognløbs_data["variabelt_vognløb"] := 1
        return vognløbs_data
    }
    if (vognløbs_data["aktiv_jul"] or vognløbs_data["aktiv_nytår"])
    {

        vognløbs_data["Garanti"] := 1
        return vognløbs_data
    }
    if (vognløbs_data["ferie_i_nuværende_uge"])
    {
        return vognløbs_data
    }
    if (vognløbs_data["Aktiv_garanti_dag"])
    {

        vognløbs_data["Garanti"] := 1
        return vognløbs_data
    }
    if (!vognløbs_data["Aktiv_garanti_dag"])
    {
        return vognløbs_data
    }

}
; test_data := [["k", [d]]]
indhent_garanti_data()
{
    gv_garantidage_fil := "../lib/gv_garantidage.tsv"

    gv_garantidage_ind := FileRead(gv_garantidage_fil)
    gv_garantidage_ind := StrReplace(gv_garantidage_ind, "`r", "")
    gv_garantidage_ind := StrSplit(gv_garantidage_ind, "`n")
    gv_garantidage := []

    for i, e in gv_garantidage_ind
    {
        gv_garantidage.Push(StrSplit(gv_garantidage_ind[i], "`t"))
    }
    gv_garantidage_ind := unset

    return gv_garantidage
}


; test

; test_garanti(k_array)
; {

; }
; ; p6_svigt_tjek_ugedag(k_aftale, dato)

; {
;         gv_svigt := []
;         FileRead, gv_svigt_ind, db\gv_svigt.txt
;         gv_svigt_ind := StrReplace(gv_svigt_ind, "`r", "")
;         gv_svigt_ind := StrSplit(gv_svigt_ind, "`n")
;         for i,e in gv_svigt_ind
;             {
;                 gv_svigt[i] := StrSplit(gv_svigt_ind[i], "`t")
;             }
; for i, e in gv_svigt
;     {
;         ; MsgBox, , , %gv_svigt%, Timeout]
;         if (gv_svigt[i][1] = k_aftale)
;             {
;             vl := gv_svigt[i][2]
;             break
;             }
;     }


; Tirsdag, 14 maj 2024
; datotid := 202405141453
; Juleaften
; datotid := 202412241100
; Nytårsdag
; datotid := 202501010900
; Søndag, uge 11
; datotid := 202403171050
; Lørdag, uge 28
; datotid := 202407131145
; Tirsdag, uge 28
; datotid := 202407130945
; Mandag, uge 31
; datotid := 202407290945

test_data :=
    [
        ["Lørdag, 11 maj 2024, uge 20 ", 20240511, 1453, "3100_47", 1],
        ["Tirsdag, 14 maj 2024, uge 20 ", 20240514, 1453, "3100_47", 1],
        ["Juleaften 2024 ", 20241224, 1453, "31002_47", 1],
        ["Nytårsdag 2025 ", 20250101, 1453, "3100_47", 1],
        ["Søndag, 17 marts 2024, uge 11 ", 20240317, 1453, "3100_47", 1],
        ["Lørdag, 13 juli 2024, uge 28 ", 20240713, 1453, "3100_47", 1],
        ["Tirsdag, 16 juli 2024, uge 29 ", 20240716, 1453, "3100_47", 1],
        ["Mandag, 29 juli 2024, uge 31 ", 20240729, 1453, "3100_47", 1]
    ]
;
test_funk(test_data)
{
    garantidata := indhent_garanti_data()
    for index, element in test_data
    {
        datotid := test_data[index][2] test_data[index][3]
        k_aftale := test_data[index][4]
        vognløbsdata := vognløb_indhent_data(garantidata, k_aftale, datotid)
        resultat := vognløb_bearbejd_data(vognløbsdata)
        MsgBox (
            "Vognløb: " test_data[index][4]
            "`nDato: " FormatTime(datotid, "dddd") " d. " FormatTime(datotid, "dd/MM") " - uge " SubStr(FormatTime(datotid, "yweek"), 5, 2)
            "`nEr garantivogn: " resultat["er_garanti_vogn"]
            "`n`nHar garanti: " resultat["Garanti"]
            "`nGarnantitid hverdag: " resultat["Garanti_hverdag"]
            "`nGarantitid weekend/helligdag" resultat["Garanti_weekend"]
            "`n`n"
            "Kører i dag garanti: " resultat["Garanti_periode_idag"]
        )
    }
    return
}

test_funk(test_data)