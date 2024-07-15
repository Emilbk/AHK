#Requires AutoHotkey v2.0

; Funktioner der interagerer med P6

sleep_konstant := 1

P6_var_sleep(sleep_var)
{
    global sleep_konstant
    sleep sleep_konstant * sleep_var

    return
}

;; P6-navigering

; Aktiverer P6-vindue, hvis ikke aktivt
P6_nav_aktiver()
{

    HotIfWinNotActive "PLANET"
    {
        WinActivate "PLANET"
        WinWaitActive "PLANET"
        sleep 100
        SendInput "{esc}" ; registrerer ikke første tryk, når der skiftes til vindue
        ; sleep 300
        return true
    }
    return false
}

; Aktiverer alt-menu i P6, tager op til to taste-sekvenser
P6_nav_alt_menu(tast1, tast2?)
{
    SendInput "{alt}"
    sleep 20
    Sendinput tast1
    if IsSet(tast2)
    {
        sleep 40
        SendInput tast2
        sleep 40
    }

    return
}

;; P6-navigering, vinduer

;
P6_nav_planbillede()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("tp")

    return
}

P6_nav_rejsesøg()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("rr")
    sleep 200
    SendInput "^t"

    return
}

; tager dato (hvis kun en søgedato), alternativt start og slutdato
P6_nav_rejsesøg_hylde(dato, datoslut?)
{
    P6_nav_rejsesøg

    SendInput "!f"
    sleep 10
    if IsSet(datoslut)
        SendInput dato "{tab 2}" datoslut
    else
        SendInput dato "{tab 2}" dato
    SendInput "!h{space}{enter}"

    return
}

P6_nav_bestilling()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("rb")

    return
}

P6_nav_kørselsaftale()
{
    P6_nav_aktiver()
    P6_nav_planbillede()
    P6_nav_alt_menu("tk")

    sleep 40
    SendInput "!{F5}"

    return

}
P6_nav_kundealarm()
{
    P6_nav_aktiver()
    P6_nav_alt_menu("ta")

    return
}


p6_nav_udråb()
{
    p6_nav_aktiver()
    p6_nav_alt_menu("ta", "!u")

    return
}

p6_nav_tal()
{
    p6_nav_aktiver()
    p6_nav_alt_menu("ta", "!t")

    return
}

;

; går til aktive vognløbs vognløbsbillede, return true når indlæst
P6_nav_vognløbsbillede(planbillede_vognløb)
{
    P6_nav_aktiver()

    sleep 30
    SendInput "^{F12}"
    sleep 150

    ; tjek for popup, igangværende ændring i vognløbsbillede
    A_Clipboard := ""
    SendInput "^c"
    ClipWait 0.3
    if (InStr(A_Clipboard, "opdateringern"))
        SendInput "!y"

    A_Clipboard := ""
    SendInput "+{F10}c"
    ClipWait 1
    vognløbsbillede_vognløb := A_Clipboard
    while (vognløbsbillede_vognløb != planbillede_vognløb)
    {
        if (A_Index == 6)
        {
            return false
        }
        SendInput "!l"
        sleep 10
        SendInput "+{F10}c"
        ClipWait 0.3
        vognløbsbillede_vognløb := A_Clipboard
        sleep 500
    }
    return true
}

;
P6_nav_vognløbsbillede_ændr_1(planbillede_kørselsaftale)
{
    SendInput "^æ"
    A_Clipboard := ""
    sleep 40
    SendInput "^c"
    ClipWait 0.5

    ; Tjek for hop til dato ved lukket vognløb
    if (A_Clipboard != "")
        if (InStr(A_Clipboard, A_Year))
            return "lukket"

    ; tjek for færdigindlæst vognløbsbillede
    sleep 100
    A_Clipboard := ""
    SendInput "+{F10}c"
    ClipWait 0.3
    while (A_Clipboard != planbillede_kørselsaftale)
    {
        if (A_Index == 10)
            return false
        A_Clipboard := ""
        SendInput "!k"
        sleep 20
        SendInput "+{F10}c"
        sleep 500
    }
    vognløbsbillede_kørselsaftale := A_Clipboard

    ; Tjek for vogngruppevognløb
    A_Clipboard := ""
    SendInput "{tab 2}"
    sleep 60
    SendInput "+{F10}c"
    ClipWait 0.3
    if (A_Clipboard = "" or A_Clipboard == vognløbsbillede_kørselsaftale)
        vognløbsbillede_vogngruppe := false
    else
        vognløbsbillede_vogngruppe := true

    SendInput "{enter}"

    return

}

P6_nav_vognløbsbillede_ændr_2()
{

    SendInput "{enter}"

}
P6_nav_vognløbsbillede_ændr_afslut()
{

    SendInput "{enter}"

}
;; P6 indhent data

; Henter tlf fra vl hvis intet parameter, indsætter tlf på vognløb hvis der er
P6_hent_data_vognløbsbillede_telefon(telefonnummer?)
{
    {
        SendInput "{enter}!ø{tab 2}"
        sleep 20
        A_Clipboard := ""
        SendInput "+{F10}c"
        ClipWait 0.3
        while (StrLen(A_Clipboard) != 8)
        {
            if (a_index == 6)
                return false
            SendInput "!ø{tab 2}"
            sleep 20
            SendInput "+{F10}c"
            ClipWait 0.3
        }
        if IsSet(telefonnummer)
        {
            SendInput telefonnummer
            sleep 20
            SendInput "{enter}"
        }
        else
        {
            A_Clipboard := ""
            SendInput "+{F10}c"
            ClipWait 0.5
            SendInput "{enter}"
            return A_Clipboard
        }
    }
    return
}

P6_hent_data_vm_telefon()
{
    P6_nav_kørselsaftale()
    SendInput "^æ"
    sleep 40
    SendInput "!a{tab 4}"
    A_Clipboard := ""
    SendInput "^c"
    ClipWait 0.5
    while (StrLen(A_Clipboard) != 8)
    {
        if (a_index == 4)
            return false

        SendInput "!a{tab 4}"
        A_Clipboard := ""
        SendInput "^c"
        ClipWait 0.5
    }

    SendInput "^a"
    return A_Clipboard
}

p6_hent_data_rejsesøg_telefon(telefon)
{
    P6_nav_rejsesøg()
    SendInput "+{tab 2}" telefon
    sleep 100
    SendInput "{enter}"
}
; Fejlbesked ved fejl i indhentning af data, tager navn på forsøgt data som parameter
P6_hent_data_vis_fejlbesked(indhentet_data)
{
    MsgBox "Der er sket en fejl i indhentning af " indhentet_data ", prøv igen.", "Fejl", 16

    return
}

;
P6_ret_data_vognløbsbillede_ændre_sluttid(vognløb, kørselsaftale, sluttid, dato?)
{
    P6_nav_aktiver()
    P6_nav_vognløbsbillede(vognløb)
    P6_nav_vognløbsbillede_ændr_1(kørselsaftale)

    SendInput "{tab 2}"
    if IsSet(dato)
    {
        SendInput dato
        SendInput "{tab}"
        SendInput sluttid
        SendInput "{tab}"
        SendInput dato
        SendInput "{tab}"
        SendInput sluttid
    }
    else
    {
        SendInput FormatTime(, "dd")
        SendInput "{tab}"
        SendInput sluttid
        SendInput "{tab}"
        SendInput FormatTime(, "dd")
        SendInput "{tab}"
        SendInput sluttid
    }

    return

}

; Return array[4], vognløbsnummer som [1]
P6_hent_data_vognløb_vognløbsnummer()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["vognløbsnummer", "", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}
; Return array[4], vognløbsnummer som [1]
; uden at aktivere planbillede
P6_hent_data_vognløb_vognløbsnummer_i_planbillede()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    for i, e in ["vognløbsnummer", "", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], vognløbsdato som [2]
P6_hent_data_vognløb_vognløbsdato()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["", "vognløbsdato", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], vognløbsnummer som [1], vognløbsdato som [2]
P6_hent_data_vognløb_vognløbsnummer_og_vognløbsdato()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["vognløbsnummer", "vognløbsdato", "", ""]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}
; Return array[4], kørselsaftale som [3], styresystem som [4]
P6_hent_data_vognløb_kørselsaftale_og_styresystem()
{
    hent_data_vognløb_output := ["", "", "", ""]
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for i, e in ["", "", "kørselsaftale", "styresystem"]
    {
        indhentet_data := P6_hent_data_vognløb_funk(e)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[i] := indhentet_data
            break

        }

        hent_data_vognløb_output[i] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Return array[4], [1] vognløbsnummer, [2] vognløbsdato, [3] kørselsaftale (uden styresystem), [4] styresystem
P6_hent_data_vognløb_alt()
{
    hent_data_vognløb_output := map()
    indhentet_data := ""

    P6_nav_aktiver()
    P6_nav_planbillede()

    for index, ønsket_data in ["vognløbsnummer", "vognløbsdato", "kørselsaftale", "styresystem"]
    {
        indhentet_data := P6_hent_data_vognløb_funk(ønsket_data)
        if (indhentet_data == "fejl")
        {
            hent_data_vognløb_output[ønsket_data] := indhentet_data
            break

        }

        hent_data_vognløb_output[ønsket_data] := indhentet_data
    }

    return hent_data_vognløb_output
}


; Tager valgt 1 datatype som input, "vognløb", "vognløbsdato", "kørselsaftale", "styresystem"
; Return "fejl" hvis fejl i indhentning
; => str
P6_hent_data_vognløb_funk(valgt_data)
{
    ; [1] planetgenvej, [2] kopieringsgenvej
    hent_data_input :=
        Map(
            "vognløbsnummer", ["!l", "+{F10}c"],
            "vognløbsdato", ["!l{tab}", "^c"],
            "kørselsaftale", ["!k", "+{F10}c"],
            "styresystem", ["!k{tab}", "+{F10}c"]
        )

    hent_data_output := ""


    if (valgt_data != "")
    {
        SendInput hent_data_input[valgt_data][1]
        A_Clipboard := ""
        SendInput hent_data_input[valgt_data][2]
        ClipWait 0.5
        hent_data_output := A_Clipboard
        while (hent_data_output == "")
        {
            if (A_Index == 3 and valgt_data == "kørselsaftale")
            {
                hent_data_output := "vogngruppe"

                return hent_data_output
            }
            if (A_Index == 10 and valgt_data != "kørselsaftale")
            {
                P6_hent_data_vis_fejlbesked(valgt_data)
                hent_data_output := "fejl"

                return hent_data_output
            }
            P6_nav_aktiver()
            P6_nav_planbillede()
            SendInput hent_data_input[valgt_data][1]
            A_Clipboard := ""
            SendInput hent_data_input[valgt_data][2]
            ClipWait 0.2
            hent_data_output := A_Clipboard
        }
    }
    return hent_data_output
}

