#Requires AutoHotkey v2.0


; Aktiverer P6-vindue, hvis ikke aktivt
P6_aktiver()
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
P6_alt_menu(tast1, tast2?)
{
    SendInput "{esc}{alt}"
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


; Aktiverer alt-menu i P6, tager op til to taste-sekvenser
P6_luk_vinduer()
{
    SendInput "{esc}{alt}"
    sleep 20
    Sendinput "{v 2}{Down 2}{Enter}"

    return
}

P6_nav_kørselsaftale()
{
    P6_aktiver()
    P6_alt_menu("t", "k")


    return
}


P6_nav_vognløb()
{
    P6_aktiver()
    P6_alt_menu("t", "l")


    return
}

p6_åben_vognløb(p_vl_obj, dato)
{
    P6_aktiver()
    SendInput(p_vl_obj.vl_data["Vognløbsnummer"])
    SendInput "{tab}"
    SendInput(dato)
    SendInput("{enter}")
    sleep 20
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    if (InStr(A_Clipboard, "eksistere ikke"))
        throw Error("Ikke registreret - TODO")
    ; tjek af korrekt vognløb
    A_Clipboard := ""
    SendInput("+{f10}c")
    ClipWait 1
    indlæst_vognløbsnummer := A_Clipboard
    SendInput("{tab}")
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    SendInput("{tab}")
    ; SendInput("^{F4}")
    indlæst_dato := A_Clipboard
    if (p_vl_obj.vl_data["Vognløbsnummer"] = indlæst_vognløbsnummer and dato = indlæst_dato)
        korrekt := 1
    ; MsgBox "korrekt"
    return
}

p6_åben_vognløb_kørselsaftale(p_vl_obj)
{
    P6_aktiver()
    SendInput("^æ")
    SendInput(p_vl_obj.vl_data["Kørselsaftale"])
    SendInput "{tab}"
    SendInput(p_vl_obj.vl_data["Styresystem"])
    SendInput("{tab}")
    ; A_Clipboard := ""
    ; SendInput("+{f10}c")
    ; ClipWait 1
    ; if (InStr(A_Clipboard, "eksistere ikke"))
    ; throw Error("Ikke registreret - TODO")
    ; tjek af korrekt vognløb
    A_Clipboard := ""
    SendInput("+{f10}c")
    ClipWait 1
    indlæst_kørselsaftale := A_Clipboard
    SendInput("{tab}")
    A_Clipboard := ""
    SendInput("+{f10}c")
    ClipWait 1
    ; SendInput("^{F4}")
    indlæst_styresystem := A_Clipboard
    if (p_vl_obj.vl_data["Kørselsaftale"] = indlæst_kørselsaftale and p_vl_obj.vl_data["Styresystem"] = indlæst_styresystem)
        korrekt := 1
    ; MsgBox "korrekt"
    return
}

p6_åben_vognløb_åbningstider(p_vl_obj)
{
    P6_aktiver()
    SendInput("{enter}")
    SendInput(p_vl_obj.vl_data["Dato"][1] "{tab 2}")
    SendInput(p_vl_obj.vl_data["Dato"][1] "{tab 2}")
    SendInput(p_vl_obj.vl_data["Dato"][1] "{tab 2}")
    SendInput(p_vl_obj.vl_data["Startzone"] "{tab}")
    SendInput(p_vl_obj.vl_data["Slutzone"] "{tab}")
    SendInput(p_vl_obj.vl_data["Hjemzone"] "{tab}")
    SendInput("{enter}")

    return
}
p6_åben_vognløb_resten(p_vl_obj)
{
    P6_aktiver()
    if (p_vl_obj.vl_data["Vognløbsnotering"])
        SendInput("!p{tab 11}+{Up}" p_vl_obj.vl_data["Vognløbsnotering"])
    if (p_vl_obj.vl_data["MobilnrChf"])
        SendInput("!ø{tab 2}" p_vl_obj.vl_data["MobilnrChf"])
    if (p_vl_obj.vl_data["Vognløbskategori"])
        SendInput("!ø{tab 3}" p_vl_obj.vl_data["Vognløbskategori"])
    if (p_vl_obj.vl_data["Planskema"])
        SendInput("!ø{tab 6}" p_vl_obj.vl_data["Planskema"])
    if (p_vl_obj.vl_data["Økonomiskema"])
        SendInput("!ø{tab 8}" p_vl_obj.vl_data["Økonomiskema"])
    if (p_vl_obj.vl_data["Statistikgruppe"])
        SendInput("!ø{tab 9}" p_vl_obj.vl_data["Statistikgruppe"])

    if (p_vl_obj.vl_data["Undtagne transporttyper"])
    {
        SendInput("!ø{tab 10}")
        for trtype in p_vl_obj.vl_data["Undtagne transporttyper"]
            SendInput("{tab}" trtype)
    }
}

p6_åben_kørselsaftale(p_vl_obj)
{
    P6_nav_kørselsaftale()
    sleep 100
    SendInput(p_vl_obj.vl_data["Kørselsaftale"])
    SendInput "{tab}"
    SendInput(p_vl_obj.vl_data["Styresystem"])
    SendInput("{enter}")
    sleep 200
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    if (InStr(A_Clipboard, "ikke registreret"))
        throw Error("Ikke registreret - TODO")
    ; tjek af korrekt kørselsaftale
    A_Clipboard := ""
    SendInput("+{f10}c")
    ClipWait 1
    indlæst_kørselaftale := A_Clipboard
    SendInput("{tab}")
    A_Clipboard := ""
    SendInput("+{f10}c")
    ClipWait 1
    SendInput("{tab}")
    ; SendInput("^{F4}")
    indlæst_styresystem := A_Clipboard
    if (p_vl_obj.vl_data["Kørselsaftale"] = indlæst_kørselaftale and p_vl_obj.vl_data["Styresystem"] = indlæst_styresystem)
        ; MsgBox "korrekt"
        return
}

p6_indlæs_data_kørselsaftale_æ()
{
    SendInput("^æ")

    return
}

p6_indlæs_data_kørselsaftale_planskema(p_vl_obj)
{
    SendInput("!p")
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    tidligere_planskema := A_Clipboard
    SendInput(p_vl_obj.vl_data["Planskema"] "{tab}!p")
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    indlæst_planskema := A_Clipboard
    if (p_vl_obj.vl_data["Planskema"] = indlæst_planskema)
        korrekt := 1
    return
}

p6_indlæs_data_kørselsaftale_økonomiskema(p_vl_obj)
{
    SendInput("!p{tab 4}")
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    tidligere_planskema := A_Clipboard
    SendInput(p_vl_obj.vl_data["Planskema"] "{tab}!p{tab 4}")
    A_Clipboard := ""
    SendInput("^c")
    ClipWait 1
    indlæst_planskema := A_Clipboard
    if (p_vl_obj.vl_data["Planskema"] = indlæst_planskema)
        korrekt := 1
    return
}