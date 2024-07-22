; sdflkjsdlfk
class vognløbObj extends Object
{

    __New(vognløbsnummer := 0) {
        this.vognløbsnummer := vognløbsnummer
        this.kørselsaftale := 0
        this.garantivogn_tjek := 0
        this.gv := 0
        this.vognløbsdato_timestamp := 0

        this.gv_variabel := 0
        this.variabel := 0
        this.vogngruppe := 0
        this.garanti_data := 0
        
        this.garanti_data := vognløbObj.indhent_garanti_data()
    }

    static indhent_garanti_data()
    {
        gv_garantidage_fil := "lib/gv_garantidage.tsv"
        garanti_data_input := FileRead(gv_garantidage_fil)
        garanti_data_input := StrReplace(garanti_data_input, "`r", "")
        garanti_data_input := StrSplit(garanti_data_input, "`n")
        garanti_data_output := []

        for i, e in garanti_data_input
        {
            garanti_data_output.Push(StrSplit(garanti_data_input[i], "`t"))
        }
        garanti_data_input := unset


        return garanti_data_output
    }



    ; input map(vognløbsnummer, vognløbsdato, kørselsaftale, styresystem)
    ; omskriv som del af anden funktion?
    hent_data_vognløb_alt_obj(data)
    {


        this.vognløbsnummer := data["vognløbsnummer"]
        this.vognløbsdato := data["vognløbsdato"]
        this.kørselsaftale_uden_styresystem := data["kørselsaftale"]
        this.styresystem := data["styresystem"]
        this.vognløbsdato_timestamp := SubStr(this.vognløbsdato, -4, 4) . SubStr(this.vognløbsdato, 4, 2) . SubStr(this.vognløbsdato, 1, 2)
        this.kørselsaftale := this.kørselsaftale_uden_styresystem "_" this.styresystem

        return
    }

    ; Tjek om vl. er garantivl. Tager  som optional parameter, ellers obj.vognløbsnummer
    ; => bool
    gv_tjek(p_kørselsaftale := 0)
    {
        if p_kørselsaftale
            this.kørselsaftale := p_kørselsaftale
        if this.kørselsaftale = 0
            throw Error("Kørselsaftale er ikke defineret.")

        for i, e in this.garanti_data
        {
            if (this.garanti_data[i][2] = this.kørselsaftale)
            {
                this.garantivogn_tjek := 1
                this.array_plads := i
                break
            }
        }

    }


    ; udfolder garantidataarray for en given kørselsaftale
    unpack_garantidata(p_kørselsaftale)
    {
        ; OnError send_fejl_meddelelse ; unassigned variable?
        if p_kørselsaftale
            kørselsaftale := p_kørselsaftale
        if !p_kørselsaftale
            kørselsaftale := this.kørselsaftale
        if !kørselsaftale
            throw Error("Der er ikke defineret en kørselsaftale")
        try
        {
            for i, e in this.garanti_data
                if (this.garanti_data[i][2] = kørselsaftale)
                {
                    this.garanti_periode_hv := e[3]
                    this.garanti_periode_we := e[4]
                    this.garanti_mandag := e[5]
                    this.garanti_tirdag := e[6]
                    this.garanti_onsdag := e[7]
                    this.garanti_tordag := e[8]
                    this.garanti_fredag := e[9]
                    this.garanti_lørdag := e[10]
                    this.garanti_søndag := e[11]
                    this.ferieuger := e[12]
                    this.garanti_25_26 := e[13]
                    this.garanti_31_01 := e[14]

                    return

                }
            throw Error("Kørselsaftale er ikke defineret i garantivognsdata", p_kørselsaftale)
        }
        catch as e
        {
            ; send_fejl_meddelelse(e) ; unassigned variable?
            return
        }
    }
    ; TODO mulighed for at tage kun dato, ikke tid, som parameter
    ; Vognløbets aktuelle status hvad angår garantiperiode/variabel
    ; Optional param. dato i YYYYMMDDHH24MI, string. Ellers dags dato
    ; => this.status
    vognløb_status(p_dato := 0, p_kørselsaftale := 0)
    {
        ; tjek kørselsaftale
        if !this.kørselsaftale or !IsSet(p_kørselsaftale)
            throw Error("Kørselsaftale er ikke defineret")
        ; tjek om dato er givet
        if p_dato
            if !IsTime(p_dato)
                throw Error("Den givne dato er ikke i validt format")
            else
                dato := p_dato
        else
            dato := A_Now

        this.gv_tjek(this.kørselsaftale)

        ; tjek om variabelt vognløb
        if !this.garantivogn_tjek
            ; skriv tjek om vogngruppevogn
        {
            this.variabel := 1
            this.status := "Variabelt driftsvognløb"
            return
        }

        ; hvis gv:

        if this.garantivogn_tjek
        {
            this.unpack_garantidata(this.kørselsaftale)
            ; skriv jul/nytårtjek

            ; tjek om tvungen ferie
            ugenr := SubStr(FormatTime(dato, "yweek"), 5, 2)
            if (InStr(this.ferieuger, ugenr))
            {
                this.status := "Garantivognløb m. tvungen ferie uge " ugenr
                this.gv_variabel := 1
                return
            }

            ; tjek om helligdag
            ; hvor er det opgjort?

            ; tjek om aktiv garantidag
            ugedag := FormatTime(dato, "dddd")

            if (this.garanti_%ugedag% = "Nej")
            {
                this.status := "Garantivognløb på tvunget lukket ugedag - " ugedag
                this.gv_variabel := 1
                return
            }

            ; tjek om indenfor garantiperiode
            if p_dato
                tidspunkt := p_dato
            else
                tidspunkt := A_Now

            tidspunkt_time := FormatTime(tidspunkt, "HH")
            tidspunkt_min := FormatTime(tidspunkt, "mm")

            ugedag_tal := FormatTime(dato, "WDay")
            ; minus 1, 1 tælles som som søndag
            if ugedag_tal != 1 ; undtaget søndag, der skal forblive 1
                ugedag_tal -= 1


            garanti_start_hv := SubStr(dato, 1, 8) SubStr(this.garanti_periode_hv, 1, 2) . "00"
            garanti_slut_hv := SubStr(dato, 1, 8) SubStr(this.garanti_periode_hv, 9, 2) . "00"

            garanti_start_we := SubStr(dato, 1, 8) SubStr(this.garanti_periode_we, 1, 2) . "00"
            garanti_slut_we := SubStr(dato, 1, 8) SubStr(this.garanti_periode_we, 9, 2) . "00"


            ; skriv tjek over midnat
            ; if (garanti_start_hv > garanti_slut_hv)
            ; MsgBox "Midnat"

            ; tjek hverdagstider
            ; TODO omskrive status_funktion, skal være en del af svigt-oprettelse
            if (ugedag_tal < 6)
            {
                if (tidspunkt <= garanti_slut_hv and tidspunkt >= garanti_start_hv)
                {
                    this.status := "Aktivt garantivognløb`n`nGaranti " ugedag ":`n" FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm")
                    this.gv := 1
                    this.garanti_periode := garanti_start_hv "-" garanti_slut_hv

                    return ; slutresultat
                }
                else
                {
                    this.status := "Garantivognløb udenfor garanti`n`nGaranti " ugedag ":`n" FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm")
                    this.gv_variabel := 1
                    return ; slutresultat
                }
            }

            ; tjek weekendtider
            if (ugedag_tal > 5)
            {

                if (tidspunkt <= garanti_slut_we or tidspunkt >= garanti_start_we)
                    this.status := "Aktivt garantivognløb`n`nGaranti " ugedag ":`n" FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm")

                if (tidspunkt >= garanti_slut_we or tidspunkt <= garanti_start_we)
                    this.status := "Garantivognløb udenfor garanti`n`nGaranti " ugedag ":`n" FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm")
            }
        }
        throw Error("Intet resultat opnået")
    }
}