; sdflkjsdlfk
class vognløbObj extends Object
{

    __New(vognløbsnummer := 0) {
        this.vognløbsnummer := vognløbsnummer
        this.kørselsaftale := 0
    }

    ; sdfsdf


    ;
    hent_data_vognløb_alt_obj()
    {

        ; P6_nav_aktiver()
        ; P6_nav_planbillede()

        for index, data in ["vognløbsnummer", "vognløbsdato", "kørselsaftale", "styresystem"]
        {
            indhentet_data := ["31320", A_Now, "3100", "47"]
            ; indhentet_data := P6_hent_data_vognløb_funk(data)

        }

        this.vognløbsnummer := indhentet_data[1]
        this.vognløbsdato := indhentet_data[2]
        this.kørselsaftale_uden_styresystem := indhentet_data[3]
        this.styresystem := indhentet_data[4]
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

        for i, e in global_garanti_data
        {
            if (global_garanti_data[i][2] = this.kørselsaftale)
            {
                this.gv := 1
                this.array_plads := i
                break
            }
        }

    }


    ; udfolder garantidataarray for en given kørselsaftale
    unpack_garantidata(p_kørselsaftale)
    {
        for i, e in global_garanti_data
            if (global_garanti_data[i][2] = p_kørselsaftale)
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

        this.unpack_garantidata(this.kørselsaftale)

        ; tjek om variabelt vognløb
        this.gv_tjek(this.kørselsaftale)
        if not this.gv
            ; skriv tjek om vogngruppevogn
        {
            this.status := "Variabelt driftsvognløb"
            return
        }
        ; hvis gv:

        ; skriv jul/nytårtjek

        ; tjek om tvungen ferie
        ugenr := SubStr(FormatTime(dato, "yweek"), 5, 2)
        if (InStr(this.ferieuger, ugenr))
        {
            this.status := "Garantivognløb m. tvungen ferie uge " ugenr
            return
        }

        ; tjek om helligdag
        ; hvor er det opgjort?

        ; tjek om aktiv garantidag
        ugedag := FormatTime(dato, "dddd")

        if (this.garanti_%ugedag% = "Nej")
        {
            this.status := "Garantivognløb på tvunget lukket ugedag - " ugedag
        
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
        if (ugedag_tal < 6)
        {
            if (tidspunkt <= garanti_slut_hv and tidspunkt >= garanti_start_hv)
            {
                this.status := "Aktivt garantivognløb - Garanti: " FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm") " i dag " ugedag
                this.gv_aktiv := 1
                this.garanti_periode := garanti_start_hv "-" garanti_slut_hv

                return ; slutresultat
            }
            else
            {
                this.status := "garantivognløb uden for garanti - Garanti: " FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm") " i dag " ugedag

                return ; slutresultat
            }
        }

        ; tjek weekendtider
        if (ugedag_tal > 5)
        {

            if (tidspunkt <= garanti_slut_we or tidspunkt >= garanti_start_we)
                this.status := "Aktivt garantivognløb - Garanti: " FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm") " i dag " ugedag

            if (tidspunkt >= garanti_slut_we or tidspunkt <= garanti_start_we)
                this.status := "garantivognløb uden for garanti - Garanti: " FormatTime(garanti_start_hv, "HH:mm") "-" FormatTime(garanti_slut_hv, "HH:mm") " i dag " ugedag
        }
        throw Error("Intet resultat opnået")
    }
}