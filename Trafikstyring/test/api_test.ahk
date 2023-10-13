#NoEnv
#SingleInstance, Force

#Include, %A_linefile%\..\..\lib\DSVParser.ahk
#Include, %A_linefile%\..\..\lib\winHTTPRequestWrapper.ahk
#Include, %A_linefile%\..\..\lib\JSON.ahk
#Include, %A_linefile%\..\..\lib\DSVParser.ahk
; #Include, %A_linefile%\..\..\lib\array.ahk
; #Include, %A_linefile%\..\..\lib\Biga-AHK\export.ahk

SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

http := WinHttpRequest(oOptions)

json_str =
(
{
   "metrics": "distance",
   "sources": "0",
   "units": "km",
    "locations": [
       "9.70093,48.477473",
       "9.207916,49.153868",
       "37.573242,55.801281",
       "115.663757,38.106467"
   ]
}
)

parsed := JSON.Load(json_str)

json_str := JSON.Dump(parsed)

parsed := JSON.Load(json_str)

InputBox, gade, Gadenavn og nr
; InputBox, postnr, Postnr
; InputBox, kommune, Kommune

geo_lookup := "https://api.openrouteservice.org/geocode/search?api_key=5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f&text=" gade "&boundary.country=DK"
geo_lookup_response := http.GET(geo_lookup )
geo_resultat := JSON.Load( geo_lookup_response.text)

;  søgt_adresse.geo_lat := søgt_adresse.bbox[3]
;  søgt_adresse.geo_long := søgt_adresse.bbox[4]
;  clipboard := søgt_adresse.geo_lat " " søgt_adresse.geo_long
;  OutputDebug, % adresse_geo_lat
;  OutputDebug, % adresse_geo_long

endpoint := "https://api.openrouteservice.org/v2/matrix/foot-walking"

;; Knudepunkt
FileRead csvStr, %A_linefile%\..\..\db\knudepunkt_geo2.csv
knudepunkt := []
knudepunkt.ind := CSVParser.ToArray(csvStr)

ObjFullyClone(obj)
{
    nobj := obj.Clone()
    for k,v in nobj
        if IsObject(v)
            nobj[k] := A_ThisFunc.(v)
    return nobj
}

for hver, r in knudepunkt.ind
    for hver2, r2 in r
    {
        knudepunkt.ind[hver][hver2] := StrReplace(r2, ",", ".")
    }

knudepunkt.ValgtLat := []
knudepunkt.ValgtLong := []
knudepunkt.ValgtLatLong := []
knudepunkt.ValgtLat.InsertAt(1, geo_resultat.bbox[2])
knudepunkt.ValgtLong.InsertAt(1, geo_resultat.bbox[1])
knudepunkt.ValgtLatLong.InsertAt(1, geo_resultat.bbox[2] ", " geo_resultat.bbox[1])

knudepunkt.resultat := objFullyClone(knudepunkt.ind)

for h, r in knudepunkt.ind
    for h2, r2 in r
    {
        if (h2 = 2)
        {
            knudepunkt.resultat[h][h2] := r2 - knudepunkt.ValgtLat[1]

        }
        else if (h2 = 3)
        {
            knudepunkt.resultat[h][3] := r2 - knudepunkt.ValgtLong[1]
        }
    }

knudepunkt.udvalg := []
antal_udvalg := 40
antal := 0
y := 0.05
x := -0.05
StartTime := A_TickCount
tid := []
tid.tid := []
tid.omgang := [0]
igen:
    for h, r in knudepunkt.resultat
        for h2, r2 in r
            if (h2 = 3 and antal < antal_udvalg)
            {
                if (knudepunkt.resultat[h][2] >= x and knudepunkt.resultat[h][2] <= y)
                    if (knudepunkt.resultat[h][3] >= x and knudepunkt.resultat[h][3] <= y)
                    {
                        knudepunkt.udvalg.Push(knudepunkt.ind[h])
                        antal++
                        knudepunkt.resultat[h].RemoveAt(3)
                    }
            }

    if (antal < antal_udvalg)
    {
        y := y + 0.05
        x := x - 0.05
        tid.omgang[1]++
        ElapsedTime := A_TickCount - StartTime
        tid.tid.Push(ElapsedTime)
        ; MsgBox,  %ElapsedTime% milliseconds have elapsed.
        Goto, igen
    }

    MsgBox, , , % antal,

    ; knudepunkt.udvalg := []
    ; antal := 0
    ; y := 0.01
    ; x := -0.01
    ; StartTime := A_TickCount
    ; tid := []
    ; igen:
    ;     for h, r in knudepunkt.resultat
    ;         for h2, r2 in r
    ;             if (h2 = 4 and antal < 15)
    ;                 if r2 Between %x% and %y%
    ;                 {
    ;                     ; MsgBox, , , % knudepunkt.resultat[h][1] " er tæt"
    ;                     knudepunkt.udvalg.Push(knudepunkt.ind[h])
    ;                     antal := antal + 1
    ;                     knudepunkt.resultat[h].RemoveAt(4)
    ;                 }
    ;     if (antal < 15)
    ;     {
    ;         y := y + 0.5
    ;         x := x - 0.5
    ;         ElapsedTime := A_TickCount - StartTime
    ;         tid.Push(ElapsedTime)
    ;         ; MsgBox,  %ElapsedTime% milliseconds have elapsed.
    ;         Goto, igen
    ;     }

    ;     MsgBox, , , % antal,
    ; str := []
    ; for h, r in knudepunkt.resultat
    ;     for h2, r2 in r
    ;         if (h2 = 4)
    ;         str .= r2 . ", "
    ; str := RTrim(str, ", ")
    ; sort str, N D,

    ; resultat := knudepunkt.ValgtLatLong[
    ; MsgBox, , , % resultat

    ; API
    ;  "Authorization": "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f",

    ; for h, r in knudepunkt.udvalg
    ;     for h2, r2 in r
    ;         MsgBox, , , r2,

    for h, r in knudepunkt.udvalg
        if (h = 98)
        {
            knudepunkt.udvalg[h].Push(knudepunkt.udvalg[h][3] "," knudepunkt.udvalg[h][2])
        }

    ; json_str = "{""locations"":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],""metrics"":[""distance""],""sources"":[0],""units":""km""}"

    str_a := knudepunkt.ValgtLong[1] "," knudepunkt.ValgtLat[1]
    str := []
    locations_a := []
    str.locations := [[]]
    str.locations[1].InsertAt(1, knudepunkt.ValgtLong[1], knudepunkt.ValgtLat[1])
    str.metrics := ["distance"]
    str.destinations := [0]
    str.units := "km"

    for h, r in knudepunkt.udvalg
        for h2, r2 in r
        {
            if (h2 = 2)
            {
                str.locations.push([])
                str.locations[h + 1].Push(knudepunkt.udvalg[h][3])
                str.locations[h + 1].Push(knudepunkt.udvalg[h][2])
            }
        }
    str.locations.RemoveAt(21)
    ; For h, r In knudepunkt.udvalg
    ;     {
    ;         locations_a .= "[" . knudepunkt.udvalg[h][4] . "],"
    ;     }
    ; locations_a := RTrim(locations_a, ",") ; remove the last pipe (|)
    ; locations_a := "[" str_a "]," . locations_a ; remove the last pipe (|)
    ; locations_a := "[" locations_a "]"
    ; str.locations.Push(locations_a)
    ; for h, r in knudepunkt.

    ; str := []
    ; str.locations := [[knudepunkt.ValgtLong[1],knudepunkt.ValgtLat[1]],[37.573242,55.801281],[115.663757,38.106467]]
    ; str.locations_rigtig := [[10.13562,56.155364],[10.203069,56.151446],[10.163294,56.128671],[10.101193,56.182173],[10.174426,56.191931],[10.072098,56.06712],[10.18317,56.004059],[10.262607,56.205151],[10.038531,56.082223],[10.150062,55.971952],[10.15516,55.983864],[10.15378,55.977893],[10.154926,55.97064],[10.133823,55.981016],[10.034847,56.031693],[10.038684,56.047296],[9.991517,56.034826],[9.964364,56.184071],[10.060216,56.263101],[9.962442,56.0499]]
    ; str.metrics := ["distance"]
    ; str.sources := [0]
    ; str.units := "km"

    json_str := JSON.Dump(str)

    MsgBox, , , % json_str,

    ; {"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distance"],"sources":[0],"units":"km"}

    headers := []
    headers["Content-Type"] := "application/json; charset=utf-8"
    headers["Accept"] := "application/json, application/geo+json, application/gpx+xml, img/png"
    headers["Authorization"] := "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f"

    ; body := Map("{"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distansources":[0],"units":"km"}")
    response_matrix := http.POST(endpoint, json_str, headers)

    parsed_matrix := JSON.Load(response_matrix.text)

    test := parsed_matrix.distances[1]

    ; knudepunkt := []
    ; knudepunkt.navn := ["Knudepunkt 1", "Knudepunkt 2", "Knudepunkt 3", "Knudepunkt 4"]
    ; knudepunkt.geo := ["10.157957810640028,56.110729831443734", "37.573242,55.801281", "115.663757,38.106467"]
    ; knudepunkt.navn_geo .= {"-FLEX Aktcen., Odg.vej": "9,019578, 56,569738"}
    ; OutputDebug, % knudepunkt.navn_geo[1]

    ; MsgBox, , Afstand, % "Nærmeste på " test[1] "km `n" test[2] " km `n" test[3] "`n" test[4]

    ; MsgBox,response.Text, "GET", 0x40040

    ; test := []
    ; test := [{"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696963717687},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}]
    ; MsgBox, 0x40040, "Get",% response.text,
    MsgBox, 0x40040, "Get",% response_matrix.text,
    ; OutputDebug, % response.text
    ; geo := SubStr(response.Text, -2, -11)
    ; MsgBox, , , % geo,

    ; {"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696964025721},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}
    knudepunkt.valgt := []
    knudepunkt.valgt.Push(gade)
    ; knudepunkt.valgt.Push(parsed_matrix.sources[2].snapped_distance)
    for h, r in parsed_matrix.sources
        for h2, r2 in r
            if (h >= 2)
                if (h2 = "snapped_distance")
                    knudepunkt.udvalg[h -1].Push(parsed_matrix.sources[h].snapped_distance)
    if (h = parsed_matrix.sources.MaxIndex())
        knudepunkt.udvalg[h].Push(parsed_matrix.sources[h].snapped_distance)

    sorteringsarray := []
    for h, r in knudepunkt.udvalg
        for h2, r2 in r
            if (h2 = 4)
                sorteringsarray.Push(r2)
knudepunkt.sorteret := []
knudepunkt.sorteret := quicksort(sorteringsarray)



quicksort(arr)
{
    if (arr.MaxIndex() <= 1)
        return arr
    mindre := [], samme := [], mere := []
    Pivot := arr[1]
    for k, v in arr
        {
            if (v < Pivot)
                mindre.push(v)
            else if (v > Pivot)
                mere.push(v)
            Else
                samme.push(v)
        }
    mindre := quicksort(mindre)
    ud := quicksort(mere)
    if (samme.MaxIndex())
        ud.InsertAt(1, samme*)
    if (mindre.MaxIndex())
        ud.InsertAt(1, mindre*)
    return ud
}

for h, r in knudepunkt.sorteret


    MsgBox, , , slut, ½