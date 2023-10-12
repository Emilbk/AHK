#NoEnv
#SingleInstance, Force

#Include, %A_linefile%\..\..\lib\DSVParser.ahk
#Include, %A_linefile%\..\..\lib\winHTTPRequestWrapper.ahk
#Include, %A_linefile%\..\..\lib\JSON.ahk
#Include, %A_linefile%\..\..\lib\DSVParser.ahk
; #Include, %A_linefile%\..\..\lib\Biga-AHK\export.ahk

SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%








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
        ; OutputDebug, % række2
        knudepunkt.ind[hver][hver2] := StrReplace(r2, ",", ".")

        ; OutputDebug, % række2
    }

; MsgBox, , , % knudepunkt[2][2]
; OutputDebug, % knudepunkt.samlet[2][2]
; OutputDebug, % knudepunkt[2]

; knudepunkt.navn := ["Knudepunkt 1", "Knudepunkt 2", "Knudepunkt 3", "Knudepunkt 4"]
; knudepunkt.geo := ["10.157957810640028,56.110729831443734", "37.573242,55.801281", "115.663757,38.106467"]
; knudepunkt.navn_geo := {"-FLEX Aktcen., Odg.vej": "9,019578, 56,569738"}

knudepunkt.ValgtLatLong := []
knudepunkt.ValgtLatLong.InsertAt(2, søgt_adresse.geo_lat)
knudepunkt.ValgtLatLong.InsertAt(3, søgt_adresse.geo_long)
knudepunkt.ValgtLatLong.InsertAt(4, søgt_adresse.geo_lat ", " søgt_adresse.geo_long)

; resultat := []
; for hver, række in knudepunkt
;     for hver, række2 in række
;     {
;     resultat := række2 - 10
;     ; resultat := knudepunkt.ValgtLatLong[2]
;     MsgBox, , , % resultat
;     }

knudepunkt.resultat := objFullyClone(knudepunkt.ind)
; knudepunkt.resultat := [3]
; knudepunkt.resultat[1] := []
; knudepunkt.resultat := knudepunkt.ind.Clone()
; knudepunkt.resultat[1].Push([])
for h, r in knudepunkt.ind
    for h2, r2 in r
    {
        if r2 is number
        {
            sum := r2 - knudepunkt.ValgtLatLong[h2]
            knudepunkt.resultat[h].RemoveAt(h2)
            knudepunkt.resultat[h].InsertAt(h2, sum)
        }
    }

; for h, r in knudepunkt.ind
;     for h2, r2 in r
;     {
;         if r2 is number
;         {
;             if (h2 = 3)
;             {
;                 if (knudepunkt.resultat[h][2] < 0 and knudepunkt.resultat[h][3] < 0 )
;                     ; MsgBox, , ,% "begge under nul " knudepunkt.resultat[h][2] " " knudepunkt.resultat[h][3]
;                     knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] - knudepunkt.resultat[h][3])
;                 if (knudepunkt.resultat[h][2] < 0 and knudepunkt.resultat[h][3] > 0 )
;                     ; MsgBox, , , % "h2 under nul " knudepunkt.resultat[h][2] " " knudepunkt.resultat[h][3]
;                     knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] + knudepunkt.resultat[h][3])
;                 if (knudepunkt.resultat[h][2] > 0 and knudepunkt.resultat[h][3] < 0 )
;                     knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] + knudepunkt.resultat[h][3])
;                 ; MsgBox, , , % "h3 under nul " knudepunkt.resultat[h][2] " " knudepunkt.resultat[h][3]
;                 if (knudepunkt.resultat[h][2] > 0 and knudepunkt.resultat[h][3] > 0 )
;                     knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] + knudepunkt.resultat[h][3])
;                 ; MsgBox, , , % "begge over nul " knudepunkt.resultat[h][2] " " knudepunkt.resultat[h][3]
;                 ; knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] - knudepunkt.resultat[h][3])
;                 ; knudepunkt.resultat[h].Push(knudepunkt.resultat[h][2] + knudepunkt.resultat[h][3])

;             }

;         }
;     }



knudepunkt.udvalg := []
antal_udvalg := 20
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
                lat := knudepunkt.ind[h][2] - knudepunkt.ValgtLatLong[3]
                long := knudepunkt.ind[h][3] - knudepunkt.ValgtLatLong[2]
                if lat Between %x% and %y%
                if long Between %x% and %y%
                {                    ; MsgBox, , , % knudepunkt.resultat[h][1] " er tæt"
                    knudepunkt.udvalg.Push(knudepunkt.ind[h])
                    antal++
                    knudepunkt.resultat[h].RemoveAt(3)
                }
            }
 
            MsgBox, , , Text,             
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

    ; MsgBox, , , % antal,

    
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



http := WinHttpRequest(oOptions)

gade := "møllevangen 23"
postnr := "8310"
komm := "Århus"

InputBox, gade, Gadenavn og nr
; InputBox, postnr, Postnr
; InputBox, kommune, Kommune 

endpoint := "https://api.openrouteservice.org/geocode/search?api_key=5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f&text=" gade "&boundary.country=DK"
response := http.GET(endpoint)

søgt_adresse := JSON.Load(response.text)
søgt_adresse.geo_lat := søgt_adresse.bbox[3]
søgt_adresse.geo_long := søgt_adresse.bbox[4]
clipboard := søgt_adresse.geo_lat " " søgt_adresse.geo_long
OutputDebug, % adresse_geo_lat
OutputDebug, % adresse_geo_long


endpoint := "https://api.openrouteservice.org/v2/matrix/foot-walking"

; json_str = "{""locations"":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],""metrics"":[""distance""],""sources"":[0],""units":""km""}"

str := []
str.locations := [[søgt_adresse.geo_lat,søgt_adresse.geo_long],[10.157957810640028,56.110729831443734],[37.573242,55.801281],[115.663757,38.106467]]
str.metrics := ["distance"]
str.sources := [0]
str.units := "km"

json_str := JSON.Dump(str)

; {"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distance"],"sources":[0],"units":"km"}


headers := []
headers["Content-Type"] := "application/json; charset=utf-8"
headers["Accept"] := "application/json, application/geo+json, application/gpx+xml, img/png"
headers["Authorization"] := "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f"

; body := Map("{"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distansources":[0],"units":"km"}")
response_matrix := http.POST(endpoint, json_str, headers)

parsed_matrix := JSON.Load(response_matrix.text)

test := parsed_matrix.distances[1]

knudepunkt := []
knudepunkt.navn := ["Knudepunkt 1", "Knudepunkt 2", "Knudepunkt 3", "Knudepunkt 4"]
knudepunkt.geo := ["10.157957810640028,56.110729831443734", "37.573242,55.801281", "115.663757,38.106467"]
knudepunkt.navn_geo .= {"-FLEX Aktcen., Odg.vej": "9,019578, 56,569738"}
OutputDebug, % knudepunkt.navn_geo[1]


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


^!e::