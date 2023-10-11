#Include, %A_linefile%\..\lib\winHTTPRequestWrapper.ahk
#Include, %A_linefile%\..\lib\JSON.ahk
#Include, %A_linefile%\..\lib\DSVParser.ahk

;  "Authorization": "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f",
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
InputBox, postnr, Postnr
InputBox, kommune, Kommune 

endpoint := "https://api.openrouteservice.org/geocode/search?api_key=5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f&text=" gade " " postnr " " komm "&boundary.country=DK"
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


MsgBox, , Afstand, % "Nærmeste på " test[1] "km `n" test[2] " km `n" test[3] "`n" test[4]


; MsgBox,response.Text, "GET", 0x40040

; test := []
; test := [{"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696963717687},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}]
; MsgBox, 0x40040, "Get",% response.text,
MsgBox, 0x40040, "Get",% response_matrix.text,
; OutputDebug, % response.text
; geo := SubStr(response.Text, -2, -11)
; MsgBox, , , % geo,

; {"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696964025721},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}