#NoEnv
#SingleInstance, Force

#Include, %A_linefile%\..\..\lib\DSVParser.ahk


SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

FileRead csvStr, %A_linefile%\..\..\db\knudepunkt_geo2.csv
knudepunkt := []
knudepunkt.samlet := CSVParser.ToArray(csvStr)

; OutputDebug, % knudepunkt.samlet[2][2]
; OutputDebug, % knudepunkt[2]

knudepunkt.navn := ["Knudepunkt 1", "Knudepunkt 2", "Knudepunkt 3", "Knudepunkt 4"]
knudepunkt.geo := ["10.157957810640028,56.110729831443734", "37.573242,55.801281", "115.663757,38.106467"]
knudepunkt.navn_geo := {"-FLEX Aktcen., Odg.vej": "9,019578, 56,569738"}

test := knudepunkt.navn_geo["-FLEX Aktcen., Odg.vej"]
; OutputDebug, % test
; MsgBox, % knudepunkt.navn_geo[1]

ny :=
; OutputDebug, % knudepunkt.samlet[3][1]

knudepunkt.samlet[3][1] := "Knudepunkt 5"
; OutputDebug, % knudepunkt.samlet[3][1]

knudepunkt.søgt_lat_long[1] := ["10.15776"]
knudepunkt.søgt_lat_long[2] := ["56.110397"]



knudepunkt.søgt_lat_long[3] := knudepunkt.søgt_lat_long[1]
; på enkelt-plads i array
for hver, række in knudepunkt.samlet
    ; for hver, punkt in række
; MsgBox, , , % række[2], 

; på alle pladser i array
; for hver, række in knudepunkt.samlet
    ; for hver, punkt in række
        ; MsgBox, , , % punkt, 
^e::