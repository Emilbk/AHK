#NoEnv
#SingleInstance, Force

#Include, %A_linefile%\..\..\lib\DSVParser.ahk
; #Include, %A_linefile%\..\..\lib\Biga-AHK\export.ahk

SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

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
knudepunkt.ValgtLatLong.Push("placeholder")
knudepunkt.ValgtLatLong.Push("56.156553197153684")
knudepunkt.ValgtLatLong.Push("8.891869399484456")
knudepunkt.ValgtLatLong.Push(knudepunkt.ValgtLatLong[2] "," knudepunkt.ValgtLatLong[3])
, 
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
antal := 0
y := 0.05
x := -0.05
StartTime := A_TickCount
tid := []
igen:
    for h, r in knudepunkt.resultat
        for h2, r2 in r
            if (h2 = 3 and antal < 8)
                {          
                lat := knudepunkt.ind[h][2] - knudepunkt.ValgtLatLong[2]
                long := knudepunkt.ind[h][3] - knudepunkt.ValgtLatLong[3]
                if lat Between %x% and %y%
                if long Between %x% and %y%
                {                    ; MsgBox, , , % knudepunkt.resultat[h][1] " er tæt"
                    knudepunkt.udvalg.Push(knudepunkt.ind[h])
                    antal := antal + 1
                    knudepunkt.resultat[h].RemoveAt(3)
                }
            }
            
    if (antal < 8)
    {
        y := y + 0.05
        x := x - 0.05
        ElapsedTime := A_TickCount - StartTime
        tid.Push(ElapsedTime)
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



^!e::