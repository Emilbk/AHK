#NoEnv
#SingleInstance, Force
#Persistent
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


array := []
array.push("a")
array.push("b")

array[3] := "c" 

test := array[3] . array[2]
knudepunkt := []

knudepunkt[1] := ["10.15776"]
knudepunkt[2] := ["56.110397"]

test2 := knudepunkt[1][1] . knudepunkt[2][1]



^e::