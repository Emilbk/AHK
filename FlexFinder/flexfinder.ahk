#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, force

vl := "bla"
FileRead, vl, liftvogn.txt
MsgBox, , , % vl[1]

!^+s::
; IfWinExist, FlexDanmark FlexFinder ;insert the window name
; WinActivate


    vl := ["3220_0032" "3221_0032" "3222_0047" "3223_0047" "3224_0012" "3225_0026" "3126_0026" "3227_0012" "3228_0032" "3229_0032" "3230_0047" "3231_0047" "3232_0047" "3233_0012" "3134_0032" "3235_0012" "3236_0047" "3237_0047" "3238_0047" "3239_0047" "3240_0047" "3241_0012" "3142_0032" "3543_0032" "3244_0026" "3245_0012" "3246_0009" "3247_0047" "3248_0047" "3249_0012" "3250_0012" "3251_0012" "3252_0012" "3253_0012" "3154_0032" "3255_0012" "3256_0012" "3257_0012" "3158_0026" "3259_0047" "3160_0047" "3261_0009" "3162_0047" "3263_0009" "3164_0047" "3265_0009" "3266_0032" "3267_0032" "3268_0012" "3269_0012" "3170_0026" "3271_0047" "3272_0012" "3273_0012" "3274_0032" "3275_0012" "3277_0009" "3178_0009" "3279_0047" "3288_0009" "3289_0047" "3290_0047" "3291_0009" "3292_0026" "3293_0026" "3294_0009" "3295_0009" "3296_0009" "3297_0009" "3298_0047" "3299_0026" "3300_0009" "3301_0009" "3302_0026" "3303_0012" "3304_0009" "3305_0026" "3308_0009" "3309_0009" "3310_0009" "3311_0009" "3312_0047" "3313_0009" "3319_0009" "3320_0012" "3321_0009" "3322_0047" "3323_0032" "3324_0009" "3325_0009" "3327_0009" "3330_0012" "3331_0047" "3336_0035" "3337_0035" "3338_0009" "3339_0035" "3340_0009" "3341_0032" "3143_0032" "3144_0047" "3345_0047" "3346_0026" "3147_0047" "3350_0047" "3351_0032" "3352_0009" "3353_0032" "3354_0026" "3355_0009" "3357_0047" "3358_0009" "3359_0009" "3361_0009" "3365_0012" "3366_0032" "3367_0012" "3368_0009" "3369_0009" "3375_0009" "3376_0009" "3377_0026" "3378_0026" "3379_0047" "3380_0047" "3381_0047" "3382_0047" "3383_0050" "3384_0050" "3385_0009" "3386_0012"]
     for index, element in vl
        {
        val = %element%
        Clipboard = %val%
        ClipWait, 2, 0
        MsgBox, , Vognløb, % val, 1
        SendInput, ^f
        SendInput, {del}
        sleep 200
        SendInput, ^v
        sleep 200
        SendInput, {tab}{tab}{Space}
        sleep 200
        PixelSearch, Px, Py, 90, 190, 1062, 621, 0x3296FF, 0, Fast ; oxo0FFFF is the pixel color fould from using the first script, insert yours there
        sleep 200
        click %Px%, %Py%
        sleep 200
        SendInput, ^f
        sleep 200
        SendInput, {del}
        SendInput, {esc}
        sleep 200
        }

    Return
    
    



    

+z::  ; Control+Z hotkey.
MouseGetPos, MouseX, MouseY
PixelGetColor, color, %MouseX%, %MouseY%
MsgBox The color at the current cursor position is %color%.
return

z::
IfWinExist, FlexDanmark FlexFinder ;insert the window name
WinActivate
PixelSearch, Px, Py, 90, 190, 1062, 621, 0x5E6FF2, 0, Fast ; oxo0FFFF is the pixel color fould from using the first script, insert yours there
if ErrorLevel
MsgBox, That color was not found in the specified region.
else
   click %Px%, %Py%

+Escape::
ExitApp
Return