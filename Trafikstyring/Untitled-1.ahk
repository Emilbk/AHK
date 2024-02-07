#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
^a::
{
ImageSearch, OutputVarX, OutputVarY, 0, 0, A_ScreenWidth, A_ScreenHeight, *80 lib\ff_ryd.png
; ImageSearch, OutputVarX, OutputVarY, 0, 0, 1400, 650, *40 lib\vl_laas.png
; ImageSearch, OutputVarX, OutputVarY, 318, 360, 550, 870, *40 lib\pl_laas.png
MsgBox, , , % ErrorLevel, 
}
