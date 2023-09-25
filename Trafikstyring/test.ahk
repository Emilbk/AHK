
+e::
{
    WinGetPos, X, Y, W, H, FlexDanmark, , , 
    MsgBox, , , % x " pixel " y " pixel " w " pixel " h " pixel ", 
    ImageSearch, Ix, Iy, 0, 0, 1920, 1080, *100 "C:\Users\ebk\AHK\MT-AHK\Trafikstyring\lib\ff.png"
    MsgBox, % Ix Iy
    click, %Ix%, %Iy%
    return
}