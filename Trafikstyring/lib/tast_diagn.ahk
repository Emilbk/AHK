CoordMode, ToolTip, Screen
q::Send, {Space Down}{LButton Down} ; Test
*F12::SetTimer, *F12 Up,% (F12:=!F12)?100:"Off"
*F12 Up::ToolTip,% KeyCombination(), 200, 300
*F11::
Loop, 0xFF
	IF GetKeyState(Key:=Format("VK{:X}",A_Index))
		SendInput, {%Key% up}
Return
KeyCombination(ExcludeKeys:="")
{ ;All pressed keys and buttons will be listed
	ExcludeKeys .= 
	Loop, 0xFF
	{
		IF !GetKeyState(Key:=Format("VK{:02X}",0x100-A_Index))
			Continue
		If !InStr(ExcludeKeys,Key:="{" GetKeyName(Key) "}")
			KeyCombination .= RegexReplace(Key,"Numpad(\D+)","$1")
	}
	Return, KeyCombination
}