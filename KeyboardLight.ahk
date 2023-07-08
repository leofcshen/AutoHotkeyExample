#Requires AutoHotkey v2.0

;CapsLock、NumLock、ScrollLock 燈號提示
myGui := Gui()
myGui.Opt("+AlwaysOnTop +ToolWindow -SysMenu -Caption +LastFound")
myGui.SetFont("S12 bold", "Verdana")
myGui.BackColor := 0xaf001d

color := {
	CapsLock: "red",
	ScrollLock: "blue",
	NumLock: "green"
}

for key, txtcolor in color.OwnProps() {
	;myGui.AddText("vTxt" . key . " c" . txtcolor, Format("{:s} {} ", key, GetKeyState(key, "T") ? "ON" : "OFF"))
	myGui.AddText("vTxt" . key . " c" . txtcolor, Format("{:s} {} ", key, "ON"))
}

WinSetTransColor 0xaf001d
myGui.Show("x" . (A_ScreenWidth / 1.2) . "y" . (A_ScreenHeight / 1.2) . " NoActivate")
CheckKeyStates() ; get intitial states on scriptstart
;return

~*NumLock::
~*CapsLock::
~*ScrollLock::CheckKeyStates()

CheckKeyStates() {
	for a, key in ["CapsLock", "ScrollLock", "NumLock"] {
		;GuiControl, % GetKeyState(key, "T") ? "Show" : "Hide", % "Txt" . key
		;切換顯示
		myGui[name:= "Txt" . key].Visible := GetKeyState(key, "T")
	}
}