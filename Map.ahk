#Requires AutoHotkey v2.0
#SingleInstance Force ;跳過對話框自動替換舊實例

;引用腳本
#Include "%A_ScriptDir%\Example\AutoHotKey_Library\ToolTipOptions.ahk"
#Include "%A_ScriptDir%\Example\AutoHotkey_Library\Snipper.ahk"
#Include "%A_ScriptDir%\Example\AutoHotkey_Library\MyLibrary.ahk"

TrayTip A_ScriptName, "啟動 AutoHotKey 腳本", "Iconi"
TraySetIcon(A_WinDir '\system32\shell32.dll', 15) ;設定圖示

reloadDelaySecond := 1
MyClass.AddReloadHotkey(reloadDelaySecond)
;===============================================================

mappy := Map(
	"Alex", 20,
	"Tom", 25,
	"Mary", 17,
)

MsgBox mappy["Alex"] ;20