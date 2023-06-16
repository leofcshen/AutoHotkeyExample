#Requires AutoHotkey v2.0
#SingleInstance force

TrayTip A_ScriptName, "啟動 AutoHotKey 腳本", "Iconi"

;幫 SciTE4AutoHotkey 加熱鍵
HotIfWinActive A_ScriptName " - SciTE4AutoHotkey"
Hotkey "~^S", ReloadOnSave

ReloadOnSave(ThisHotkey) {
	MsgBox "reload"
	Sleep 2000
	Reload
}

;Win + C
#C::{
	timesTotal := 10
	timesRun := 1

	array := [
		["C:\Program Files", "%ProgramFiles%"],
		["C:\Program Files (x86)", "%ProgramFiles(x86)%"],
		["C:\Users\User\AppData", "appdata"],
		["C:\Users\User\AppData\Roaming", "%appdata%"],
		["cmd", "cmd"],
		["PowerShell", "powershell"],
		["PowerShell 7", "pwsh"],
		["事件檢視器", "eventvwr"],
		["剪取工具", "snippingtool"],
		["啟動時執行資料夾", "shell:startup"],
		["小畫家", "snippingtool"],
		["小畫家 3D", "ms-paint:"],
		["小算盤", "calc"],
		["工作排程器", "taskschd.msc"],
		["控制台", "control"],
		["服務", "services.msc"],
		["本機群組原則編輯器", "gpedit.msc"],
		["登錄編輯程式", "regedit"],
		["程式和功能", "appwiz.cpl"],
		["系統資訊", "msinfo32"],
		["螢幕小鍵盤", "osk"],
		["記事本", "notepad"],
		["遠端桌面", "mstsc"],
		["關於 Windows", "winver"],
	]

	;Goto 新測試
	NewTest:

	if timesRun = timesTotal + 1 {
		return
	}

	;隨機取一個題目
	number := Random(1, array.Length)
	answer := array[number][2]

	;Goto 重新測試
	ReTest:

	IB := InputBox("第 " timesRun " / " timesTotal "次測試，請輸入啟動 " array[number][1] " 的指令", "指令測試")

	if IB.Result = "OK" {
		if IB.value = answer{
			MsgBox "輸入 '" IB.value "' 正確."
			timesRun++
			Goto NewTest
		} else {
			if MsgBox("輸入 '" IB.value "' 錯誤.", , "R/C") ="Retry"
				Goto ReTest
			else
				MsgBox "啟動 " array[number][1] " 的指令為 " array[number][2]
		}
	}
}