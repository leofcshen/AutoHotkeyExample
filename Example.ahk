#Requires AutoHotkey v2.0
#SingleInstance force ;只允許一個實例

TrayTip A_ScriptName, "啟動 AutoHotKey 腳本", "Iconi"
TraySetIcon(A_WinDir '\system32\shell32.dll', 15) ; Scissors ; Scissors

;幫 SciTE4AutoHotkey 加熱鍵
HotIfWinActive A_ScriptName " - SciTE4AutoHotkey"
Hotkey "^S", ReloadOnSave

ReloadOnSave(ThisHotkey) {
	Sleep 1000
	Reload
}

;執行外部方法並回傳結果(會閃 cmd 視窗)
RunCommand(pType, pCommand) {
	return ComObject("WScript.Shell").Exec(pType " " pCommand).StdOut.ReadAll()
}

;執行外部方法並回傳結果(會閃 cmd 視窗)
RunWaitOne(pCommand){
	return ComObject("WScript.Shell").Exec(A_ComSpec " /C " pCommand).StdOut.ReadAll()
}

;執行外部方法並回傳結果
RunResult(pType, pCommand) {
	;指定執行結果暫存位置
	tempFile := A_ScriptDir "\" DllCall("GetCurrentProcessId") ".txt"
	;執行並輸出結果到暫存位置
	if (pType = "python") {
		RunWait A_ComSpec " /c " pType " " pCommand " >>" tempFile,,"Hide"
	} else {
		RunWait A_ComSpec " /c " pType " -ExecutionPolicy Bypass " pCommand " >>" tempFile,,"Hide"
	}
	;開啟暫存檔
	;Run tempFile
	;取結果值
	result := FileRead(tempFile)
	;刪除暫存檔
	FileDelete(tempFile)
	return result
}

;Win + Q ;複製今天日期到剪貼簿
#Q::{
	A_Clipboard := FormatTime(A_Now, "yyyyMMdd")
	TrayTip "已複製今天日期：" A_Clipboard, "Win + Q", "Iconi"
}

;Win + B ;呼叫外部 function
#B::{
	;呼叫 .ps1 PowerShell
	;result := RunCommand("powershell", A_ScriptDir '\PowerShellFunction.ps1 3 3')
	;result := RunWaitOne('powershell ' A_ScriptDir '\PowerShellFunction.ps1 2 3')
	;result := RunResult("powershell", A_ScriptDir "\PowerShellFunction.ps1 4 4")

	;呼叫 .py Python
	;result := RunCommand("python", A_ScriptDir '\PythonFunction.py 3 3 1')
	;result := RunWaitOne('python ' A_ScriptDir '\PythonFunction.py 2 3 3')
	result := RunResult("python", A_ScriptDir "\PythonFunction.py 2 3 4")

	MsgBox result
}

;Win + Z
#Z::{
	timesTotal := 3
	timesRun := 0

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

	timesRun++
	if timesRun = timesTotal + 1 {
		MsgBox "測試結束"
		return
	}

	;隨機取一個題目
	number := Random(1, array.Length)
	answer := array[number][2]

	;Goto 重新測試
	Retry:

	exam :=	(
		"題目數：" array.Length "`n"
		"第 " timesRun " / " timesTotal "次測試，請輸入啟動下列服務的指令`n"
		array[number][1]
	)

	IB := InputBox(exam, "快捷鍵指令練習")

	if IB.Result = "Cancel" {
		return
	}

	if IB.value = answer {
		MsgBox "輸入 '" IB.value "' 正確。"
		goto NewTest
	}

	if MsgBox("輸入 '" IB.value "' 錯誤.", , "R/C") = "Retry" {
		goto Retry
	}

	MsgBox Format("啟動 {1} 的指令為 {2}", array[number][1], array[number][2])
	goto NewTest
}