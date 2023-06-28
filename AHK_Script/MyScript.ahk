#Requires AutoHotkey v2.0

class Common {

}

class MyClass extends Common {
	;{ 回傳今天日期_yyyyMMdd
	GetToday_yyyyMMdd() => FormatTime(A_Now, "yyyyMMdd")
	;}
	;{ 執行外部方法並回傳結果(會閃 cmd 視窗)
	static RunCommand(pType, pCommand) => ComObject("WScript.Shell").Exec(pType " " pCommand).StdOut.ReadAll()
	;}
	;{ 執行外部方法並回傳結果(會閃 cmd 視窗)
	static RunWaitOne(pCommand) => ComObject("WScript.Shell").Exec(A_ComSpec " /C " pCommand).StdOut.ReadAll()
	;}
	;{ 執行外部方法並回傳結果(不會閃 cmd 視窗，結果先存檔再取值)
	static RunResult(pType, pCommand) {
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
	;}
	;{ 新增重新載入腳本熱鍵在編輯器存檔時觸發
	static AddReloadHotkey(reloadDelay) {
		editorList := [" - SciTE4AutoHotkey", " - Notepad++", " - 記事本"]
		for index, editor in editorList {
			HotIfWinActive A_ScriptName editor
			Hotkey "~^S", ReloadOnSave.Bind(reloadDelay)
		}

		ReloadOnSave(reloadDelay, ThisHotkey) {
			Tooltip reloadDelay " 秒後重新載入腳本"
			Sleep reloadDelay * 1000
			Reload
		}
	}
	;}
}
