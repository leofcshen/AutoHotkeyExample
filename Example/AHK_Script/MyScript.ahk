#Requires AutoHotkey v2.0

;{ 自訂錯誤
; 檔案不存在
class FileNotExistError extends Error {
	__New(pFile){
		message := "檔案不存在" pFile
		super.__New(message, -1)
	}
}
;}

class Common {

}

class MyClass extends Common {
	;{ 回傳今天日期_yyyyMMdd
	GetToday_yyyyMMdd() => FormatTime(A_Now, "yyyyMMdd")
	;}
	;{ 呼叫外部檔案並回傳結果(會閃 cmd 視窗)
	static RunCommand(pLanguage, pFile, pParameters) {
		if not FileExist(pFile){
			throw FileNotExistError(pFile)
		}

		return	ComObject("WScript.Shell").Exec(pLanguage " " pFile " " pParameters).StdOut.ReadAll()
	}
	;}
	;{ 呼叫外部檔案並回傳結果(會閃 cmd 視窗)
	static RunWaitOne(pLanguage, pFile, pParameters) {
		if not FileExist(pFile){
			throw FileNotExistError(pFile)
		}

		return ComObject("WScript.Shell").Exec(A_ComSpec " /C " pLanguage " " pFile " " pParameters).StdOut.ReadAll()
	}
	;}
	;{ 呼叫外部檔案並回傳結果(不會閃 cmd 視窗，結果先存檔再取值)
	static RunResult(pLanguage, pFile, pParameters) {
		if not FileExist(pFile){
			throw FileNotExistError(pFile)
		}

		;指定執行結果暫存位置
		tempFile := A_ScriptDir "\" DllCall("GetCurrentProcessId") ".txt"
		;執行並輸出結果到暫存位置
		RunWait A_ComSpec " /c " pLanguage ((pLanguage = "python") ? " " : " -ExecutionPolicy Bypass ") pFile " " pParameters " >>" tempFile,,"Hide"
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
	static AddReloadHotkey(pReloadDelaySecond) {
		; 編輯器新增熱鍵把存檔觸發到重新載入腳本
		editorList := [" - SciTE4AutoHotkey", " - Notepad++", " - 記事本"]
		for index, editor in editorList {
			HotIfWinActive A_ScriptName editor
			Hotkey "~^S", ReloadOnSave.Bind(pReloadDelaySecond)
		}
		; 重新載入腳本
		ReloadOnSave(pReloadDelaySecond, ThisHotkey) {
			MyTooltip pReloadDelaySecond " 秒後重新載入腳本", 0
			Sleep pReloadDelaySecond * 1000
			Reload
		}
	}
	;}
	;{ 顯示錯誤
	static ShowError(e) {
		stackList := StrSplit(e.Stack, "`n")
		stackString := ""
		for stack in stackList {
			stackString .= stack "`n`n"
		}

		MsgBox Format("
		(
			發生錯誤：`n
			Type:`t{1}`n
			Message:`t{2}`n
			File:`t{3}`n
			Line:`t{4}`n
			What:`t{5}`n
			Stack:
			{6}
		)",
		type(e), e.Message, e.File, e.Line, e.What, stackString)
	}
	;}
}

;{ 取的螢幕資訊
GetMonitorInfo() {
	MonitorCount := MonitorGetCount()
	MonitorPrimary := MonitorGetPrimary()
	MsgBox "Monitor Count:`t" MonitorCount "`nPrimary Monitor:`t" MonitorPrimary
	Loop MonitorCount	{
		MonitorGet A_Index, &L, &T, &R, &B
		MonitorGetWorkArea A_Index, &WL, &WT, &WR, &WB
		MsgBox
		(
			"Monitor:`t#" A_Index "
			Name:`t" MonitorGetName(A_Index) "
			Left:`t" L " (" WL " work)
			Top:`t" T " (" WT " work)
			Right:`t" R " (" WR " work)
			Bottom:`t" B " (" WB " work)"
		)
	}
}
;}

;{ 使用自訂 ToolTip
;		pQuadrant： ToolTip 顯示的象限，0 為置中
MyTooltip(pText, pQuadrant := 1, pTime := 3000) {
	ToolTipOptions.Init()
	ToolTipOptions.SetFont("s48 underline italic", "Consolas")
	ToolTipOptions.SetMargins(12, 12, 12, 12)
	ToolTipOptions.SetTitle("test title" , 4)
	ToolTipOptions.SetColors("Green", "White")

	SetWinDelay -1
	; 取得工作列高度
	WinGetPos , , , &taskBarHeight, "ahk_class Shell_TrayWnd"
	; 設定 ToolTip
	ToolTip pText
	; 設定持續時間
	SetTimer () => ToolTip(), -pTime
	; 取得 ToolTip handler:
	toolTipHandler := WinExist("ahk_class tooltips_class32 ahk_pid " ProcessExist())
	; 取得 ToolTip dimensions:
	WinGetPos ,, &toolTipWidth, &toolTipHeight, toolTipHandler
	; 移動 ToolTip 到該象限
	switch pQuadrant {
		case 0:
			WinMove (A_ScreenWidth - toolTipWidth) / 2, (A_ScreenHeight - toolTipHeight) / 2, , , toolTipHandler
		case 1:
			WinMove (A_ScreenWidth - toolTipWidth), taskBarHeight, , , toolTipHandler
		case 2:
			WinMove 0, taskBarHeight, , , toolTipHandler
		case 3:
			WinMove 0, (A_ScreenHeight - toolTipHeight - taskBarHeight), , , toolTipHandler
		case 4:
			WinMove (A_ScreenWidth - toolTipWidth), (A_ScreenHeight - toolTipHeight - taskBarHeight), , , toolTipHandler
	}
	;ToolTipOptions.Reset()
}
;}