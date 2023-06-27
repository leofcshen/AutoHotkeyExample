#Requires AutoHotkey v2.0

if WinExist("Untitled - Notepad")
    WinActivate ; 使用由 WinExist 找到的窗口
else
    WinActivate "小算盤"