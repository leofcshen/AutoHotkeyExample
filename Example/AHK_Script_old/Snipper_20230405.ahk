; Snipper
; Fanatic Guru
;
; Version 2023 04 05
;
; #Requires AutoHotkey v2
;
; Copy Area of Screen
;
;{-----------------------------------------------
;
; AutoHotkey alternate to Windows Snipping Tool with additional features
; Creates Gui of Snip that can be manipulated onscreen
;
;	Credits:
;	The work of dozens of people inspired this script.
;	Many of them listed in the threads below:
;
;	Screen Clipping		https://www.autohotkey.com/boards/viewtopic.php?f=6&t=12088
;	Gdip				https://www.autohotkey.com/boards/viewtopic.php?t=72011
;
;}

;; INITIALIZATION - ENVIROMENT
;{-----------------------------------------------
;
#Requires AutoHotkey v2
#Warn All, Off
#SingleInstance force ; Ensures that only the last executed instance of script is running
SetWinDelay(0)
;}

;; DEFAULT SETTING - VARIABLES
;{-----------------------------------------------
;
Settings_SavePath_Image := GetFullPathName('.\Snipper - Images\')
Settings_SavePath_Image_Ext := 'PNG' ; BMP|DIB|RLE|JPG|JPEG|JPE|JFIF|GIF|TIF|TIFF|PNG
Settings_SavePath_PDF := GetFullPathName('.\Snipper - PDF\')
;}

;; INITIALIZATION - VARIABLES
;{-----------------------------------------------
;
guiSnips := Map()
;}

;; INITIALIZATION - GUI
;{-----------------------------------------------
;
ContextSnipMenu := Menu()
ContextSnipMenu.Add('COPY:  &Clipboard', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  Clipboard (with Border)', ContextSnipMenu_Handler)
ContextSnipMenu.Add('')
ContextSnipMenu.Add('COPY:  Acrobat &PDF', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  Acrobat PDF (with Border)', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  Acrobat P&DF - Saved', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  Acrobat PDF - Saved (with Border)', ContextSnipMenu_Handler)
ContextSnipMenu.Add('')
ContextSnipMenu.Add('COPY:  ' Settings_SavePath_Image_Ext ' &File', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  ' Settings_SavePath_Image_Ext ' File (with Border)', ContextSnipMenu_Handler)
ContextSnipMenu.Add('COPY:  ' Settings_SavePath_Image_Ext ' File && Outlook Email', ContextSnipMenu_Handler)
ContextSnipMenu.Add('')
ContextSnipMenu.Add('')
ContextSnipMenu.Add('CLOSE:  Snip Image', ContextSnipMenu_Handler)
;}

;; AUTO-EXECUTE
;{-----------------------------------------------
;
;TraySetIcon(A_WinDir '\system32\shell32.dll', 262) ; Scissors ; Scissors
;TraySetIcon(A_WinDir '\SystemResources\shell32.dll.mun', 260)
DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
;}

;; HOTKEYS
;{-----------------------------------------------
;
#Lbutton::		;	<-- Snip Image Only
{
	Area := SelectScreenRegion('LButton')
	If (Area.W > 8 and Area.H > 8)
		SnipArea(Area, false)
}

#^Lbutton::	;	<-- Snip Image and Copy to Clipboard
{
	Area := SelectScreenRegion('LButton')
	If (Area.W > 8 and Area.H > 8)
		SnipArea(Area, true)
}

#HotIf WinActive('SnipperWindow ahk_class AutoHotkeyGUI')

RButton::	;	<-- @@ Click for Context Menu
{
	MouseGetPos(,,&OutputVarWin)
	WinActivate('ahk_id ' OutputVarWin)
	ContextSnipMenu.Show()
}

Esc:: CloseSnip()	;	<-- @@ Close Active Snip

#HotIf
;}

;; CLASSES & FUNCTIONS - GUI
;{-----------------------------------------------
;
OnMessage(0x201, WM_LBUTTONDOWN)
WM_LBUTTONDOWN(wParam, lParam, msg, hwnd)
{
	PostMessage(0xA1, 2, , hwnd)
}

ContextSnipMenu_Handler(ItemName, ItemPos, MyMenu)
{
	Sleep 350 ; give Menu time to fade out
	Switch ItemPos
	{
		Case 1: Snip2Clipboard(Borders := false)											; Clipboard
		Case 2: Snip2Clipboard(Borders := true)												; Clipboard (with Border)
		Case 4: Snip2Clipboard(Borders := false), Clipboard2Acrobat()						; PDF
		Case 5: Snip2Clipboard(Borders := true), Clipboard2Acrobat()						; PDF (with Border)
		Case 6: Snip2Clipboard(Borders := false), Clipboard2Acrobat(Settings_SavePath_PDF)	; PDF - Saved
		Case 7: Snip2Clipboard(Borders := true), Clipboard2Acrobat(Settings_SavePath_PDF)	; PDF - Saved (with Border)
		Case 9: Snip2File(Settings_SavePath_Image)											; File
		Case 10: Snip2File(Settings_SavePath_Image, Borders := true)						; File (with Border)
		Case 11: File2Outlook(Snip2File(Settings_SavePath_Image, Borders := false))			; File and Email
		Case 14: CloseSnip																	; Snip Image
	}
}

CloseSnip()
{
	Hwnd := WinGetID('A')
	WinClose('A')
	guiSnips.Delete(Hwnd)
}
;}

;; CLASSES & FUNCTIONS
;{-----------------------------------------------
;
SelectScreenRegion(Key, Color := 'Lime', Transparent := 80)
{
	static 	guiSSR := Gui('+AlwaysOnTop -caption +Border +ToolWindow +LastFound -DPIScale')
	CoordMode('Mouse', 'Screen')
	MouseGetPos(&sX, &sY)
	WinSetTransparent(Transparent, guiSSR)
	guiSSR.BackColor := Color
	Loop
	{
		Sleep 10
		MouseGetPos(&eX, &eY)
		W := Abs(sX - eX), H := Abs(sY - eY)
		X := Min(sX, eX), Y := Min(sY, eY)
		guiSSR.Show('x' X ' y' Y ' w' W ' h' H)
	} Until !GetKeyState(Key, 'p')
	guiSSR.Hide()
	Return { X: X, Y: Y, W: W, H: H, X2: X + W, Y2: Y + H }
}

SnipArea(Area, SetClipboard := false)
{
	Global guiSnips
	GDIp.Startup()
	pBitmap := GDIp.BitmapFromScreen(Area)
	If SetClipboard
		GDIp.SetBitmapToClipboard(pBitmap)
	Hwnd := CreateLayeredWinMod(pBitmap, Area)
	guiSnips[Hwnd] := GuiFromHwnd(Hwnd)
	GDIp.DisposeImage(pBitmap)
	GDIp.Shutdown()
}

CreateLayeredWinMod(pBitmap, Area, BorderColorA := 0xff6666ff, BorderColorB := 0xffffffff)
{
	guiSnip := Gui('-Caption +E0x80000 +LastFound +ToolWindow +AlwaysOnTop +OwnDialogs', 'SnipperWindow')
	guiSnip.Show('NA')
	hbm := GDIp.CreateDIBSection(Area.W + 6, Area.H + 6), hdc := GDIp.CreateCompatibleDC(), obm := GDIp.SelectObject(hdc, hbm)
	G := GDIp.GraphicsFromHDC(hdc), GDIp.SetSmoothingMode(G, 4), GDIp.SetInterpolationMode(G, 7)
	GDIp.DrawImage(G, pBitmap, 3, 3, Area.W, Area.H)
	GDIp.DisposeImage(pBitmap)
	pPen1 := GDIp.CreatePen(BorderColorA, 3), pPen2 := GDIp.CreatePen(BorderColorB, 1)
	GDIp.DrawRectangle(G, pPen1, 1, 1, Area.W + 3, Area.H + 3)
	GDIp.DrawRectangle(G, pPen2, 1, 1, Area.W + 3, Area.H + 3)
	GDIp.DeletePen(pPen1), GDIp.DeletePen(pPen2)
	GDIp.UpdateLayeredWindow(guiSnip.hwnd, hdc, Area.X - 3, Area.Y - 3, Area.W + 6, Area.H + 6)
	GDIp.SelectObject(hdc, obm), GDIp.DeleteObject(hbm), GDIp.DeleteDC(hdc), GDIp.DeleteGraphics(G)
	Return guiSnip.hwnd
}

Snip2Clipboard(Borders := false, Hwnd?)
{
	If !IsSet(Hwnd)
		Hwnd := WinGetID('A')
	WinGetPos(&X, &Y, &W, &H, 'ahk_id ' Hwnd)
	If !Borders
		X += 3, Y += 3, W -= 6, H -= 6
	GDIp.Startup()
	pBitmap := GDIp.BitmapFromScreen({ X: X, Y: Y, W: W, H: H })
	GDIp.SetBitmapToClipboard(pBitmap)
	GDIp.DisposeImage(pBitmap)
	GDIp.Shutdown()
}

Snip2File(SavePath, Borders := false, Hwnd?)
{
	If !FileExist(SavePath)
		DirCreate(SavePath)
	If !IsSet(Hwnd)
		Hwnd := WinGetID('A')
	WinGetPos(&X, &Y, &W, &H, 'ahk_id ' Hwnd)
	If !Borders
		X += 3, Y += 3, W -= 6, H -= 6
	GDIp.Startup()
	pBitmap := GDIp.BitmapFromScreen({ X: X, Y: Y, W: W, H: H })
	TimeStamp := FormatTime(, 'yyyy_MM_dd @ HH_mm_ss')
	FileName := TimeStamp ' (' W 'x' H ').' Settings_SavePath_Image_Ext
	GDIp.SaveBitmapToFile(pBitmap, SavePath FileName)
	GDIp.DisposeImage(pBitmap)
	GDIp.Shutdown()
	Return SavePath FileName
}

Clipboard2Acrobat(SavePath := '')		; Adobe Acrobat must be installed
{
	App := ComObject('AcroExch.App')
	App.Show()
	App.MenuItemExecute('ImageConversion:Clipboard')
	If SavePath
	{
		If !FileExist(SavePath)
			DirCreate(SavePath)
		TimeStamp := FormatTime(, 'yyyy_MM_dd @ HH_mm_ss')
		FileName := TimeStamp '.PDF'
		AVDoc := App.GetActiveDoc()
		PVDoc := AVDoc.GetPDDoc()
		PDSaveIncremental := 0x0000   ;/* write changes only */
		PDSaveFull := 0x0001   ;/* write entire file */
		PDSaveCopy := 0x0002   ;/* write copy w/o affecting current state */
		PDSaveLinearized := 0x0004   ;/* write the file linearized for */
		PDSaveBinaryOK := 0x0010   ;/* OK to store binary in file */
		PDSaveCollectGarbage := 0x0020   ;/* perform garbage collection on */
		PVDoc.Save(PDSaveFull | PDSaveLinearized, SavePath FileName)
	}
}

File2Outlook(File, URL := '')	; Outlook must be installed
{
	TimeStamp := FormatTime(RegExReplace(File, '^.*\\|[_ @]|\(.*$'), "dddd MMMM d, yyyy 'at' h:mm:ss tt")
	Try
		IsObject(MailItem := ComObjActive('Outlook.Application').CreateItem(olMailItem := 0)) ; Get the Outlook application object if Outlook is open
	Catch
		MailItem := ComObject('Outlook.Application').CreateItem(olMailItem := 0) ; Create if Outlook is not open
	MailItem.BodyFormat := (olFormatHTML := 2)
	;~ MailItem.TO :='somejunkemail@yahoo.com'
	;~ MailItem.CC :='somejunkemail@yahoo.com'
	MailItem.Subject := 'Screenshot taken: ' TimeStamp ; Subject line of email
	HTMLBody := "<H2 style='BACKGROUND-COLOR: red'><br></H2><HTML>Please find attached the screenshot taken on " TimeStamp "<br><br>"
	If URL
		HTMLBody .= "<span style='color:black'>The image can also be accessd here: <a href=" URL ">" URL "</a><br><br></span>"
	HTMLBody .= '</HTML>'
	MailItem.HTMLBody := HTMLBody
	MailItem.Attachments.Add(File)
	MailItem.Display
}

GetFullPathName(path)
{
	cc := DllCall('GetFullPathNameW', 'str', path, 'uint', 0, 'ptr', 0, 'ptr', 0, 'uint')
	buf := Buffer(cc * 2)
	DllCall('GetFullPathNameW', 'str', path, 'uint', cc, 'ptr', buf, 'ptr', 0, 'uint')
	Return StrGet(buf)
}
;}

;; LIBRARIES - GDIp
;{-----------------------------------------------
;
;{ GDIp Class - Select GDIp library functions converted to a class specifically for this script
#DllLoad 'GdiPlus'
Class GDIp
{
	;{ Startup
	Static Startup()
	{
		If (this.HasProp("Token"))
			Return
		input := Buffer((A_PtrSize = 8) ? 24 : 16, 0)
		NumPut("UInt", 1, input)
		DllCall("gdiplus\GdiplusStartup", "UPtr*", &pToken := 0, "UPtr", input.ptr, "UPtr", 0)
		this.Token := pToken
	}
	;}
	;{ Shutdown
	Static Shutdown()
	{
		If (this.HasProp("Token"))
			DllCall("Gdiplus\GdiplusShutdown", "UPtr", this.DeleteProp("Token"))
	}
	;}
	;{ BitmapFromScreen
	Static BitmapFromScreen(Area)
	{
		chdc := this.CreateCompatibleDC()
		hbm := this.CreateDIBSection(Area.W, Area.H, chdc)
		obm := this.SelectObject(chdc, hbm)
		hhdc := this.GetDC()
		this.BitBlt(chdc, 0, 0, Area.W, Area.H, hhdc, Area.X, Area.Y)
		this.ReleaseDC(hhdc)
		pBitmap := this.CreateBitmapFromHBITMAP(hbm)
		this.SelectObject(chdc, obm), this.DeleteObject(hbm), this.DeleteDC(hhdc), this.DeleteDC(chdc)
		Return pBitmap
	}
	;}
	;{ SetBitmapToClipboard
	Static SetBitmapToClipboard(pBitmap)
	{
		off1 := A_PtrSize = 8 ? 52 : 44
		off2 := A_PtrSize = 8 ? 32 : 24

		pid := DllCall("GetCurrentProcessId", "uint")
		hwnd := WinExist("ahk_pid " . pid)
		r1 := DllCall("OpenClipboard", "UPtr", hwnd)
		If !r1
			Return -1

		hBitmap := this.CreateHBITMAPFromBitmap(pBitmap, 0)
		If !hBitmap
		{
			DllCall("CloseClipboard")
			Return -3
		}

		r2 := DllCall("EmptyClipboard")
		If !r2
		{
			this.DeleteObject(hBitmap)
			DllCall("CloseClipboard")
			Return -2
		}

		oi := Buffer((A_PtrSize = 8) ? 104 : 84, 0)
		DllCall("GetObject", "UPtr", hBitmap, "int", oi.size, "UPtr", oi.ptr)
		hdib := DllCall("GlobalAlloc", "uint", 2, "UPtr", 40 + NumGet(oi, off1, "UInt"), "UPtr")
		pdib := DllCall("GlobalLock", "UPtr", hdib, "UPtr")
		DllCall("RtlMoveMemory", "UPtr", pdib, "UPtr", oi.ptr + off2, "UPtr", 40)
		DllCall("RtlMoveMemory", "UPtr", pdib + 40, "UPtr", NumGet(oi, off2 - A_PtrSize, "UPtr"), "UPtr", NumGet(oi, off1, "UInt"))
		DllCall("GlobalUnlock", "UPtr", hdib)
		this.DeleteObject(hBitmap)
		r3 := DllCall("SetClipboardData", "uint", 8, "UPtr", hdib) ; CF_DIB = 8
		DllCall("CloseClipboard")
		DllCall("GlobalFree", "UPtr", hdib)
		E := r3 ? 0 : -4    ; 0 - success
		Return E
	}
	;}
	;{ CreateCompatibleDC
	Static CreateCompatibleDC(hdc := 0)
	{
		Return DllCall("CreateCompatibleDC", "UPtr", hdc)
	}
	;}
	;{ CreateDIBSection
	Static CreateDIBSection(w, h, hdc := "", bpp := 32, &ppvBits := 0, Usage := 0, hSection := 0, Offset := 0)
	{
		hdc2 := hdc ? hdc : this.GetDC()
		bi := Buffer(40, 0)
		NumPut("UInt", 40, bi, 0)
		NumPut("UInt", w, bi, 4)
		NumPut("UInt", h, bi, 8)
		NumPut("UShort", 1, bi, 12)
		NumPut("UShort", bpp, bi, 14)
		NumPut("UInt", 0, bi, 16)

		hbm := DllCall("CreateDIBSection"
			, "UPtr", hdc2
			, "UPtr", bi.ptr    ; BITMAPINFO
			, "uint", Usage
			, "UPtr*", &ppvBits
			, "UPtr", hSection
			, "uint", Offset, "UPtr")

		If !hdc
			this.ReleaseDC(hdc2)
		Return hbm
	}
	;}
	;{ SelectObject
	Static SelectObject(hdc, hgdiobj)
	{
		Return DllCall("SelectObject", "UPtr", hdc, "UPtr", hgdiobj)
	}
	;}
	;{ BitBlt
	Static BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, raster := "")
	{
		Return DllCall("gdi32\BitBlt"
			, "UPtr", ddc
			, "int", dx, "int", dy
			, "int", dw, "int", dh
			, "UPtr", sdc
			, "int", sx, "int", sy
			, "uint", raster ? raster : 0x00CC0020)
	}
	;}
	;{ CreateBitmapFromHBITMAP
	Static CreateBitmapFromHBITMAP(hBitmap, hPalette := 0)
	{
		DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "UPtr", hBitmap, "UPtr", hPalette, "UPtr*", &pBitmap := 0)
		Return pBitmap
	}
	;}
	;{ CreateHBITMAPFromBitmap
	Static CreateHBITMAPFromBitmap(pBitmap, Background := 0xffffffff)
	{
		DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "UPtr", pBitmap, "UPtr*", &hBitmap := 0, "int", Background)
		Return hBitmap
	}
	;}
	;{ DeleteObject
	Static DeleteObject(hObject)
	{
		Return DllCall("DeleteObject", "UPtr", hObject)
	}
	;}
	;{ ReleaseDC
	Static ReleaseDC(hdc, hwnd := 0)
	{
		Return DllCall("ReleaseDC", "UPtr", hwnd, "UPtr", hdc)
	}
	;}
	;{ DeleteDC
	Static DeleteDC(hdc)
	{
		Return DllCall("DeleteDC", "UPtr", hdc)
	}
	;}
	;{ DisposeImage
	Static DisposeImage(pBitmap, noErr := 0)
	{
		If (StrLen(pBitmap) <= 2 && noErr = 1)
			Return 0

		r := DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
		If (r = 2 || r = 1) && (noErr = 1)
			r := 0
		Return r
	}
	;}
	;{ GetDC
	Static GetDC(hwnd := 0)
	{
		Return DllCall("GetDC", "UPtr", hwnd)
	}
	;}
	;{ GetDCEx
	Static GetDCEx(hwnd, flags := 0, hrgnClip := 0)
	{
		Return DllCall("GetDCEx", "UPtr", hwnd, "UPtr", hrgnClip, "int", flags)
	}
	;}
	;{ GetWindowRect
	Static GetWindowRect(hwnd, &W, &H)
	{
		rect := Buffer(16, 0)
		er := DllCall("dwmapi\DwmGetWindowAttribute"
			, "UPtr", hwnd        ; HWND  hwnd
			, "UInt", 9           ; DWORD dwAttribute (DWMWA_EXTENDED_FRAME_BOUNDS)
			, "UPtr", rect.ptr    ; PVOID pvAttribute
			, "UInt", rect.size   ; DWORD cbAttribute
			, "UInt")             ; HRESULT

		If er
			DllCall("GetWindowRect", "UPtr", hwnd, "UPtr", rect.ptr, "UInt")

		r := {}
		r.x1 := NumGet(rect, 0, "Int"), r.y1 := NumGet(rect, 4, "Int")
		r.x2 := NumGet(rect, 8, "Int"), r.y2 := NumGet(rect, 12, "Int")
		r.w := Abs(Max(r.x1, r.x2) - Min(r.x1, r.x2))
		r.h := Abs(Max(r.y1, r.y2) - Min(r.y1, r.y2))
		W := r.w, H := r.h
		Return r
	}
	;}
	;{ GraphicsFromHDC
	Static GraphicsFromHDC(hDC, hDevice := "", InterpolationMode := "", SmoothingMode := "", PageUnit := "", CompositingQuality := "")
	{
		If hDevice
			DllCall("Gdiplus\GdipCreateFromHDC2", "UPtr", hDC, "UPtr", hDevice, "UPtr*", &pGraphics := 0)
		Else
			DllCall("gdiplus\GdipCreateFromHDC", "UPtr", hDC, "UPtr*", &pGraphics := 0)

		If pGraphics
		{
			If (InterpolationMode != "")
				this.SetInterpolationMode(pGraphics, InterpolationMode)
			If (SmoothingMode != "")
				this.SetSmoothingMode(pGraphics, SmoothingMode)
			If (PageUnit != "")
				this.SetPageUnit(pGraphics, PageUnit)
			If (CompositingQuality != "")
				this.SetCompositingQuality(pGraphics, CompositingQuality)
		}

		Return pGraphics
	}
	;}
	;{ SetInterpolationMode
	Static SetInterpolationMode(pGraphics, InterpolationMode)
	{
		Return DllCall("gdiplus\GdipSetInterpolationMode", "UPtr", pGraphics, "int", InterpolationMode)
	}
	;}
	;{ SetSmoothingMode
	Static SetSmoothingMode(pGraphics, SmoothingMode)
	{
		Return DllCall("gdiplus\GdipSetSmoothingMode", "UPtr", pGraphics, "int", SmoothingMode)
	}
	;}
	;{ SetPageUnit
	Static SetPageUnit(pGraphics, Unit)
	{
		Return DllCall("gdiplus\GdipSetPageUnit", "UPtr", pGraphics, "int", Unit)
	}
	;}
	;{ SetCompositingQuality
	Static SetCompositingQuality(pGraphics, CompositionQuality)
	{
		Return DllCall("gdiplus\GdipSetCompositingQuality", "UPtr", pGraphics, "int", CompositionQuality)
	}
	;}
	;{ DrawImage
	Static DrawImage(pGraphics, pBitmap, dx := "", dy := "", dw := "", dh := "", sx := "", sy := "", sw := "", sh := "", Matrix := 1, Unit := 2, ImageAttr := 0)
	{
		usrImageAttr := 0
		If !ImageAttr
		{
			If !IsNumber(Matrix)
				ImageAttr := this.SetImageAttributesColorMatrix(Matrix)
			Else If (Matrix != 1)
				ImageAttr := this.SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")
		} Else usrImageAttr := 1

		If (dx != "" && dy != "" && dw = "" && dh = "" && sx = "" && sy = "" && sw = "" && sh = "")
		{
			sx := sy := 0
			sw := dw := this.GetImageWidth(pBitmap)
			sh := dh := this.GetImageHeight(pBitmap)
		} Else If (sx = "" && sy = "" && sw = "" && sh = "")
		{
			If (dx = "" && dy = "" && dw = "" && dh = "")
			{
				sx := dx := 0, sy := dy := 0
				sw := dw := this.GetImageWidth(pBitmap)
				sh := dh := this.GetImageHeight(pBitmap)
			} Else
			{
				sx := sy := 0
				this.GetImageDimensions(pBitmap, &sw, &sh)
			}
		}

		_E := DllCall("gdiplus\GdipDrawImageRectRect"
			, "UPtr", pGraphics
			, "UPtr", pBitmap
			, "float", dx, "float", dy
			, "float", dw, "float", dh
			, "float", sx, "float", sy
			, "float", sw, "float", sh
			, "int", Unit
			, "UPtr", ImageAttr ? ImageAttr : 0
			, "UPtr", 0, "UPtr", 0)

		If (ImageAttr && usrImageAttr != 1)
			this.DisposeImageAttributes(ImageAttr)

		Return _E
	}
	;}
	;{ CreateImageAttributes
	Static CreateImageAttributes()
	{
		DllCall("gdiplus\GdipCreateImageAttributes", "UPtr*", &ImageAttr := 0)
		Return ImageAttr
	}
	;}
	;{ DisposeImageAttributes
	Static DisposeImageAttributes(ImageAttr)
	{
		Return DllCall("gdiplus\GdipDisposeImageAttributes", "UPtr", ImageAttr)
	}
	;}
	;{ SetImageAttributesColorMatrix
	Static SetImageAttributesColorMatrix(clrMatrix, ImageAttr := 0, grayMatrix := 0, ColorAdjustType := 1, fEnable := 1, ColorMatrixFlag := 0)
	{
		GrayscaleMatrix := 0

		If (StrLen(clrMatrix) < 5 && ImageAttr)
			Return -1

		If StrLen(clrMatrix) < 5
			Return

		ColourMatrix := Buffer(100, 0)
		Matrix := RegExReplace(RegExReplace(clrMatrix, "^[^\d-\.]+([\d\.])", "$1", , 1), "[^\d-\.]+", "|")
		Matrix := StrSplit(Matrix, "|")
		Loop 25
		{
			M := (Matrix[A_Index] != "") ? Matrix[A_Index] : Mod(A_Index - 1, 6) ? 0 : 1
			NumPut("Float", M, ColourMatrix, (A_Index - 1) * 4)
		}

		Matrix := ""
		Matrix := RegExReplace(RegExReplace(grayMatrix, "^[^\d-\.]+([\d\.])", "$1", , 1), "[^\d-\.]+", "|")
		Matrix := StrSplit(Matrix, "|")
		If (StrLen(Matrix) > 2 && ColorMatrixFlag = 2)
		{
			GrayscaleMatrix := Buffer(100, 0)
			Loop 25
			{
				M := (Matrix[A_Index] != "") ? Matrix[A_Index] : Mod(A_Index - 1, 6) ? 0 : 1
				NumPut("Float", M, GrayscaleMatrix, (A_Index - 1) * 4)
			}
		}

		If !ImageAttr
		{
			created := 1
			ImageAttr := this.CreateImageAttributes()
		}

		E := DllCall("gdiplus\GdipSetImageAttributesColorMatrix"
			, "UPtr", ImageAttr
			, "int", ColorAdjustType
			, "int", fEnable
			, "UPtr", ColourMatrix.ptr
			, "UPtr", GrayscaleMatrix ? GrayscaleMatrix.ptr : 0
			, "int", ColorMatrixFlag)

		E := created = 1 ? ImageAttr : E
		Return E
	}
	;}
	;{ GetImageDimensions
	Static GetImageDimensions(pBitmap, &Width, &Height)
	{
		If StrLen(pBitmap) < 3
			Return -1

		Width := 0, Height := 0
		E := this.GetImageDimension(pBitmap, &Width, &Height)
		Width := Round(Width)
		Height := Round(Height)
		Return E
	}
	;}
	;{ GetImageDimension
	Static GetImageDimension(pBitmap, &w, &h)
	{
		Return DllCall("gdiplus\GdipGetImageDimension", "UPtr", pBitmap, "float*", &w := 0, "float*", &h := 0)
	}
	;}
	;{ GetImageWidth
	Static GetImageWidth(pBitmap)
	{
		DllCall("gdiplus\GdipGetImageWidth", "UPtr", pBitmap, "uint*", &Width := 0)
		Return Width
	}
	;}
	;{ GetImageHeight
	Static GetImageHeight(pBitmap)
	{
		DllCall("gdiplus\GdipGetImageHeight", "UPtr", pBitmap, "uint*", &Height := 0)
		Return Height
	}
	;}
	;{ DrawRectangle
	Static DrawRectangle(pGraphics, pPen, x, y, w, h)
	{
		Return DllCall("gdiplus\GdipDrawRectangle", "UPtr", pGraphics, "UPtr", pPen, "float", x, "float", y, "float", w, "float", h)
	}
	;}
	;{ UpdateLayeredWindow
	Static UpdateLayeredWindow(hwnd, hdc, x := "", y := "", w := "", h := "", Alpha := 255)
	{
		If ((x != "") && (y != ""))
			pt := Buffer(8, 0), NumPut("UInt", x, pt, 0), NumPut("UInt", y, pt, 4)

		If (w = "") || (h = "")
			this.GetWindowRect(hwnd, &w, &h)

		Return DllCall("UpdateLayeredWindow"
			, "UPtr", hwnd                                  ; layered window hwnd
			, "UPtr", 0                                     ; hdcDst (screen) - usually 0
			, "UPtr", ((x = "") && (y = "")) ? 0 : pt.ptr   ; POINT x,y of layered window
			, "int64*", w | h << 32                          ; SIZE w,h of layered window
			, "UPtr", hdc                                   ; hdcSrc - source bitmap to be drawn on to layered window - NULL if not changing
			, "int64*", 0                                ; x,y offset of bitmap to be drawn
			, "uint", 0                                  ; crKey - bgcolor to use?  meaningless when using full alpha
			, "UInt*", Alpha << 16 | 1 << 24                   ;
			, "uint", 2)
	}
	;}
	;{ CreatePen
	Static CreatePen(ARGB, w, Unit := 2)
	{
		E := DllCall("gdiplus\GdipCreatePen1", "UInt", ARGB, "float", w, "int", Unit, "UPtr*", &pPen := 0)
		Return pPen
	}
	;}
	;{ DeletePen
	Static DeletePen(pPen)
	{
		Return DllCall("gdiplus\GdipDeletePen", "UPtr", pPen)
	}
	;}
	;{ DeleteGraphics
	Static DeleteGraphics(pGraphics)
	{
		Return DllCall("gdiplus\GdipDeleteGraphics", "UPtr", pGraphics)
	}
	;}
	;{ SaveBitmapToFile
	Static SaveBitmapToFile(pBitmap, sOutput, Quality := 75, toBase64 := 0)
	{
		_p := 0

		SplitPath sOutput, , , &Extension
		If !RegExMatch(Extension, "^(?i:BMP|DIB|RLE|JPG|JPEG|JPE|JFIF|GIF|TIF|TIFF|PNG)$")
			Return -1

		Extension := "." Extension
		DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", &nCount := 0, "uint*", &nSize := 0)
		ci := Buffer(nSize)
		DllCall("gdiplus\GdipGetImageEncoders", "uint", nCount, "uint", nSize, "UPtr", ci.ptr)
		If !(nCount && nSize)
			Return -2

		Static IsUnicode := StrLen(Chr(0xFFFF))
		If (IsUnicode)
		{
			StrGet_Name := "StrGet"
			Loop nCount
			{
				sString := %StrGet_Name%(NumGet(ci, (idx := (48 + 7 * A_PtrSize) * (A_Index - 1)) + 32 + 3 * A_PtrSize, "UPtr"), "UTF-16")
				If !InStr(sString, "*" Extension)
					Continue

				pCodec := ci.ptr + idx
				Break
			}
		} Else
		{
			Loop nCount
			{
				Location := NumGet(ci, 76 * (A_Index - 1) + 44, "UPtr")
				nSize := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "uint", 0, "int", 0, "uint", 0, "uint", 0)
				sString := Buffer(nSize)
				DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "str", sString, "int", nSize, "uint", 0, "uint", 0)
				If !InStr(sString, "*" Extension)
					Continue

				pCodec := ci.ptr + 76 * (A_Index - 1)
				Break
			}
		}

		If !pCodec
			Return -3

		If (Quality != 75)
		{
			Quality := (Quality < 0) ? 0 : (Quality > 100) ? 100 : Quality
			If (Quality > 90 && toBase64 = 1)
				Quality := 90

			If RegExMatch(Extension, "^\.(?i:JPG|JPEG|JPE|JFIF)$")
			{
				DllCall("gdiplus\GdipGetEncoderParameterListSize", "UPtr", pBitmap, "UPtr", pCodec, "uint*", &nSize)
				EncoderParameters := Buffer(nSize, 0)
				DllCall("gdiplus\GdipGetEncoderParameterList", "UPtr", pBitmap, "UPtr", pCodec, "uint", nSize, "UPtr", EncoderParameters.ptr)
				nCount := NumGet(EncoderParameters, "UInt")
				Loop nCount
				{
					elem := (24 + A_PtrSize) * (A_Index - 1) + 4 + (pad := (A_PtrSize = 8) ? 4 : 0)
					If (NumGet(EncoderParameters, elem + 16, "UInt") = 1) && (NumGet(EncoderParameters, elem + 20, "UInt") = 6)
					{
						_p := elem + EncoderParameters.ptr - pad - 4
						NumPut(Quality, NumGet(NumPut(4, NumPut(1, _p + 0, "UPtr") + 20, "UInt"), "UPtr"), "UInt")
						Break
					}
				}
			}
		}

		If (toBase64 = 1)
		{
			; part of the function extracted from ImagePut by iseahound
			; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=76301&sid=bfb7c648736849c3c53f08ea6b0b1309
			DllCall("ole32\CreateStreamOnHGlobal", "UPtr", 0, "int", true, "UPtr*", &pStream := 0)
			_E := DllCall("gdiplus\GdipSaveImageToStream", "UPtr", pBitmap, "UPtr", pStream, "UPtr", pCodec, "uint", _p)
			If _E
				Return -6

			DllCall("ole32\GetHGlobalFromStream", "UPtr", pStream, "uint*", &hData)
			pData := DllCall("GlobalLock", "UPtr", hData, "UPtr")
			nSize := DllCall("GlobalSize", "uint", pData)

			bin := Buffer(nSize, 0)
			DllCall("RtlMoveMemory", "UPtr", bin.ptr, "UPtr", pData, "uptr", nSize)
			DllCall("GlobalUnlock", "UPtr", hData)
			ObjRelease(pStream)
			DllCall("GlobalFree", "UPtr", hData)

			; Using CryptBinaryToStringA saves about 2MB in memory.
			DllCall("Crypt32.dll\CryptBinaryToStringA", "UPtr", bin.ptr, "uint", nSize, "uint", 0x40000001, "UPtr", 0, "uint*", &base64Length := 0)
			base64 := Buffer(base64Length, 0)
			_E := DllCall("Crypt32.dll\CryptBinaryToStringA", "UPtr", bin.ptr, "uint", nSize, "uint", 0x40000001, "UPtr", &base64, "uint*", base64Length)
			If !_E
				Return -7

			bin := Buffer(0)
			Return StrGet(base64, base64Length, "CP0")
		}

		_E := DllCall("gdiplus\GdipSaveImageToFile", "UPtr", pBitmap, "WStr", sOutput, "UPtr", pCodec, "uint", _p)
		Return _E ? -5 : 0
	}
	;}
}
;}
;}