#noEnv
#singleinstance, force
sendMode input
setWorkingDir, % a_scriptDir

/* NOTES
	----------
	replace WorkbookName:="ProjectStarter.xlam!" with your workbook where the macros are
	replace your sub names
	duplicate the hotkey sections as needed (each hotkey must be unique)
*/

;#IfWinActive ahk_class wndclass_desked_gsk ;if vbeditor window is active
	
WorkbookName:="'vbArc-Addin.xlsm'!" 

return

;control 1
^1::
item := "test1" ;"YOUR_SUB_NAME"
Gosub,RunExcelMacro
return

^2:: ;alt 2
item := "test2" ;"YOUR_SUB_NAME"
Gosub,RunExcelMacro
return


RunExcelMacro:
macro:= WorkbookName . item

;check if macro exists in workbook

; If Excel_Get fails it returns an error message instead of an object.
XL := Excel_Get() 
if !IsObject(XL)  {
	MsgBox, 16, Excel_Get Error, % XL
	return
}else{
;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes
}

try {
	XL.Run(macro)  
} catch {
	MsgBox, 16,, Can't find %item% in %WorkbookName%
}
return


; Excel_Get by jethrow (modified)
; Forum:    https://autohotkey.com/boards/viewtopic.php?f=6&t=31840
; Github:   https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Excel/Excel_Get.ahk

Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
	static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
	WinGetClass, WinClass, %WinTitle%
	if !(WinClass == "XLMAIN")
		return "Window class mismatch."
	ControlGet, hwnd, hwnd,, Excel7%Excel7#%, %WinTitle%
	if (ErrorLevel)
		return "Error accessing the control hWnd."
	VarSetCapacity(IID_IDispatch, 16)
	NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
	if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc) != 0
		return "Error calling AccessibleObjectFromWindow."
	window := ComObject(9, pacc, 1)
	if ComObjType(window) != 9
		return "Error wrapping the window object."
	Loop
		try return window.Application
	catch e
		if SubStr(e.message, 1, 10) = "0x80010001"
			ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
	else
		return "Error accessing the application object."
}

