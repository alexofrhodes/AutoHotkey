#IfWinActive, ahk_exe EXCEL.EXE

WorkbookName:="ProjectStarter.xlam!" ; replace this as needed

;ctrl & h
^h:: 
MouseGetPos, xpos, ypos, window

Gui, Add, Button, gRunExcelMacro, [&1] abc ; the last part is the Sub name, [&1] and space before sub name act as accelerator, must be unique key
Gui, Add, Button, gRunExcelMacro, abcx
;Gui, Add, Button, ys gRunExcelMacro, [&3] Place tags  ;ys starts new column
;Gui, Add, Button, gRunExcelMacro, [&4] Open Notepad

Gui, Add, Text
Gui, Add, Button, xs gExitApp, ExitApp ;xs starts new row

Gui, -Border
Gui, Show, x%xpos% y%ypos%
return 

ExitApp:
GuiEscape:
 ;MsgBox escape pressed
Gui, Destroy
return

RunExcelMacro:
try {
	XL := Excel_Get()
} catch {
	MsgBox, 16,, Can't obtain Excel! 
	return
}
;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes


if instr(A_GuiControl, A_Space){
	StringSplit, Procedure, A_GuiControl, %A_Space%
	macro:= WorkbookName . Procedure2 
}else{
	macro:= WorkbookName . A_GuiControl
}
;MsgBox, %A_GuiControl%
;MsgBox, %procedure2%
MsgBox, %macro%

try {
	XL.Run(macro)  
} catch {
	MsgBox, 16,, Can't find %A_GuiControl% in the opened workbook!
}

;Gosub, ExitApp
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

; References
;   https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
;   https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
;   https://autohotkey.com/boards/viewtopic.php?p=134048#p134048
