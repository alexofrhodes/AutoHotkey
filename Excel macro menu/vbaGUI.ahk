/*	Info
AUTHOR: Anastasiou Alex
		anastasioualex@gmail.com
		https://github.com/alexofrhodes
		https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
		
PURPOSE:
----------
Permanent GUI to run excel macros
Allows the use of same text file as when using EZmenu (look at my excel menu variant - vbaMENU.ahk)

SETUP these 3 parameters
---------
1. Name of your existing workbook which holds the procedures to run
2. Number of max controls per column
3. Path to a text file with a tree structure menu of said procedures
---------
example:
---------
Header1
	ProcedureName
	ProcedureName
Header2
	ProcedureName
	ProcedureName

	ProcedureName
	ProcedureName
*/

;#IfWinActive, ahk_exe EXCEL.EXE
#SingleInstance, force
SetWorkingDir, %A_ScriptDir%

WorkbookName:="ProjectStarter.xlam!" 
MenuFile:=	 A_ScriptDir . "\vba.menu" 
ItemsPerColumn:=11

;get menu headers for dropdown menu
gosub GetGuis 		; will continue to ChooseGui

;Hotkey
^+h:: 	;ctrl + shift + H

;I modified the script from the following link as a base to switch GUIS	--->		https://www.autohotkey.com/board/topic/58189-change-gui-from-dropdown-choice/

;Setup GUI with a dropdown to allow switching
ChooseGui:
{
if (ChooseGui!="") {
	ChooseGui=0
}
Gui, Submit
Gui, Destroy

/*	
Add a dropdown to the GUI
The selection change will create the corresponding menu
AltSubmit passes the element's index instead of text to the variable fo the control
*/

Gui, Add, DropDownList, section choose%ChooseGui% AltSubmit gChooseGui vChooseGui, %Guis%

;sections resets the x y position for subsequent controls
Gui, Add, Text, section 

;put text file's content into a variable
FileRead, content, %MenuFile%

;split it into an array
var:=StrSplit(content, "---")

;chose the array element matching the dropdown's selected option (matching index)
lines:= var[ChooseGui]

;Create GUI according to dropdown selection (by index)
Gosub, LoadMenu

Return
}

;get menu headers for dropdown menu
GetGuis:
{	
;put text file's content into a variable
FileRead, content, %MenuFile%

;Split into an aray of its lines
for each, line in StrSplit(content, "`n", "`r")
{
	if trim(line) =
		continue			;<--- means to skip this element of the loop
	
	FirstCharacter:= substr(line,1,1)	
	if firstCharacter in !,-,.,;
		Continue
	
	;Choose the non indented elements as headers for the menu
	if not (firstCharacter = A_Tab) 
	{
		if Guis=
		{
			Guis:= Guis . line
		}else{
			Guis:= Guis . "|" . line
		}
	}
}
return
}

;Create GUI according to dropdown selection (by index)
LoadMenu:
{
	counter:=0
	for each, line in StrSplit(lines, "`n", "`r")
	{
		if ErrorLevel
			break

		;Some formatting which allows us to use the same text file as when using EZmenu (look at my excel menu variant)	
		FirstCharacter:= substr(line,1,1)
		if FirstCharacter in !,.
			continue
		if not (firstCharacter=A_Tab)
			continue
		Position:= InStr(line, ";")
		if Position>0
			line:=SubStr(line, 1, Position - 1)
		Position:= InStr(line, "!")
		if Position>0
			line:=SubStr(line, 1, Position - 1)
		line:=trim(line)

		;How many controls to allow per column
		LimitReached:=Mod(counter, ItemsPerColumn)
		
		;alternatively to force new column leave empty line (resets counter for the rest of the controls in the new column)
		if line =
		{
			Gui, Add, Text,ys
			counter:=0
			continue
		;or if controls placed in active column reached set limit, start new column
		}else if (%LimitReached%=0)
		{
			Gui, Add, Text,ys 
		}
		
		/*
		Add a button with the caption = the text of the line, in this case 
		the name of a VBA procedure which will run on click with this script's GunExcelMacro 
		if the workbook (WorkbookName) which was set at the top is open and contains said Procedure
		*/
		Gui, Add, Button, gRunExcelMacro, %line%
		counter++
	}
	
	;Add dividing line and exit button
	Gui, Add, Text, x5 h0 w100 0x10
	Gui, Add, Button, y+10 xs gExitApp, ExitApp 
	
	;AUTHOR links
	Gui, Add, Link,y+10, <a href="https://github.com/alexofrhodes/AutoHotkey">GitHub</a> 
	Gui, Add, Link,x+40, <a href="https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg">YouTube</a> 
	
	;Gui options
	Gui, +AlwaysOnTop ;-Border +resize 
	Gui, Show	;, x%xpos% y%ypos%, Main
	return 
	
	;Close GUI when exit button pressed or ESC pressed. This doesn't stop the script's execution.
	ExitApp:
	GuiEscape:
	Gui, Destroy
	return

	
/*	NOTES
	Gui, Add, Text, x5 y5 w150 0x10  ;Horizontal Line > Etched Gray
	Gui, Add, Text, x5 y5 h150 0x11  ;Vertical Line > Etched Gray
	Gui, Add, Text, x5 y155 w150 h1 0x7  ;Horizontal Line > Black
	Gui, Add, Text, x155 y5 w1 h150 0x7  ;Vertical Line > Black

	Gui, Add, Button, gRunExcelMacro, [&1] abc ; the last part is the Sub name, [&1] and space before sub name act as accelerator, must be unique key
	Gui, Add, Button, gRunExcelMacro, abcx
	Gui, Add, Button, ys gRunExcelMacro, [&3] Place tags  ;ys starts new column
	Gui, Add, Button, gRunExcelMacro, [&4] Open Notepad
*/
}

RunExcelMacro:
{
	try {
		XL := Excel_Get()
	} catch {
		MsgBox, 16,, Can't obtain Excel! 
		return
	}
;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes
	
	;a space is allowed in the following format: [&1] MacroName
	;to allow a GUI accelerator between the braces eg. [accelerator]

	if instr(A_GuiControl, A_Space){
		StringSplit, Procedure, A_GuiControl, %A_Space%
		macro:= WorkbookName . Procedure2 
	}else{
		macro:= WorkbookName . A_GuiControl
	}
	;MsgBox, %A_GuiControl%
	;MsgBox, %procedure2%
	;MsgBox, %macro%
	
	try {
		XL.Run(macro)  
	} catch {
		MsgBox, 16,, Can't find %A_GuiControl% in the opened workbook!
	}
	;Gosub, ExitApp
	return
}

Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
	/*
		Excel_Get by jethrow (modified)
		Forum:    https://autohotkey.com/boards/viewtopic.php?f=6&t=31840
		Github:   https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Excel/Excel_Get.ahk
		
		References
		https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
		https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
		https://autohotkey.com/boards/viewtopic.php?p=134048#p134048
	*/
	
	
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

