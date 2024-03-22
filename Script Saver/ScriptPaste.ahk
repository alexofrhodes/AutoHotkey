#SingleInstance Force
#Include FolderStructure.ahk

global MyMenu
global BaseFolder := A_ScriptDir . "\Snippets"
global extensions := ["txt"]

#Include customTray.ahk
SetupTray()

;-----------------------
;Listen to hotkeys if vbeditor is active window
;-----------------------

	; #HotIf WinActive("ahk_class wndclass_desked_gsk")
	
;-----------------------
;Long press right button
;-----------------------

	; RButton:: 
	; {
	; 	startTime := A_TickCount 
	; 	KeyWait("RButton", "U")  
	; 	keypressDuration := A_TickCount-startTime 
	; 	if (keypressDuration > 200) 
	; 	{
	; 		Main()
	; 	}
	; 	else 
	; 	{
	; 		Send("{RButton}")
	; 	}        
	; }

;-----------------------	
;Double press ctrl
;-----------------------

~Ctrl Up:: 
{
	If A_ThisHotkey = A_PriorHotkey && A_TimeSincePriorHotkey < 400
		Main()
}

Main(){
	global
	try	
		myMenu.Delete
	myMenu:= Menu()
	MyMenu.Add("Cancel", DoNothing)
	MyMenu.SetIcon("Cancel","icons\cancel.ico")
	MyMenu.Add

	allfilePaths:= AddFolderStructureToMenu(myMenu, BaseFolder, extensions, "theHandlerFunction")																				
	myMenu.Show()
}

theHandlerFunction(filePath, *) {
	;MsgBox("File selected is" filePath)

	if GetKeyState("Ctrl"){
		;if ctrl is pressed
		Run 'edit '  filePath
	}else{
		A_Clipboard := FileRead(filePath)
		Send("{Ctrl down}v{Ctrl up}")			
	}
}



DoNothing(*){
	return
}

