; #Requires AutoHotkey >= v2.0

SetupTray

global scriptsFolder := A_ScriptDir . "\Scripts\"
global extensions := ["ahk", "exe"]

RButton:: 
{
    If WinActive("ahk_class CabinetWClass")
        {
        startTime := A_TickCount 
        KeyWait("RButton", "U")  
        keypressDuration := A_TickCount-startTime 
        if (keypressDuration > 200) 
        {
            Main
        }
        else 
        {
            Send("{RButton}")
        }        
        }
    else
        Send("{RButton}")
}

Main(){
    try
        mymenu.Delete
    MyMenu := Menu()
    MyMenu.Add "Cancel", DoNothing
    MyMenu.Add "Quit", Quit    
    MyMenu.Add
    Loop Files, scriptsFolder . "*", "D"
    {
        Submenu := Menu()
        currentFolder := A_LoopFileName
        Loop Files, scriptsFolder . currentFolder . "\*", "F" 
            {
                if HasVal(extensions, A_LoopFileExt)
                    {
                        ahkFile := A_LoopFileName
                        Submenu.Add ahkFile, MenuHandler 
                    }        
            }
        MyMenu.add currentFolder, Submenu
    }
    Loop Files, scriptsFolder . "*", "F"
    {
        if HasVal(extensions, A_LoopFileExt)
            {
                ahkFile := A_LoopFileName
                MyMenu.add ahkFile, MenuHandler
            }
    }    
    MyMenu.Show()
}

MenuHandler(Item, *) {
    Loop Files "Scripts\*", "R"  ; Recurse into subfolders.
    {
        SplitPath A_LoopFileFullPath, &name, &dir, &ext, &name_no_ext, &drive
        if (name = Item)
            {
                Run A_LoopFileFullPath
                ; MsgBox(A_LoopFileFullPath)
                return
            }
    }   
}

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}

DoNothing(*)
{
    return
}

Quit(*)
{
    ExitApp
}

