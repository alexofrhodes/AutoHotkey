SendMode, Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance, Force
SetKeyDelay, 50

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"

nTime := A_TickCount
WinGetClass, sClass, A

;Custom Tray Icon
I_Icon = ScriptSaver.ico ;<a href="https://www.flaticon.com/free-icons/script" title="script icons">Script icons created by Pixel perfect - Flaticon</a>
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

;Event for Tray icon left click
OnMessage(0x404, "AHK_NOTIFYICON")
AHK_NOTIFYICON(wParam, lParam)
{
    if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
        gosub CreateGui
        return 0
    }
}

CreateGui:
{
    gui, destroy
    SettingsFile = config.ini
    Gosub, ReadIni ;If the file doesn't exist yet, the default lists will be used.

    Gui , Add, ComboBox, w200 hwndComboPath vTargetPath gSaveSettings, %MyPaths%
    Gui , Font, Bold
    Gui , Add, Button, ys w18 h18 gAddPaths, + ;Buttons go to their g-label when clicked.
    Gui , Add, Button, ys w18 h18 gRemPaths, -
    Gui , Font, Normal
    Gui , Add, ComboBox, w200 hwndComboExtension xs section vTargetExtension gSaveSettings, %MyExtensions%
    Gui , Font, Bold
    Gui , Add, Button, ys w18 h18 gAddExtensions, +
    Gui , Add, Button, ys w18 h18 gRemExtensions, -
    Gui , Font, Normal

    gui, add, Checkbox, xs section gSaveSettings vChAskFileName, Ask for file name
    gui, add, Checkbox, ys gSaveSettings vChEditAfterSave, Edit after save  

    Gui , Add, Button, w200 xs y+20 section gSaveScript, Save Script

	;AUTHOR links
	Gui, Add, Link,ys y+-35 section, 	<a href="https://github.com/alexofrhodes/">					        GitHub		</a> 
	Gui, Add, Link,xs, 			        <a href="https://alexofrhodes.github.io/">							Blog		</a> 
	Gui, Add, Link,xs, 			        <a href="https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg">	YouTube		</a> 

    iniread, DefaultPath, Config.ini, Settings, Path
    iniread, DefaultExtension, Config.ini, Settings, Extension

    GuiControl, Text, %ComboPath%, %DefaultPath%
    GuiControl, Text, %ComboExtension%, %DefaultExtension%

    iniread, DefaultEditAfterSave, Config.ini, Settings, EditAfterSave
    iniread, DefaultAskForName, Config.ini, Settings, AskForName
    
    GuiControl, , ChEditAfterSave, %DefaultEditAfterSave%
    GuiControl, , ChAskForName, %DefaultAskForName%

    gui, +AlwaysOnTop
    Gui , Show, , Script Saver 
    Return
}

GuiEscape:
GuiClose:
{
    gosub, SaveSettings
    gui,hide
    return
}

SaveSettings:
{
    Gui, Submit, NoHide ;Get the new type.   
    IniWrite, %TargetPath%, Config.ini, Settings, Path
    IniWrite, %TargetExtension%, Config.ini, Settings, Extension   
    IniWrite, %ChEditAfterSave%, Config.ini, Settings, EditAfterSave   
    IniWrite, %ChAskFileName%, Config.ini, Settings, AskFileName  

    if A_ThisHotkey = "GuiClose"
        gui, hide
    Return
}

AddPaths:
{
    Gui, Submit, NoHide ;Get the new type.
    MyPaths .= "|" . TargetPath ;Add a pipe and the new type onto the list.
    GuiControl,, TargetPath, %TargetPath% ;Add the new type to the GUI.
    Gosub, SaveSettings
    Return
}

RemPaths:
{
    Gui, Submit, NoHide
    Temp =
    Loop, Parse, MyPaths, | ;Loop over each item in the list.
    {
        If !A_LoopField Or (A_LoopField = TargetPath) ;If this is the item you want to delete (or if it's blank), skip it.
            Continue
        Temp .= "|" . A_LoopField ;Put all the other items back into a list.
    }
    MyPaths := Temp
    GuiControl,, TargetPath, %MyPaths% ;Update the GUI with the new list.
    Gosub, SaveSettings
    Return
}

AddExtensions:
{
    Gui, Submit, NoHide
    MyExtensions .= "|" . Extension
    GuiControl,, Extension, %Extension%
    Gosub, SaveSettings
    Return
}

RemExtensions:
{
    Gui, Submit, NoHide
    Temp =
    Loop, Parse, MyExtensions, |
    {
        If !A_LoopField Or (A_LoopField = Extension)
            Continue
        Temp .= "|" . A_LoopField
    }
    MyExtensions := Temp
    GuiControl,, Extension, %MyExtensions%
    Gosub, SaveSettings
    Return
}

ReadIni:
{
    IniRead, MyPaths, %SettingsFile%, Settings, MyPaths, VBA|AHK
    IniRead, MyExtensions, %SettingsFile%, Settings, MyExtensions, TXT|AHK
    If (InStr(MyPaths, "|") = 1) ;This just makes sure there is not a pipe at the beginning so that there's not a blank space at the top of the list.
        MyPaths := SubStr(MyPaths, 2)
    If (InStr(MyExtensions, "|") = 1)
        MyExtensions := SubStr(MyExtensions, 2)
    Return
}

SaveScript:
{
    gui, hide
    Gui, Submit, NoHide ;Get the new type.   

    savedClipboard := ClipboardAll	; Backup data from clipboard.
    Clipboard := ""
	While !Clipboard	; Try to copy selected (it may fail, and then we have to retry).
        {
            Send, {Ctrl Down}c{Ctrl Up}
            Sleep, 5
        }

    textToSave := Clipboard			
    Clipboard := savedClipboard
    sURL := GetActiveBrowserURL()

    Extension := TargetExtension
    saveToFolder := TargetPath
    
    if %saveToFolder% := "" 
        FileSelectFolder, saveToFolder, *%A_ScriptDir% 
    if %saveToFolder% := "" 						
        return
    FileCreateDir, %saveToFolder%
    
    ; RegRead, editor, HKCR, AutoHotkeyScript\Shell\Edit\Command	; Get user's editor assigned for the filetype.
    ; editor := StrReplace(editor, "%SystemDrive%", SystemDrive)
    
	If (ChAskFileName = 1)
        {
            InputBox, filename, Enter the name for the script's file:,,, 300, 100
            If ErrorLevel	; Flush the variable's contents in case user canceled or closed the input box.
                filename := ""
        }

    saveTofile := (saveToFolder ? saveToFolder :A_ScriptDir "\" %saveToFolder%) "\" (filename ? filename : A_Now) "." TargetExtension

    FileAppend, %sURL% `n %textToSave%, %saveTofile%, UTF-8	 ; Save selected into a file next to this script.

    if (ChEditAfterSave = 1)
        { 
            ; Run, % (editor ? StrReplace(editor, "%1", """" saveTofile """") : notepad.exe  """" saveTofile """") 
            Run, edit %saveTofile%
        }
    textToSave := filename := savedClipboard := saveTofile := ""	; Restore clipboard from backup and clean temporary variables.
    gui, show
    return
}

GetActiveBrowserURL() {
	global ModernBrowsers, LegacyBrowsers
	WinGetClass, sClass, A
	If sClass In % ModernBrowsers
		Return GetBrowserURL_ACC(sClass)
	Else If sClass In % LegacyBrowsers
		Return GetBrowserURL_DDE(sClass) ; empty string if DDE not supported (or not a browser)
	Else
		Return ""
}

; "GetBrowserURL_DDE" adapted from DDE code by Sean, (AHK_L version by maraskan_user)
; Found at http://autohotkey.com/board/topic/17633-/?p=434518

GetBrowserURL_DDE(sClass) {
	WinGet, sServer, ProcessName, % "ahk_class " sClass
	StringTrimRight, sServer, sServer, 4
	iCodePage := A_IsUnicode ? 0x04B0 : 0x03EC ; 0x04B0 = CP_WINUNICODE, 0x03EC = CP_WINANSI
	DllCall("DdeInitialize", "UPtrP", idInst, "Uint", 0, "Uint", 0, "Uint", 0)
	hServer := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", sServer, "int", iCodePage)
	hTopic := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "WWW_GetWindowInfo", "int", iCodePage)
	hItem := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "0xFFFFFFFF", "int", iCodePage)
	hConv := DllCall("DdeConnect", "UPtr", idInst, "UPtr", hServer, "UPtr", hTopic, "Uint", 0)
	hData := DllCall("DdeClientTransaction", "Uint", 0, "Uint", 0, "UPtr", hConv, "UPtr", hItem, "UInt", 1, "Uint", 0x20B0, "Uint", 10000, "UPtrP", nResult) ; 0x20B0 = XTYP_REQUEST, 10000 = 10s timeout
	sData := DllCall("DdeAccessData", "Uint", hData, "Uint", 0, "Str")
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hServer)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hTopic)
	DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hItem)
	DllCall("DdeUnaccessData", "UPtr", hData)
	DllCall("DdeFreeDataHandle", "UPtr", hData)
	DllCall("DdeDisconnect", "UPtr", hConv)
	DllCall("DdeUninitialize", "UPtr", idInst)
	csvWindowInfo := StrGet(&sData, "CP0")
	StringSplit, sWindowInfo, csvWindowInfo, `" ;"; comment to avoid a syntax highlighting issue in autohotkey.com/boards
	Return sWindowInfo2
}

GetBrowserURL_ACC(sClass) {
	global nWindow, accAddressBar
	If (nWindow != WinExist("ahk_class " sClass)) ; reuses accAddressBar if it's the same window
	{
		nWindow := WinExist("ahk_class " sClass)
		accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindow))
	}
	Try sURL := accAddressBar.accValue(0)
	If (sURL == "") {
		WinGet, nWindows, List, % "ahk_class " sClass ; In case of a nested browser window as in the old CoolNovo (TO DO: check if still needed)
		If (nWindows > 1) {
			accAddressBar := GetAddressBar(Acc_ObjectFromWindow(nWindows2))
			Try sURL := accAddressBar.accValue(0)
		}
	}
	If ((sURL != "") and (SubStr(sURL, 1, 4) != "http")) ; Modern browsers omit "http://"
		sURL := "http://" sURL
	If (sURL == "")
		nWindow := -1 ; Don't remember the window if there is no URL
	Return sURL
}

; "GetAddressBar" based in code by uname
; Found at http://autohotkey.com/board/topic/103178-/?p=637687

GetAddressBar(accObj) {
	Try If ((accObj.accRole(0) == 42) and IsURL(accObj.accValue(0)))
		Return accObj
	Try If ((accObj.accRole(0) == 42) and IsURL("http://" accObj.accValue(0))) ; Modern browsers omit "http://"
		Return accObj
	For nChild, accChild in Acc_Children(accObj)
		If IsObject(accAddressBar := GetAddressBar(accChild))
			Return accAddressBar
}

IsURL(sURL) {
	Return RegExMatch(sURL, "^(?<Protocol>https?|ftp)://(?<Domain>(?:[\w-]+\.)+\w\w+)(?::(?<Port>\d+))?/?(?<Path>(?:[^:/?# ]*/?)+)(?:\?(?<Query>[^#]+)?)?(?:\#(?<Hash>.+)?)?$")
}

; The code below is part of the Acc.ahk Standard Library by Sean (updated by jethrow)
; Found at http://autohotkey.com/board/topic/77303-/?p=491516

Acc_Init()
{
	static h
	If Not h
		h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Acc_ObjectFromWindow(hWnd, idObject = 0)
{
	Acc_Init()
	If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return ComObjEnwrap(9,pacc,1)
}
Acc_Query(Acc) {
	Try Return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}
Acc_Children(Acc) {
	If ComObjType(Acc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	Else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		If DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren%
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=9?Acc_Query(child):child), NumGet(varChildren,i-8)=9?ObjRelease(child):
			Return Children.MaxIndex()?Children:
		} Else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
}