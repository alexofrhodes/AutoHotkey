#SingleInstance, Force
SetKeyDelay, 50

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"

; Custom Tray Icon
I_Icon = ScriptSaver.ico
IfExist, %I_Icon%
    Menu, Tray, Icon, %I_Icon%

; Event for Tray icon left-click
OnMessage(0x404, "AHK_NOTIFYICON")

global MyPaths := ""
global MyExtensions := ""

global SettingsFile := "config.ini"
global TargetPath
global TargetExtension
global ChAskFileName
global ChEditAfterSave
global EditFileNameValue 

AHK_NOTIFYICON(wParam, lParam)
{
    if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
        CreateGui()
        return 0
    }
}

createGui()

return

CreateGui(){
    global 
    Gui, destroy
    LoadPaths()
    LoadExtensions()

    gui, add, Text,x12 section,File Name 
    IniRead, EditFileNameValue, %SettingsFile%, Settings, FileName
    Gui, Add, Edit, xs vEditFileName gSaveSettings -wrap -wanttab -WantReturn w200 section, %EditFileNameValue%

    IniRead, ChAskFileName, %SettingsFile%, Settings, AskFileName
    Gui, Add, CheckBox, gSaveSettings vChAskFileName Checked%ChAskFileName%, Ask for file name instead
  
    Gui, Add, Button, ys w100 section gSaveScript, Save Script
    IniRead, ChEditAfterSave, %SettingsFile%, Settings, EditAfterSave    
    Gui, Add, CheckBox, gSaveSettings vChEditAfterSave Checked%ChEditAfterSave%, Edit after save

    Gui, Add, Text, section x12 w300 0x10  ;Horizontal Line > Etched Gray
    ; Gui, Add, Text, x5 y5 h150 0x11  ;Vertical Line > Etched Gray
    ; Gui, Add, Text, x5 y155 w150 h1 0x7  ;Horizontal Line > Black
    ; Gui, Add, Text, x155 y5 w1 h150 0x7  ;Vertical Line > Black

    gui, add, text, x12 ys+10  section, Target Folder
    Gui, Add, Button, ys-3 gOpenFolder, Open
    gui, add, button, ys-3 gAddNewFolder, New
    Gui, Add, ListBox, x12 w150 hwndhTargetPath vTargetPath gSaveSettings, %MyPaths%
    ItemHeight := LB_GetItemHeight(hTargetPath)
    hIndex := StrSplit(MyPaths, "|").length() * ItemHeight
    GuiControl, move, TargetPath , h%hIndex%
   
    iniRead, selItem, %SettingsFile%, Settings, Path   
    ControlGet, Items, List, , TargetPath
    Loop, Parse, MyPaths, |
        if A_LoopField = %selItem%
            {
                GuiControl, Choose, TargetPath, %A_Index% 
                break
            }

    gui, add, text, ys section  , File Extension
    Gui, Add, Button, ys-3  w18 h18 gAddExtension, +
    Gui, Add, Button, ys-3 w18 h18 gRemoveExtension, -

    Gui, Add, ListBox, w150 hwndhTargetExtension vTargetExtension xs section gSaveSettings, %MyExtensions%
    ItemHeight := LB_GetItemHeight(hTargetExtension)
    hIndex := StrSplit(MyExtensions, "|").length() * ItemHeight
    GuiControl, move, TargetExtension , h%hIndex%

    iniRead, selItem, %SettingsFile%, Settings, Extension   
    ControlGet, Items, List, , TargetExtension
    Loop, Parse, MyExtensions, |
        if A_LoopField = %selItem%
            {
                GuiControl, Choose, TargetExtension, %A_Index% 
                break
            }


    Gui, +AlwaysOnTop -MinimizeBox -MaximizeBox
    Gui, Show, autosize, Script Saver
    return
}

LoadExtensions(){
    IniRead, MyExtensions, config.ini, Settings, MyExtensions 
}

AddNewFolder(){
    global
    InputBox, NewFolder , Creating a Folder, Choose New Folder's Name,,,130
    if StrLen(NewFolder) =0
        return
    FileCreateDir, saved scripts\%NewFolder%
    MyPaths .= "|" . NewFolder
    GuiControl, , TargetPath, %NewFolder%
    ItemHeight := LB_GetItemHeight(hTargetPath)
    hIndex := StrSplit(MyPaths, "|").length() * ItemHeight
    GuiControl, move, TargetPath , h%hIndex%
    gui, show, AutoSize
}

LoadPaths(){
    global
    MyPaths .= ""
    Loop Files, %A_ScriptDir%\saved scripts\*, D  
        {
            SplitPath, A_LoopFileFullPath, FileName, Folder, Extension, Filename_no_ext	
            MyPaths .= "|" . FileName
        }    
        MyPaths := SubStr(mypaths,2)
        return
}

GuiEscape:
GuiClose:
{
    SaveSettings()
    Gui, Hide
    return
}


OpenFolder:
    Gui, Submit, NoHide
    Run, %A_ScriptDir%\saved scripts\%TargetPath%\
    return

SaveSettings()
{
    Gui, Submit, NoHide
    IniWrite, %TargetPath%, %SettingsFile%, Settings, Path
    IniWrite, %TargetExtension%, %SettingsFile%, Settings, Extension
    IniWrite, %ChEditAfterSave%, %SettingsFile%, Settings, EditAfterSave
    IniWrite, %ChAskFileName%, %SettingsFile%, Settings, AskFileName
    IniWrite, %MyExtensions%, %SettingsFile%, Settings, MyExtensions
    
    GuiControlGet, EditFileNameValue, , EditFileName
    iniWrite, %EditFileNameValue%, %SettingsFile%, Settings, FileName
    return
}

AddExtension()
{
    global
    Gui, Submit, NoHide
    InputBox, Extension, add new Extension,,,,130
    if StrLen(Extension) = 0
        return
    MyExtensions .= "|" . Extension
    GuiControl, , TargetExtension, %Extension%
    b_index := StrSplit(MyExtensions, "|").length()
    ; Control, Choose, 3, hTargetExtension    ;;not working???

    ItemHeight := LB_GetItemHeight(hTargetExtension)
    hIndex := StrSplit(MyExtensions, "|").length() * ItemHeight
    GuiControl, move, TargetExtension , h%hIndex%
    
    SaveSettings()
    gui,show, AutoSize
    return
}

RemoveExtension()
{
    global
    Gui, Submit, NoHide
    Temp := ""
    Loop, Parse, MyExtensions, |
    {
        if !A_LoopField || (A_LoopField = TargetExtension)
            continue
        Temp .= "|" . A_LoopField
    }

    StringMid, MyExtensions, Temp, 2
    GuiControl,,TargetExtension,|
    GuiControl, , TargetExtension, %MyExtensions%

    ItemHeight := LB_GetItemHeight(hTargetExtension)
    hIndex := StrSplit(MyExtensions, "|").length() * ItemHeight
    GuiControl, move, TargetExtension , h%hIndex%
        
    SaveSettings()
    gui,show, AutoSize
    return
}

SaveScript()
{
    Gui, Hide
    Gui, Submit, NoHide

    savedClipboard := ClipboardAll
    Clipboard := ""
    While !Clipboard
    {
        Send, ^c
        Sleep, 1
    }

    textToSave := Clipboard
    Clipboard := savedClipboard

    IfWinActive, %ModernBrowsers% ahk_exe chrome.exe ; Add your browser's executable name
    {        
        sURL := GetActiveBrowserURL()
    }else{
        WinGetTitle, sURL, A
    }

    Extension := TargetExtension
    saveToFolder := TargetPath

    if (saveToFolder = "")
        FileSelectFolder, saveToFolder, %A_ScriptDir%
    Else
        saveToFolder := A_ScriptDir "\saved scripts\" saveToFolder

    ; FileCreateDir, %saveToFolder%
    GuiControlGet, EditFileNameValue, , EditFileName
    if StrLen(EditFileNameValue) > 0
    {
        filename := EditFileNameValue
    }else if (ChAskFileName = 1) or (%EditFileNameValue% = ""){
        InputBox, filename, Enter the name for the script's file,,,,130
        if ErrorLevel
            filename := ""
    }
    
    
    saveToFile := (saveToFolder ? saveToFolder : A_ScriptDir) "\" (filename ? filename : A_Now) "." TargetExtension

    FileAppend, %sURL% `n %textToSave%, %saveToFile%, UTF-8

    if (ChEditAfterSave = 1)
    {
        Run, edit %saveToFile%
    }
    textToSave := filename := savedClipboard := saveToFile := ""
    Gui, Show
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


; ----------------------------------------------------------------------------------------
GetClientHeight(HWND) { ; Retrieves the height of the client area.
    VarSetCapacity(RECT, 16, 0)
    DllCall("User32.dll\GetClientRect", "Ptr", HWND, "Ptr", &RECT)
    Return NumGet(RECT, 12, "Int")
 }
 ; ----------------------------------------------------------------------------------------
 LB_GetItemHeight(HLB) { ; Retrieves the height of a single list box item.
    ; LB_GETITEMHEIGHT = 0x01A1
    SendMessage, 0x01A1, 0, 0, , ahk_id %HLB%
    Return ErrorLevel
 }
 ; ----------------------------------------------------------------------------------------
 LB_GetCount(HLB) { ; Retrieves the number of items in the list box.
    ; LB_GETCOUNT = 0x018B
    SendMessage, 0x018B, 0, 0, , ahk_id %HLB%
    Return ErrorLevel
 } 