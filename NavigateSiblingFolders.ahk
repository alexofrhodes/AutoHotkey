;based on:
;Explorer window interaction (folder windows/Desktop, file/folder enumeration/selection/navigation/creation) - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?f=6&t=35041

#IfWinActive, ahk_class CabinetWClass
^PgUp:: ;explorer - navigate to sibling folder
^PgDn:: ;explorer - navigate to sibling folder
#IfWinActive, ahk_class ExploreWClass
^PgUp:: ;explorer - navigate to sibling folder
^PgDn:: ;explorer - navigate to sibling folder
vGetNext := !!InStr(A_ThisHotkey, "Dn")
WinGet, hWnd, ID, A
WinGetClass, vWinClass, % "ahk_id " hWnd
if !(vWinClass = "CabinetWClass") && !(vWinClass = "ExploreWClass")
	return
for oWin2 in ComObjCreate("Shell.Application").Windows
	if (oWin2.HWND = hWnd)
	{
		vDir2 := oWin2.Document.Folder.Self.Path
		vDir2 := RTrim(vDir2, "\")
		oWin := oWin2
		break
	}
oWin2 := ""
if !FileExist(vDir2)
{
	oWin := ""
	return
}
SplitPath, vDir2, vName1, vDir1
vIsMatch := 0
vName := ""
Loop, Files, % vDir1 "\*", D
{
	if vIsMatch
	{
		vName := A_LoopFileName
		break
	}
	if (A_LoopFileName = vName1)
		if vGetNext
			vIsMatch := 1
		else
			break
	vName := A_LoopFileName
}
if (vName = vName1) || (vName = "")
{
	oWin := ""
	return
}
vDir := vDir1 "\" vName
;MsgBox, % vDir

if !InStr(vDir, "#") ;folders that don't contain #
	oWin.Navigate(vDir)
else ;folders that contain #
{
	DllCall("shell32\SHParseDisplayName", WStr,vDir, Ptr,0, PtrP,vPIDL, UInt,0, Ptr,0)
	VarSetCapacity(SAFEARRAY, A_PtrSize=8?32:24, 0)
	NumPut(1, &SAFEARRAY, 0, "UShort") ;cDims
	NumPut(1, &SAFEARRAY, 4, "UInt") ;cbElements
	NumPut(vPIDL, &SAFEARRAY, A_PtrSize=8?16:12, "Ptr") ;pvData
	NumPut(DllCall("shell32\ILGetSize", Ptr,vPIDL, UInt), &SAFEARRAY, A_PtrSize=8?24:16, "Int") ;rgsabound[1]
	oWin.Navigate2(ComObject(0x2011, &SAFEARRAY), 0)
	DllCall("shell32\ILFree", Ptr,vPIDL)
}
oWin := ""
return
#IfWinActive