#Requires AutoHotkey v2.0

;based on:
;Explorer window interaction (folder windows/Desktop, file/folder enumeration/selection/navigation/creation) - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?f=6&t=35041

#HotIf WinActive("ahk_class CabinetWClass", )
^Down:: ;explorer - navigate to sibling folder
^Up:: ;explorer - navigate to sibling folder
#HotIf WinActive("ahk_class ExploreWClass", )
^Down:: ;explorer - navigate to sibling folder
^Up:: ;explorer - navigate to sibling folder
{ ; V1toV2: Added bracket
vGetNext := !!InStr(A_ThisHotkey, "D")
hWnd := WinGetID("A")
vWinClass := WinGetClass("ahk_id " hWnd)
if !(vWinClass = "CabinetWClass") && !(vWinClass = "ExploreWClass")
	return
for oWin2 in ComObject("Shell.Application").Windows
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
SplitPath(vDir2, &vName1, &vDir1)
vIsMatch := 0
vName := ""
Loop Files, vDir1 "\*", "D"
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
	DllCall("shell32\SHParseDisplayName", "WStr", vDir, "Ptr", 0, "PtrP", &vPIDL, "UInt", 0, "Ptr", 0)
	SAFEARRAY := Buffer(A_PtrSize=8?32:24, 0) ; V1toV2: if 'SAFEARRAY' is a UTF-16 string, use 'VarSetStrCapacity(&SAFEARRAY, A_PtrSize=8?32:24)'
	NumPut("UShort", 1, &SAFEARRAY, 0) ;cDims
	NumPut("UInt", 1, &SAFEARRAY, 4) ;cbElements
	NumPut("Ptr", vPIDL, &SAFEARRAY, A_PtrSize=8?16:12) ;pvData
	NumPut("Int", DllCall("shell32\ILGetSize", "Ptr", vPIDL, "UInt"), &SAFEARRAY, A_PtrSize=8?24:16) ;rgsabound[1]
	oWin.Navigate2(ComValue(0x2011, &SAFEARRAY), 0)
	DllCall("shell32\ILFree", "Ptr", vPIDL)
}
oWin := ""
return
} ; V1toV2: Added bracket in the end
#HotIf