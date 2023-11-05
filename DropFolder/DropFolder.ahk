﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Skrommel
;DropFolder.ahk
; Drop files on a floating icon to move or copy files to user defined folders
; Press Ctrl while dropping to copy files

; Anastasiou Alex
;----------------
; load lists from txt files
; right click menu to 
;   switch lists
;   Add Current Folder or Selected Folder(s) To List

#SingleInstance,Force
SetWindelay,0
global loadedList :=
global selectedFiles := []

I_Icon := A_ScriptDir "\img\FileCabinet.jpg" 
iF FileExist(I_Icon)
  Menu, Tray, Icon , %I_Icon%

OnMessage(0x404, "AHK_NOTIFYICON")
AHK_NOTIFYICON(wParam, lParam){
  if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
      WinGet, win_minimized, MinMax, DropFolder
      if (win_minimized != 1)
        Gui, Hide
      else
        gosub, start
    }
}

Start:
{
  gui, destroy
  gosub, LoadSettings
  transparency= 500
  x=1500
  y=100
  w=100
  h=
  location=AlwaysOnTop
  Gosub,MENU
  OnMessage(0x112,"WM_SYSCOMMAND")
  If location=AlwaysOnTop
    guioptions:="+" location
  If w<>
    pictureoptions:="H-1 W" w
  Else
    pictureoptions:="W-1 H" h
  Gui,Margin,0,0
  Gui,% "-Caption +ToolWindow +LastFound -Resize " guioptions
  guiid:=WinExist()
  Gui,Add,Picture,% pictureoptions " GMOVEWINDOW",% picture

  Gui, Color, Black
  gui, font, s10 w600 cYellow, Consolas
  gui,add,Text, +BackgroundBlack vListCaption h16, AAAAAAAAAAA
  gosub, assignCaption
  Gui,Show,% "X" x " Y" y, DropFolder
  WinSet,Transparent,% transparency
}
Return

AssignCaption:
{
  cap := SubStr(loadedList, instr(loadedList,"\")+1)
  cap := substr(cap, 1, InStr(cap, ".")-1)
  GuiControl, Text, ListCaption , % " " cap
}
return

LoadSettings:
{
  picture=img\FileCabinet.jpg
  IniRead, loadedList, DropFolder.ini, Settings, loadedList, DropLists.txt
  loadedList=LISTS\%loadedList%
  selectedFiles := []
  FileRead, fileContents, %loadedList%
  Loop, Parse, fileContents, `n
  {
      line := Trim(A_LoopField)
      if line
          selectedFiles.Push(line)
  }
}
return

GetExplorerHwnd(hwnd=0){
	if(hwnd==0){
		explorerHwnd := WinActive("ahk_class CabinetWClass")
		if(explorerHwnd==0)
			explorerHwnd := WinExist("ahk_class CabinetWClass")
	}
	else
		explorerHwnd := WinExist("ahk_class CabinetWClass ahk_id " . hwnd)
	
	if (explorerHwnd){
		for window in ComObjCreate("Shell.Application").Windows{
			try{
				if (window && window.hwnd && window.hwnd==explorerHwnd)
					return window.hwnd
        ;return window.Document.Folder.Self.Path
			}
		}
	}
	return false
}
AddCurrentOrSelectedFoldersToList:
{
  folderPath := Explorer_GetSelection()
  FileAppend, `n%folderPath%, %loadedList%
  GoSub, LoadSettings
  GoSub, Menu
}

Explorer_GetSelection() {
  hwnd:=GetExplorerHwnd()
  for window in ComObjCreate("Shell.Application").Windows
    if (hWnd = window.HWND) && (oShellFolderView := window.document)
        break
  for item in oShellFolderView.SelectedItems
    {
      if instr(FileExist(item.path),"D")
        result .= (result = "" ? "" : "`n") . item.path
    }
  if !result
      result := oShellFolderView.Folder.Self.Path
  Return result
}

Nothing:
Return

GuiContextMenu:
Switch A_GuiEvent {
	Case "Normal":
    
	Case "RightClick":
    Try 
      Menu, MyMenu, DeleteAll
    Menu, MyMenu, Add, Hide, GuiClose
    Menu, MyMenu, Add, Open folder, OpenFolder       
    Menu, MyMenu, Add, Edit Current List, EditList   
    Menu, MyMenu, Add, Add Current or Selected Folder(s) to List, AddCurrentOrSelectedFoldersToList 
    Menu, MyMenu, Add
    Menu, MyMenu, add, SWITCH LIST, nothing
    menu, mymenu, disable, SWITCH LIST
    Menu, MyMenu, Add    

    Loop Files, %A_ScriptDir%\LISTS\*.txt, R  ; Recurse into subfolders.
      {
        SplitPath, A_LoopFileFullPath, FileName, Folder, Extension, Filename_no_ext
        Menu, MyMenu, Add, %Filename_no_ext%, SwitchList
      }    
    CoordMode, Mouse, Screen
    Menu, MyMenu, Show
}
Return

SwitchList:
  IniWrite, %A_ThisMenuItem%.txt, DropFolder.ini, Settings, loadedList
  GoSub, LoadSettings
  GoSub, Menu
  gosub, assignCaption
return

GuiDropFiles:
{  
  GetKeyState,ctrl,Control,P
  where:=A_GuiControl
  filecount:=A_EventInfo
  guix:=A_GuiX
  guiy:=A_GuiY
  files:=A_GuiEvent
  Menu,menu,Show 
  Return
}

MENU:
{
  Try 
    Menu, Menu, DeleteAll
  Menu, menu, Add, &Browse..., BROWSE
  Menu, menu, Add
  for index, filePath in selectedFiles
    Menu, menu, Add, % filePath, COPY
  Return
}

EditList:
  run, %loadedList% 
Return

OpenFolder:
  run, LISTS
return

COPY:
{
  DestinationFolder := A_ThisMenuItem  
  If ctrl = D 
  {
    Loop, Parse, files, `n
    {
      ItemPath := A_LoopField
      if isFile(ItemPath) {
          FileCopy, %ItemPath%, %DestinationFolder%\%OutFileName%
      } else if isDir(ItemPath) {
          FileCopyDir, %ItemPath%, %DestinationFolder%\%OutFileName%
      }
    }
  } else {
    Loop, Parse, files, `n
    {
      ItemPath := A_LoopField
      SplitPath, A_LoopField, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
      
      if isFile(ItemPath) {
          FileMove, %ItemPath%, %DestinationFolder%\%OutFileName%
      } else if isDir(ItemPath) {
          FileMoveDir, %ItemPath%, %DestinationFolder%\%OutFileName%
      }
    }
  }
  Return
}

isFile(Path){
  Return !InStr(FileExist(Path), "D") 
}

isDir(Path){
  Return !!InStr(FileExist(Path), "D") 
}

BROWSE:
{
  FileSelectFolder,target,,3
  If target=
    Return
  FileCreateDir,%target%
  If ctrl=D
    Loop,Parse,files,`n
      FileCopy,%A_LoopField%,%target%
  Else
    Loop,Parse,files,`n
      FileMove,%A_LoopField%,%target%
  Return
}

;Stolen from SKAN at http://www.autohotkey.com/forum/topic32768.html
MOVEWINDOW: 
  PostMessage,0xA1,2,,,A   ;WM_NCLBUTTONDOWN=0x00A1 HTCAPTION=2
Return

;Stolen from Lexicos at http://www.autohotkey.com/forum/topic18260.html
WM_SYSCOMMAND(wParam) { 
  Global guiid
  If (A_Gui && wParam = 0xF020) ; SC_MINIMIZE 
    WinRestore,ahk_id %guiid% 
} 

GuiClose: 
  Gui, Hide 
return

