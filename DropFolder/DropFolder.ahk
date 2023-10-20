#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Skrommel
;DropFolder.ahk
; Drop files on a floating icon to move or copy files to user defined folders
; Press Ctrl while dropping to copy files


;Anastasiou Alex
; added option to load lists from txt files
; added right click menu to switch lists

#SingleInstance,Force
SetWindelay,0

Start:
{
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
  Gui,Show,% "X" x " Y" y
  WinSet,Transparent,% transparency
}
Return

LoadSettings:
{
  picture=Settings\FileCabinet.jpg
  IniRead, loadedList, Settings\DropFolder.ini, Settings, loadedList, DropLists.txt
  loadedList=Settings\%loadedList%
  ; Define the array to store the selected files
  selectedFiles := []

  ; Read the list of file paths from a text file
  FileRead, fileContents, %loadedList%
  ; Split the file paths by line and add them to the selectedFiles array
  Loop, Parse, fileContents, `n
  {
      ; Trim leading and trailing spaces from the line
      line := Trim(A_LoopField)
      ; Check if the line is not empty
      if line
      {
          ; Add the file path to the selectedFiles array
          selectedFiles.Push(line)
      }
  }
}
return

GuiContextMenu:
Switch A_GuiEvent {
	Case "Normal":
    
	Case "RightClick":
    Try 
      Menu, MyMenu, DeleteAll
    Menu, MyMenu, Add, GuiClose
    Menu, MyMenu, Add, EditLists    
        
    Menu, MyMenu, Add

    Loop Files, %A_ScriptDir%\Settings\*.txt, R  ; Recurse into subfolders.
      {
        SplitPath, A_LoopFileFullPath, FileName, Folder, Extension, Filename_no_ext
        Menu, MyMenu, Add, %Filename_no_ext%, SwitchList
      }    

    ; Menu, MyMenu, Add, List1
    ; Menu, MyMenu, Add, Droplists    

    ; Show the context menu at the cursor's position
    CoordMode, Mouse, Screen
    Menu, MyMenu, Show
}
Return

SwitchList:
  IniWrite, %A_ThisMenuItem%.txt, Settings\DropFolder.ini, Settings, loadedList
  GoSub, LoadSettings
  GoSub, Menu
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

  for index, filePath in selectedFiles
  {
    Menu, menu, Add, % filePath, COPY
  }
  Menu, menu, Add
  Menu, menu, Add, &Browse..., BROWSE
  Menu, menu, Add
  Menu, menu, Add, PickList
  ; Menu, menu, Add, &Cancel, CANCEL
  Return
}

PickList:
  try
    Menu, MyMenu, Show
return

EditLists:
  run, Settings
Return

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

isFile(Path)
{
  Return !InStr(FileExist(Path), "D") 
}

isDir(Path)
{
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

CANCEL:
Return

;Stolen from SKAN at http://www.autohotkey.com/forum/topic32768.html
MOVEWINDOW: 
  PostMessage,0xA1,2,,,A   ;WM_NCLBUTTONDOWN=0x00A1 HTCAPTION=2
Return

;Stolen from Lexicos at http://www.autohotkey.com/forum/topic18260.html
WM_SYSCOMMAND(wParam) 
{ 
  Global guiid
  If (A_Gui && wParam = 0xF020) ; SC_MINIMIZE 
    WinRestore,ahk_id %guiid% 
} 

GuiClose: 
  ExitApp 
return

