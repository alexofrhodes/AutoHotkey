;====================================================================
; 
; Demo of GuiSave and GuiRestore functions
;
; Programmer: Alan Lilly 
; AutoHotkey: v1.1.03.00 (autohotkey_L ANSI version)
;
;====================================================================

VERSION := 1.0

RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe
WINTITLE := basename1 " " VERSION

#SingleInstance force  
#NoENV              ; Avoids checking empty variables to see if they are environment variables (recommended for all new scripts and increases performance).
SetBatchLines -1    ; have the script run at maximum speed and never sleep
ListLines Off       ; a debugging option

;outputdebug DBGVIEWCLEAR

;============================================================
; 1. When this ahk program is compiled into an exe, fileinstall indicates which files should be embedded inside the exe.
; 2. When the program is run, fileinstall extracts the embedded file to the specified folder.
;============================================================

RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath 

; if Not InStr(FileExist(A_AppData "\" basename1), "D")    ; create appdata folder if doesnt exist
;     FileCreateDir , % A_AppData "\" basename1
    
;============================================================
; Build gui:
;============================================================ 

Gui, Add, Text, xm section, Preset
Gui, Add, ComboBox, x+5 vfrmSAVEDPRESET gPresetChange
    
Gui, Add, Button, x+5 h21 w60 gSavePreset, Save    
Gui, Add, Button, x+5 h21 w60 gDeletePreset vDELETEBUTTON, Delete 

Gui, Add, Text, xm
Gui, Add, Text, xm section , Edit1
Gui, Add, Edit, ys w200 vfrmJOBTITLE, 

Gui, Add, Text, xs section right, DropDown
Gui, Add, DropDownList, x+5 vfrmCONVERT, First||Second|Third

Gui, Add, Checkbox, vfrmRECURSE section xm , Checkbox1
Gui, Add, Checkbox, vfrmAUTODUMP xs , Checkbox2

Gui, Add, StatusBar

GetXY(winx, winy)
Gui, Show,x%winx% y%winy%,%WINTITLE%

GoSub, UpdatePresetList  ; update drop down to show all preset section names in ini file

return

;============================================================
; do a guirestore for newly selected preset
;============================================================

PresetChange:

    gui, submit, nohide
    
    ; if drop down text is blank then error message and return
    if (frmSAVEDPRESET = "") 
        return
    
    ; save gui values after combobox1 to ini file under given section
    guirestore("presets.ini",frmSAVEDPRESET)
    
Return

;============================================================
; save preset to presets.ini
;============================================================

SavePreset:

    gui, submit, nohide
    
    ; if drop down text is blank then error message and return
    if (frmSAVEDPRESET = "") {
        SB_SetText("Preset name required")
        return
    }
    
    guisave("presets.ini", frmSAVEDPRESET, "DELETEBUTTON")
    
    GoSub, UpdatePresetList  ; update drop down to show all preset section names in ini file
    
    GuiControl, Text, frmSAVEDPRESET, % frmSAVEDPRESET  ; update the control
    
    SB_SetText(frmSAVEDPRESET " preset saved") 
    
Return

;============================================================
; delete selected preset section from presets.ini
;============================================================

DeletePreset:

    gui, submit, nohide
    
    RegExMatch(A_ScriptName, "^(.*?)\.", basename) 
    
    ; if drop down text is blank then error message and return
    if (frmSAVEDPRESET = "") {
        SB_SetText("Preset name required")
        return
    }
    
    ; delete entire section from ini file
    ; IniDelete, %A_AppData%\%basename1%\presets.ini, %frmSAVEDPRESET%
    IniDelete, %A_ScriptDir%\presets.ini, %frmSAVEDPRESET%

    SB_SetText(frmSAVEDPRESET " preset deleted" ) 
    
    GoSub, UpdatePresetList  ; update drop down to show all preset section names in ini file
    
Return

;============================================================
; update drop down to show all preset section names in ini file, except section1
;============================================================

UpdatePresetList:

    gui, submit, nohide
    
    RegExMatch(A_ScriptName, "^(.*?)\.", basename) 
    
    ; get all section names in ini file
    ; IniRead, sectionNames, %A_AppData%\%basename1%\presets.ini 
    IniRead, sectionNames, %A_ScriptDir%\presets.ini

    sectionNames := RegExReplace(sectionNames , "\n", "|")         ; change newline to pipe
    sectionNames := RegExReplace(sectionNames , "section1[\|]?", "")    ; exclude section1
    sectionNames := "|" sectionNames
    
    ; update drop down to show all preset section names in ini file
    GuiControl, , frmSAVEDPRESET, % sectionNames  ; update the control
    
Return
    
;============================================================
; when you click x or close button
;============================================================ 

GuiClose:

    Gui, Submit, NoHide      ; update control variables
    
    ; use script's basename to define ini file panel position and presets.ini
    RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe

    ; get window state
    WinGet, winstate, MinMax, %WINTITLE%
    ; do not save window position if minimized, winx and winy would be something like -32000
    if (winstate != -1) {      
        ; save window dimensions, location, and column widths!    
        WinGetPos , x, y, Width, Height, %WINTITLE%
        ; IniWrite, %x%, %A_AppData%\%basename1%\%basename1%.ini, Window Position, winx
        ; IniWrite, %y%, %A_AppData%\%basename1%\%basename1%.ini, Window Position, winy
        IniWrite, %x%, %A_ScriptDir%\%basename1%.ini, Window Position, winx
        IniWrite, %y%, %A_ScriptDir%\%basename1%.ini, Window Position, winy

    }
        
ExitApp

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; FUNCTIONS
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;============================================================
; Return window x and y position from ini file.
;============================================================

GetXY(byref winx, byref winy)   
{

    RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe

    ;============================================================
    ; position gui based on values from ini file
    ;============================================================
    
    ; IniRead, winx, %A_AppData%\%basename1%\%basename1%.ini, Window Position, winx, 0
    ; IniRead, winy, %A_AppData%\%basename1%\%basename1%.ini, Window Position, winy, 0
    IniRead, winx,  %A_ScriptDir%\%basename1%.ini, Window Position, winx, 0
    IniRead, winy,  %A_ScriptDir%\%basename1%.ini, Window Position, winy, 0

    ; get the width and height of the entire desktop (even if it spans multiple monitors)
    SysGet, VirtualWidth, 78
    SysGet, VirtualHeight, 79

    ; prevent display of gui off-screen (somehow this was still happening to jess, so I added this logic)
    if (winx < 0) OR (winx > VirtualWidth) 
        winx := 0

    if (winy < 0) OR (winy > VirtualHeight)     
        winy := 0    
    
    Return
    
}

;============================================================
; save all gui control values for active gui to ini file
;============================================================

GuiSave(inifile,section,begin="",end="")   
{
    SplitPath, inifile, file, path, ext, base, drive     ; splitpath expects paths with \
    
    if (path = "") {   ; if no path given then use default path
        RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe
        ; inifile := A_AppData "\" basename1 "\" inifile
        inifile := A_ScriptDir "\" inifile
    }
    
    WinGet, List_controls, ControlList, A    ; get list of all controls active gui

    if (begin = "")
        flag := 0
    else 
        flag := 1
        
    Loop, Parse, List_controls, `n
    { 
        ;ControlGet, cid, hWnd,, %A_LoopField%         ; get the id of current control
        GuiControlGet, textvalue,,%A_Loopfield%,Text  ; get associated text
        GuiControlGet, vname, Name, %A_Loopfield%     ; get controls vname
    
        If (vname = "")   ; only save controls which have a vname 
            continue

        if (begin = vname) {
            flag := 0
            continue
        }
            
        if (flag) 
            continue
            
        if (end = vname) 
            break
            
        GuiControlGet, value ,, %A_Loopfield%         ; get controls value
        value := RegExReplace(value, "`n", "|")       ; convert newlines to pipes (for multiline edit fields, because newlines are not valid for ini file)
        
        ; todo: truncate edit values to not exceed ini fieldsize limit (1024?)  OR blank (all or nothing)
        
        IniWrite, % value, %inifile%, %section%, %vname%
        
    }
   
   return
}

;============================================================
; Update gui controls with values from ini file.
;============================================================

GuiRestore(inifile,section)   
{

    SplitPath, inifile, file, path, ext, base, drive     ; splitpath expects paths with \
    
    if (path = "") {   ; if no path given then use default path
        RegExMatch(A_ScriptName, "^(.*?)\.", basename)    ; dont use splitpath to get basename because it cant handle DeltaRush.1.3.exe
        ; inifile := A_AppData "\" basename1 "\" inifile
        inifile := A_ScriptDir "\" inifile
    }

    ;============================================================
    ; update gui controls with values from ini file
    ;============================================================    

    WinGet, List_controls, ControlList, A   ; get list of all controls for active gui
    
    Loop, Parse, List_controls, `n
    { 
    
        ;ControlGet, cid, hWnd,, %A_LoopField%         ; get the id of current control
        ;GuiControlGet, textvalue,,%A_Loopfield%,Text  ; get controls associated text
        GuiControlGet, vname, Name, %A_Loopfield%     ; get controls vname
        GuiControlGet, value ,, %A_Loopfield%         ; get controls value
        
        If (vname = "")   ; only process controls which have a vname 
            continue
        
        IniRead, value, %inifile%, %section%, %vname%, ERROR
        
        if (value != "ERROR") {
            
            value := RegExReplace(value, "\|", "`n")       ; convert pipes to newlines (for multiline edit fields, because newlines are not valid for ini file)
            
            RegExMatch( A_Loopfield, "(.*?)\d+", name)   ; extract the control name without numbers
            if (name1 = "ComboBox") {
                GuiControl, ChooseString, %A_Loopfield%, %value%   ; select item in dropdownlist
            } else {
                GuiControl,  ,%A_Loopfield%, %value%    ; update the control
            }
        }
        
    }
    
    return
   
}