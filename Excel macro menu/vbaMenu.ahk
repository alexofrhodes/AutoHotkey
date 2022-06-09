#noEnv
#singleinstance, force
sendMode input
setWorkingDir, % a_scriptDir

;#IfWinActive ahk_exe EXCEL.EXE ;sometimes error
;#IfWinActive ahk_class XLMAIN ;sometimes error
#IfWinActive ahk_class wndclass_desked_gsk ;if vbeditor window is active

/* NOTES
	----------
	replace WorkbookName:="ProjectStarter.xlam!" with your workbook where the macros are
	edit/create the vba.menu file in same folder as this file to contain your macros
	ezMenu will create a menu from the vba.menu file
*/

WorkbookName:="ProjectStarter.xlam!" 

^h::
; If Excel_Get fails it returns an error message instead of an object.
XL := Excel_Get() 
if !IsObject(XL)  {
	MsgBox, 16, Excel_Get Error, % XL
	return
}else{
;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes
}

ezMenu("vbaMenu", "vba.menu")
Reload ;otherwise throws error (menu item has no parent)
return

RunExcelMacro:
item := a_thisMenuItem
;split macro name if menu item contains accelerator
if instr(item, A_Space){
	StringSplit, Procedure, item, %A_Space%
	macro:= WorkbookName . Procedure2 
}else{
	macro:= WorkbookName . item
}
;check if macro exists in workbook
try {
	XL.Run(macro)  
} catch {
	MsgBox, 16,, Can't find %item% in %WorkbookName%
}
return




; Excel_Get by jethrow (modified)
; Forum:    https://autohotkey.com/boards/viewtopic.php?f=6&t=31840
; Github:   https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Excel/Excel_Get.ahk

Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) {
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

; References
;   https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
;   https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
;   https://autohotkey.com/boards/viewtopic.php?p=134048#p134048






/*
[script info]
version     = 1.2.1
description = easy menu creation using indentation and markdown-like syntax
ahk version = 1.1.26.01
author      = davebrny
source      = https://github.com/davebrny/ezMenu
*/

ezMenu_init:
global menu_created, default_label, default_item, disabled_item, tab_width, s_tab
default_label := ""
tab_width = 4
loop, % tab_width
    s_tab .= a_space
return


ezMenu(menu_name, string_or_file, modify_func="") {
    goSub, ezMenu_init

    if fileExist(string_or_file)
         menu_text := ezMenu_get(string_or_file)
    else menu_text := trim(string_or_file, "`n")
    if (menu_text = "")
        error_return("""" menu_name """ menu is empty")

    if !inStr(menu_text, "`n") and inStr(menu_text, "|") 
        stringReplace, menu_text, menu_text, |, `n, all

    if isFunc(modify_func)
        menu_text := %modify_func%(menu_text)

    loop, parse, menu_text, `n, `r
        {
        if a_loopfield is space
            continue    ; if whitespace or empty
        if (inStr(LTrim(a_loopfield), ";") = 1)
            continue    ; if commented
        line_text := a_loopfield
        menu_level(line_text, level)    ; (byRef: line_text, level)
        error_check(line_text, level, a_index)
        menu_add(menu_name, line_text, level)
        }
    menu_created := true

    menu, % menu_name, show
    menu, % menu_name, delete
}



ezMenu_get(filepath) {
    fileRead, contents, % filepath
    stringReplace, contents, contents, `r`n, `n, all
    if (subStr(filepath, -8, 9) = ".menu.ahk")
        {
        if !inStr(contents, "[ezMenu]")
            error_return("[ezMenu] not found")
        else if !inStr(contents, "[ezMenu_end]")
            error_return("[ezMenu_end] not found")
        stringGetPos, pos, contents, [ezMenu], L1
        stringMid, right_text, contents, pos + 9
        stringGetPos, pos, right_text, [ezMenu_end], L1
        stringMid, menu_text, right_text, pos, , L
        }
    else menu_text := contents
    return trim(menu_text, "`n")
}



error_return(msg) {
    msgBox, % msg
    exit
}



menu_level(byRef line_text, byRef level) {
    line_text   := rTrim(line_text)
    line_text   := strReplace(line_text, a_tab, s_tab)   ; replace tabs with spaces
    replaced    := line_text
    line_text   := LTrim(line_text)
    white_space := strLen(replaced) - strLen(line_text)  ; (whitespace count before text)
    level       := (white_space // tab_width) + 1        ; use floor divide to move incorrectly indented item up or down 1 level
}



error_check(line_text, level, line_number) {
    static last_level

    if (line_number = 1)
        last_level := "0"

    if (level > 21)
        error_return("Error on menu line " line_number " `n" line_text " `n`n"
            . "Maximum levels allowed: 21")

    if (level > last_level) and (level > last_level + 1)
        error_return("Error on menu line " line_number " `n" line_text " `n`n"
            . "Item has no parent menu")

    last_level := level
}



menu_add(menu_name, item_name, menu_level) {
    if (menu_level = "")
        menu_level := "1"
    if (default_label = "")
        default_label := menu_name
    item_label := default_label

        ;# default menu item
    if (inStr(LTrim(item_name), "*") = 1)
        {
        stringTrimLeft, item_name, item_name, 1
        default_item := true
        }
        ;# disable menu item
    if (inStr(LTrim(LTrim(item_name, "*")), ".") = 1)
        {
        stringTrimLeft, item_name, item_name, 1
        disabled_item := true
        }

        ;# custom label action
    item_name := LTrim(item_name)
    if instr(item_name, "!")
        {
        loop, parse, % item_name, !, % a_space
            {
            if (a_index = 1)
                continue
            first_word := strSplit(a_loopField, a_space)
            if isLabel(first_word[1])
                found_label := first_word[1]
            }
        until (found_label)
        if (found_label)
            {
            item_label := found_label
            if inStr(item_name, "!" found_label "!")     ; if setting a new global default
                {
                default_label := found_label
                if (inStr(item_name, "!" found_label "!") = 1)   ; if at start, then skip menu add
                    return
                item_name := strReplace(item_name, "!" found_label "!", "")
                }
            else if inStr(item_name, "!" found_label)    ; if setting a custom label for this line
                item_name := strReplace(item_name, "!" found_label, "")
            }
        item_name := trim(item_name)
        }

        ;# save value for later use
    save_value(item_name, menu_level)

        ;# add item to menu
    if (menu_level = 1)
        {
        if (trim(item_name) = "---")
            menu(menu_name, separator)
        last_menu := item_label
        }
    else if (menu_level > 1)
        {
        loop,    ; loop down through levels
            {
            sub_item    := stored_value("menu" menu_level)
            parent_item := stored_value("menu" menu_level - 1)
            item_action := (a_index = 1) ? (item_label) : (last_menu)
            if (trim(item_name) = "---") and (menu_created != true)
                menu(parent_item, separator)
            else if (trim(item_name) != "---")
                menu(parent_item, sub_item, item_action)
            last_menu := ":" parent_item
            --menu_level
            }
        until (menu_level = 1)
        }

        ;# add final root level (level 1)
    final_menu := stored_value("menu" menu_level)
    if (trim(item_name) != "---")
        menu(menu_name, final_menu, last_menu)

    disabled_item := ""
    item_label    := ""
}   ; end of menu_add()



save_value(item_name, menu_level) {
    global
    menu%menu_level% := item_name
}


stored_value(name) {
    stored_value := %name%
    return stored_value
}


menu(menu, item="", action="") {
	menu, % menu, add, % item, % action
	if (disabled_item = true)
		menu, % menu, disable, % item
	if (default_item = true)
		menu, % menu, default, % item
	disabled_item := ""
	default_item  := ""
}