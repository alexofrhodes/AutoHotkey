#Requires AutoHotkey v2.0
#SingleInstance Force

global previousClipboard := ""


myGUI := Gui()
mygui.Opt("+AlwaysOnTop -LastFound") ; +E0x08000000 = don't get focus?

btnFind := mygui.Add("Button","w80","Find")
btnFind.Enabled:=false
txtFind := mygui.Add("ComboBox", "w300 x+m")

btnReplace := mygui.Add("Button","xm w80","Replace")
btnReplace.OnEvent("Click",ClipboardReplace)
txtReplace := mygui.Add("ComboBox","x+m w300")

chbRegex := mygui.Add("Checkbox","xm w80","regex")
chbRegex.OnEvent("Click", ToggleRegexOptions)

chbCaseSense := myGUI.Add("Checkbox","x+m","CaseSense")

lblLimit := mygui.Add("Text","x+m section","Limit")
txtLimit := mygui.Add("Edit","x+m w20 h16" ,"")

lblStartingPos := mygui.Add("Text","x+m","Start")
txtStartingPos := mygui.Add("Edit","x+m w20 h16" ,"")

lblOccurrence := mygui.Add("Text","x+m","Occurrence")
txtOccurrence := mygui.Add("Edit","x+m w20 h16" ,"")

myGUI.Add("text","xm h1 w400 0x10") 

lblPlaceholder1 := mygui.Add("Text","xm w70 section","Placeholder1")

lbxCommands := myGui.Add("ListBox","xm w70 r30", [
                                        "Undo",
                                        "Clear",
                                        "Copy",
                                        "Append",
                                        "Inject",
                                        "Transclude"
                                    ])

lbxCommands.OnEvent("DoubleClick",HandleCommands)

optDirect := MyGui.Add("Radio", "ys section checked ", "Clipboard")
optDirect := MyGui.Add("Radio", "x+m", "Sidebar")

; btnUndo := mygui.Add("Button","xm w80 section","Undo")
; btnUndo.OnEvent("Click", ClipboardUndo)

; btnClear := mygui.Add("Button","x+m w80","Clear")
; btnClear.OnEvent("Click",GuiClear)

; btnCopy := mygui.Add("Button","x+m w80","Copy")
; btnCopy.OnEvent("Click",GuiCopy)

; btnAppend := mygui.Add("Button","x+m w80","Append")
; btnAppend.OnEvent("Click",GuiAppend)

; btnGuiPaste := mygui.Add("Button","x+m w80","Inject")
; btnGuiPaste.OnEvent("Click",GuiPaste)

; btnReadTextFileContent := mygui.Add("Button","xm w80","Read txt(s)")
; btnReadTextFileContent.OnEvent("Click",ClipboardReadTextFileContent)

; call function to loadOptions

ToggleRegexOptions

txtClipboard := mygui.Add("Edit","xs w500 h500", A_Clipboard)

HandleCommands(*){
    Switch lbxCommands.text, false
    {
    Case "Undo"       : ClipboardUndo
    Case "Clear"      : GuiClear
    Case "Copy"       : GuiCopy
    Case "Append"     : GuiAppend
    Case "Inject"     : GuiPaste
    Case "Transclude" : ClipboardReadTextFileContent
    Default: ;nothing
    }
}
 
^+!c::showForm

showForm(){
    global
    mygui.Show("NoActivate")
}

ToggleRegexOptions(*){
    chbCaseSense.Visible := !chbRegex.Value

    lblOccurrence.Visible := chbRegex.Value
    txtOccurrence.Visible  := chbRegex.Value
    lblStartingPos.Visible  := chbRegex.Value
    txtStartingPos.Visible  := chbRegex.Value
}

$^C::
{
    SimCopy
}

GuiCopy(*){
    myGUI.Hide
    SimCopy
    mygui.show
}

SimCopy(*){
    global 
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    ClipWait
    try 
        txtClipboard.Value := A_Clipboard
}

GuiPaste(*){
    global
    previousClipboard := A_Clipboard
    myGUI.Hide
    A_Clipboard := ""
    Send("^c")
    ClipWait    
    EditPaste A_Clipboard, txtClipboard
    A_Clipboard := txtClipboard.Text
    mygui.show    
}

GuiClear(*){
    global
    previousClipboard := A_Clipboard
    A_Clipboard := ""
    txtClipboard.text := ""
}
!C::ClipboardAppend

^!z::ClipboardUndo

ClipboardUndo(*){
    global 
    tmp := A_Clipboard
    A_Clipboard := previousClipboard
    previousClipboard := tmp
    txtClipboard.Value := A_Clipboard
}

GuiAppend(*){
    myGUI.Hide
    ClipboardAppend
    mygui.show
}

ClipboardAppend(*){ 
    global 
    previousClipboard := A_Clipboard

    This := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    ClipWait
    if this = ""{
        
    }else{
        A_Clipboard := This "`r`n" A_Clipboard
    }
    ; MsgBox(A_Clipboard)
    try 
        txtClipboard.Value := A_Clipboard
}

; puts the content of all txt-file-matches to clipboard
#C:: ClipboardReadTextFileContent

ClipboardReadTextFileContent(*){ 
    global 
    previousClipboard := A_Clipboard

    This := A_Clipboard
    A_Clipboard := ""
    FileContent := ""
    
    ; Use a regular expression to extract file paths
    FileRegex := "(?i)c:\\([a-z0-9\s_\\.():-])+?\.txt"

    ; Use a loop to find all matches in the clipboard content
    pos := 1
    while (RegExMatch(This, FileRegex, &Match, pos))
    {
        FileName := Match[0]
        FileContent := Fileread(FileName)
        A_Clipboard := A_Clipboard ? A_Clipboard "`r`n" FileContent : FileContent
        pos := Match.Pos + Match.Len
    }
    ; MsgBox(A_Clipboard)
    try 
        txtClipboard.Value := A_Clipboard
} 

ClipboardReplace(*){
    global
    previousClipboard := A_Clipboard
    if txtStartingPos.text = "" 
        txtStartingPos.text := 1
    if txtLimit.text = ""
        txtLimit.value := -1
    limit := txtLimit.Value
    if limit < 1
        limit := -1
    if txtOccurrence.text = "" 
        txtOccurrence.text := 1
    if chbRegex.value = true {
        RegExReplace(A_Clipboard, txtFind.text, txtReplace.text, , Limit, StartingPos)
    }else{
        if IsNumber(txtStartingPos.text){
            StartingPos := txtStartingPos.value
        }else{
            StartingPos := InStr(txtClipboard.text, txtfind.text, chbCaseSense.value, 1, txtOccurrence.value)
        }
        A_Clipboard := StrReplace(A_Clipboard, txtFind.text , txtReplace.text, chbCaseSense.value, , limit)
        try 
            txtClipboard.Value := A_Clipboard
    }
}
