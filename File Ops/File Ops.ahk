/* INFO ----------------------------------------------------------------------------------------
    Author        :   Anastasiou Alex
    Email         :   AnastasiouAlex@gmail.com
    Blog          :   https://AlexOfRhodes.github.io/
    Github        :   https://github.com/AlexOfRhodes
    Youtube       :   https://bit.ly/3aLZU9M
    Vk            :   https://vk.com/video/playlist/735281600_1
------------------------------------------------------------------------------------------------
    Project       :   File Ops
    Description   :   Interface for common operations on files:
                        * Selected in ListView 
                                - you can add/remove folders and list their files
                                - custom list of files by drag and drop
                        * Selected in File Explorer
                        
                        For new Functions just include its name to the CommandList
------------------------------------------------------------------------------------------------ 
    ChangeLog       :   Revision    Date            Note
                        1.0.0       2023-06-21      First release          
------------------------------------------------------------------------------------------------ */

#Requires AutoHotkey >= 2.0
#SingleInstance Force
SetWorkingDir A_ScriptDir
#include anchorv2.ahk
; #include v2_Class_LV_Colors.ahk

global CommandList := 
[
    "aTest",
    "SmartUnzip",
    "TxtRemoveBlankLines",
    "TxtMerge",
    "TxtSplitVbaProcedures",
]

; === GUI 
{
    ; TRAY MENU
    {
        I_Icon := A_ScriptDir "\icons\FILE.PNG" 
        iF FileExist(I_Icon)
            TraySetIcon(I_Icon)
        OnMessage(0x404, AHK_NOTIFYICON)
        
        Tray := A_TrayMenu
        ; tray.delete ; Delete the standard items.
        Tray.Add()  ; Creates a separator line.

        ; VOID, when right clicking on the tray, the file explorer loses focus (needed for function getTarget)
            ; for ,item in CommandList
            ;     Tray.add(item, CallByTray)

            ; CallByTray(ctr,*){
            ;     oByFileExplorer.Value := 1
            ;     ListFiles
            ;     GetTarget
            ;     f:=%ctr%, f()
            ; }
        
        Tray.Add("Github", FollowLink)            ; +BarBreak starts new column
        Tray.SetIcon("Github","icons\github.ico")

        Tray.Add("YouTube", FollowLink)
        Tray.SetIcon("YouTube","icons\youtube.ico")

        Tray.Add("Blog", FollowLink)   
        Tray.SetIcon("Blog","icons\Blog.ico")

        Tray.Add("Email", FollowLink)  
        Tray.SetIcon("Email","icons\gmail.ico")

        Tray.Add("BuyMeACoffee", FollowLink) 
        Tray.SetIcon("BuyMeACoffee","icons\BuyMeACoffee.ico")
    }

    ; LINKS
    {        
        FollowLink(ControlName, info, *)
        {
            f:=% "openLink" Controlname%, f()

            ; switch ControlName, false
            ; {
            ; case "Github":          OpenLinkGithub  
            ; case "Youtube":         OpenLinkYouTube
            ; case "Blog":            OpenLinkBlog
            ; case "BuyMeACoffee":    OpenLinkBuyMeACoffee
            ; case "Email":           OpenLinkEmail 
            ; }   
        }

        OpenLinkGithub(*) {
            run "https://github.com/alexofrhodes/" 
        }
        OpenLinkYoutube(*) {
            run "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
        }
        OpenLinkBlog(*) {
            run "https://alexofrhodes.github.io/"
        }
        OpenLinkBuyMeACoffee(*) {
            run "https://www.buymeacoffee.com/AlexOfRhodes"
        }
        OpenLinkEmail(*){
            Run "mailto:anastasioualex@gmail.com?subject=AutoHotkey - File Ops&body=Hi! I would like to talk about a bug/suggestion/request."
        }
    }

    MyGui := Gui()
    mygui.SetFont("q5")

    global SB_Default := ""   
    SB := MyGui.AddStatusBar(, SB_Default)

    ; TARGET FILES FROM
    {        
        ; By Folder

            oByFolder := mygui.AddRadio("ym+10","By Folder")
            oByFolder.OnEvent("Click", ListFiles)
            oByFolder.Focus

        ; By Selection in File Explorer

            oByFileExplorer := mygui.AddRadio("ym+10","By File Explorer")
            oByFileExplorer.OnEvent("Click", ListFiles)
            
        ; Drag and Drop

            oByDragDrop := mygui.AddRadio("ys","By Drag n Drop")
            oByDragDrop.OnEvent("Click", ListFiles)
            oRemoveSelected := mygui.addbutton("ys-5 section -Tabstop","Remove Selected")
            oRemoveSelected.OnEvent("Click",RemoveSelected)
            oClearDD := mygui.AddButton("ys -Tabstop","Clear DD")
            oClearDD.OnEvent("Click", ClearDD)
    }
    ; FOLDER PICKER
    {    
        global myFolders := []
        myFolders := StrSplit(IniRead(A_ScriptDir . "\config.ini", "Settings", "myFolders", A_ScriptDir . "\"), "|")

        oFolderPicker := MyGui.AddComboBox("xm section w500 ", myFolders)
        oFolderPicker.OnEvent("Change",ListFiles)
        ; oFolderPicker.Focus

        AddBtn := MyGui.AddButton("-Tabstop ys section h22 w22", "+")
        AddBtn.OnEvent("Click", AddBtn_Click)  
        AddBtn.SetFont("bold s12 q5")

        RemoveBtn := MyGui.AddButton("-Tabstop ys h22 w22", "-")
        RemoveBtn.OnEvent("Click", RemoveBtn_Click)  
        RemoveBtn.SetFont("bold s12 q5")
    
    }
    ; FILTER
    {
        mygui.AddText("xm y+10 section","Filter")
        oFilter := mygui.AddEdit("ys-5 w150", "*")
        oFilter.OnEvent("Change", ListFiles)
        ofilter.SetFont("s9 ")
    }
    ; COMMANDS                   
    {
        mygui.AddText("ys","Commands")
        oCommands := mygui.AddDropDownList("ys-3 w190 Backgroundffffff  ", CommandList) 
        
        oRunner := MyGui.AddButton("ys-5 h22 ", "&R U N")
        orunner.OnEvent("Click", RunCommand)
        
    }

    ; PROGRESSBAR

    myProgressBar := mygui.AddProgress("xs section w450 Range0-10 smooth" )   ;Vertical
    myProgressBar.value := 0
    mygui.AddText("ys+1 x+10","%")

    ; LISTVIEW
    {

        ; global LVS_EX_TRANSPARENTBKGND := "LV0x00400000"
        oListView := myGui.AddListView(" xm w500 section R11", ["FILE"]) 
        oListView.Opt("+Grid +multi +BackgroundF2F9F1")   ;Background FAFDD6 FFFCEF 
        ; WantF2 -ReadOnly   ;allow edit  
        ; LV.ModifyCol()     ;Auto-sizes all columns to fit their contents.

        ; oListView.OnEvent("ColClick", LV_SortArrow)

        ;TODO drag drop rearrange

        ; oListView.OnNotify(-109, LVN_BEGINDRAG)
        ; LVN_BEGINDRAG(LV, LPARAM) {
        ;     Row := NumGet(LPARAM + (A_PtrSize * 3), "Int") + 1 ; 1-based row number
        ;     MsgBox("You started to drag row " . Row)
        ; }
        
        ; --- v2_Class_LV_Colors.ahk
        ;-------------------------------------
        ; CLV := LV_Colors(oListView)
        ; If !IsObject(CLV) {
        ;     MsgBox("Couldn't create a new LV_Colors object!", "ERROR", 16)
        ;     ExitApp
        ; }
        ; ; Set the colors for selected rows
        ; CLV.SelectionColors(0xF0F0F0)
        ; oListView.Opt("-Redraw")
        ; CLV.AlternateRows(0x808080, 0xFFFFFF)
        ; oListView.Opt("+Redraw")
        ; oListView.Focus()
        ;-------------------------------------
    }
    ; LOAD OPTIONS / FILES
    {
        LoadOptions
        ListFiles
    }
    ; FINALIZE GUI
    {   
        MyGui.OnEvent("Size", OnSize)
        mygui.OnEvent("Escape", onEscape)
        mygui.OnEvent("Close", onClose)
        myGui.OnEvent("DropFiles",OnDropFiles)

        global WM_MOUSEMOVE  := 0x0200
        global WM_MOUSELEAVE := 0x02A3
        OnMessage(WM_MOUSEMOVE , OnMouseEvent)
        OnMessage(WM_MOUSELEAVE, OnMouseEvent)

        imageGithub          := MyGui.AddPicture( "ys+25 x+25 w28 h28 SECTION", "icons\github.ico").OnEvent("Click",OpenLinkGithub )
        imageBlog            := MyGui.AddPicture( "Xs         w28 h28", "icons\blog.ico").OnEvent("Click",OpenLinkBlog )
        imageEmail           := MyGui.AddPicture( "Xs         w28 h28", "icons\gmail.ico").OnEvent("Click",OpenLinkEmail )
        imageBuyMeACoffee    := MyGui.AddPicture( "Xs         w28 h28", "icons\BuyMeACoffee.ico").OnEvent("Click",OpenLinkBuyMeACoffee )
        imageYouTube         := MyGui.AddPicture( "Xs         w28 h28", "icons\YouTube.ico").OnEvent("Click",OpenLinkYoutube )

        myGui.Opt("Resize +AlwaysOnTop -DPIScale " )    ; + +MinSize
        mygui.BackColor := "FFFFFF"
       if (GuiX > 0){
             MyGui.Show("x" GuiX "y" GuiY "autosize" )
            if oByFileExplorer.value=true 
                mygui.Move(,,,220)
        }else{
            mygui.show("AutoSize center")
        }
        
    }
}

; === FUNCTIONS                 
{

    aTest(*){
        ; FileArray := StrSplit(SelectedFiles, "`n`r")
        myProgressBar.Opt("Range0-" SelectedFiles.length)
        For Index, TargetFile in SelectedFiles
            {
                Sleep(200)
                if !FileExist(TargetFile) 
                    continue
                msg .= "`n" TargetFile
                myProgressBar.value += 1
            }        
            SB_Set("message complete")
            MsgBox(msg)
    }

    SB_Set(msg, MessageAlert:=false, PlaySound:=false){
        global SB_Default := msg
        SB.text := msg
        if messageAlert
            MsgBox(msg)
        else if PlaySound
            SoundBeep
    }

    TxtSplitVbaProcedures(*){
        TxtRegexExport("mi)((Public|Private|Friend)\s){0,1}(Static\s){0,1}(Function|Sub|Property\sGet|Property\sLet|Property\sSet)\s{0,1}([a-zA-Z0-9_]*?)\s{0,1}\([\S\s]*?End\s(Function|Sub|Property)")
    }

    TxtRegexExport(RegexPattern, MergeThem:=False, sep := "`n`r"){  
        ;  Define the regex pattern
        ;  regexPattern := "\b\w+\b"  ; Example regex pattern - Word boundary pattern

        myProgressBar.Opt("Range0-" SelectedFiles.length)
    
        For Index, TargetFile in SelectedFiles
        {
            Sleep(200)
            if !FileExist(targetFile) or (SubStr(targetFile, -4) != ".txt"){
                myProgressBar.value += 1
                continue
            }
            SplitPath(targetFile, &fileName, &folder, , &fileName_noext)
            outputFolder := folder "\regex output"
            ; Create the output folder if it doesn't exist
            if (!DirExist(outputFolder))
                DirCreate(outputFolder)
            fileContent := FileRead(targetFile, "`n UTF-8")

            TimeStamp := A_Now
            n :=1
            out := []
            while (RegExMatch(filecontent, RegexPattern, &match, n)) {
                str:=match[]
                out.Push(str)
                foundName:= match.5
                if !(foundName="")
                    outputFileName :=  outputFolder "\" foundName ".txt"
                else
                    outputFileName := outputFolder "\" TimeStamp " " out.Length ".txt"
                
                if !MergeThem
                    FileAppend("`n" str, outputFileName,"`n UTF-8")
                n := match.pos + strlen(str)                        
            }

            matchCount := out.length

            if MergeThem{
                str:=""
                for index, value in out
                    str .= value . sep
                str := substr(str,1,StrLen(str)-StrLen(sep))
                outputFileName := outputFolder "\" TimeStamp " - " matchCount " - items.txt"
                FileAppend("`n" str, outputFileName,"`n UTF-8")
            }
            myProgressBar.value += 1    
        }
        SB_Set("Exported " matchCount " matches from " SelectedFiles.Length " files")
    }

    TxtRemoveBlankLines(){
        myProgressBar.Opt("Range0-" SelectedFiles.length)
        For Index, TargetFile in SelectedFiles
        {
            Sleep(200)
            if !FileExist(targetFile) or (SubStr(targetFile, -4) != ".txt"){
                myProgressBar.value += 1
                continue
            }
            SplitPath(targetFile, &fileName, &folder, , &fileName_noext)
            fileContent := FileRead(targetFile, "`n UTF-8")
            FileDelete(targetFile) 
            FileAppend(fileContent, folder "\bkup_" . fileName)
            lines := StrSplit(fileContent, "`n")
            newContent := ""
            ; Loop through the lines and remove empty lines
            for index, line in lines {
                if (line != ""){
                    newContent .= line
                    if (index < lines.Length)
                        newContent .= "`n"
                }
            }
            FileAppend(newContent, targetFile,"`n UTF-8")
            FileDelete(folder "\bkup_" . fileName) 
            myProgressBar.value += 1
        }
        MsgBox("Empty lines removed from " SelectedFiles.Length " files")  ;oListView.GetCount("Selected")
    }

    TxtMerge(OutputFile := ""){
        myProgressBar.Opt("Range0-" SelectedFiles.length)
        Output := ""
        For Index, TargetFile in SelectedFiles
            {
                Sleep(200)
                if !FileExist(targetFile) or (SubStr(targetFile, -4) != ".txt"){
                    myProgressBar.value += 1
                    continue
                }
                if (OutputFile = ""){
                    SplitPath(targetFile, , &folder, , &zipname_noext)
                    OutputFile := FileSelect(24 , folder, "Merging - Select output file", "Text Documents (*.txt)")
                    if (OutputFile = "")
                        return                    
                }

                Output .= "`n" . FileRead(targetFile, "`n UTF-8")
                if (targetFile = OutputFile)
                    FileDelete(targetFile) 
                myProgressBar.value += 1
        }
        FileAppend(Output, OutputFile)
        MsgBox("Merged " SelectedFiles.Length " files to " OutputFile)
    } 

    ListFiles(*){
        EM_SETSEL := 0x00B1
        if (oFilter.Text = ""){
            ofilter.text := "*"
            SendMessage(EM_SETSEL, 1, 1,, "ahk_id" ofilter.Hwnd)
;                                  ^  ^            
;                           startPos, endPos                                        
        }
        oListView.Delete()
        if  (oByFolder.value = true){
            oFolderPicker.enabled := true
            Loop Files, oFolderPicker.text . oFilter.Text
                oListView.Add(,A_LoopFileName)
            controlshow oListView
            mygui.Move(,,,460) 

        }else if (oByDragDrop.value = true){
            oFolderPicker.Enabled := false  
            ListViewContent := StrSplit(IniRead(A_ScriptDir . "\config.ini", "Settings", "DragDropList"), "|")
            if (ListViewContent.Length > 0)
                for index, item in ListViewContent
                    if FileExist(item)
                        if (RegExMatch(item, "i)" StrReplace(oFilter.Text,"*",".*")))
                            oListView.Add(,item)                              
    
            controlshow oListView
            mygui.Move(,,,460) 
        }else if (oByFileExplorer.value = true){
            oFolderPicker.Enabled := false 
            controlhide oListView
            mygui.Move(,,,220) 
        }
        setStatusBarMessage
        ; TODO remove the following test
        ; WinGetPos(,,,&H,mygui.Hwnd) 
        ; MsgBox(H)
    }
    
    RegExEscape(text) {
        return RegExReplace(text, "[.*+?^$|(){}[\]\\]", "`$&")
    }

    setStatusBarMessage(*){
        if  (oByFolder.value = true)
            SB_Set("Select files from ListView")
        else if (oByDragDrop.value = true)
            SB_Set("Drag n Drop Baby")
        if (oByFileExplorer.value = true)
            SB_Set("Targeting selected files in File Explorer")
    }

    GetTarget(*){
        global SelectedFiles := []  ;""
        if (oByFileExplorer.value=true){
            mygui.Hide
            Sleep(200)
            SelectedFiles := Explorer_GetSelection() 
            
        }else if (oByDragDrop.value=true){
            RowNumber := 0                                  ; This causes the first loop iteration to start the search at the top of the list.
            Loop
            {
                RowNumber := oListView.GetNext(RowNumber)   ; Resume the search at the row after that found by the previous iteration.
                if not RowNumber                            ; The above returned zero, so there are no more selected rows.
                    break
                Text := oListView.GetText(RowNumber)
                SelectedFiles.Push(Text) ; .= Text
            }
        }else if (oByFolder.value=true){
            RowNumber := 0                                  ; This causes the first loop iteration to start the search at the top of the list.
            Loop
            {
                RowNumber := oListView.GetNext(RowNumber)   ; Resume the search at the row after that found by the previous iteration.
                if not RowNumber                            ; The above returned zero, so there are no more selected rows.
                    break
                Text := oFolderPicker.text . oListView.GetText(RowNumber)
                SelectedFiles.Push(Text) ; .= Text
            }
        }

        if (SelectedFiles.Length=0) ;(SelectedFiles = "")
            return
    }

    SmartUnzip(*)
    {
        ;Smarter unzip by nod5  181114
        ;https://www.donationcoder.com/forum/index.php?topic=46655.0
        myProgressBar.Opt("Range0-" SelectedFiles.length)
        For Index, TargetFile in SelectedFiles
        {
            Sleep(200)
            if !FileExist(targetFile) or (SubStr(targetFile, -4) != ".zip"){
                myProgressBar.value += 1
                continue
            }

            ;unzip to temporary folder
            SplitPath(targetFile, , &folder, , &zipname_noext)
            tempfolder := folder "\temp_" A_Now
            DirCreate(tempfolder)
            Unzip(targetFile, tempfolder)

            ;count unzipped files/folders recursively
            Loop Files, tempfolder "\*.*", "FDR"
            {
                allcount := A_Index
                onlyitem := A_LoopFilePath
            }
            
            ;count unzipped files/folders at first level only
            Loop Files, tempfolder "\*.*", "FD"
            {
                firstlevelcount := A_Index
                firstlevelitem  := A_LoopFilePath
            }
            
            ;case1: only one file/folder in whole zip
            if (allcount = 1)
            {
                if InStr(FileExist(onlyitem), "D")
                {
                    SplitPath(onlyitem, &onlyfoldername)
                    ; if (onlyfoldername = zipname_noext)   ; TODO consider some cases..
                    ; DirMove(onlyitem, folder "\" onlyfoldername)  ; pull that folder out  
                    DirMove(tempfolder, folder "\" zipname_noext)   ; rename that folder to the zip's name
                }else{
                    DirMove(tempfolder, folder "\" zipname_noext)   
                    ;FileMove(onlyitem, folder)                     ; obsolete
                }
            }
            ;case2: only one folder (and no files) at the first level in zip
            else if (firstlevelcount = 1) and InStr(FileExist(firstlevelitem), "D")
            {
                SplitPath(firstlevelitem, &firstlevelfoldername)
                DirMove(firstlevelitem, folder "\" firstlevelfoldername)
            }else{ ;case3: multiple files/folders at the first level in zip
                DirMove(tempfolder, folder "\" zipname_noext)
            }
            ;cleanup temp folder
            try
                DirDelete(tempfolder, 1)
            myProgressBar.value += 1
        }

        ;refresh Explorer to show results
        If WinActive("ahk_class CabinetWClass")
            Send("{F5}")
        
    }

    ;function: unzip files to already existing folder - zip file can have subfolders
    Unzip(zipfile, folder)
    {
        psh := ComObject("Shell.Application")
        psh.Namespace(folder).CopyHere( psh.Namespace(zipfile).items, 4|16)
    }

    Explorer_GetSelection() {
        ; https://www.autohotkey.com/boards/viewtopic.php?style=17&t=60403#p255169
        result := [] ; ""
        hHwnd := hWnd := WinExist("A")
        winClass := WinGetClass("ahk_id" . hHwnd)
        if !(winClass ~= "^(Progman|WorkerW|(Cabinet|Explore)WClass)$")
            Return 
        shellWindows := ComObject("Shell.Application").Windows
        if (winClass ~= "Progman|WorkerW")
            shellFolderView := shellWindows.Item( ComValue(VT_UI4 := 0x13, SWC_DESKTOP := 0x8) ).Document
        else {
            for window in shellWindows 
            if (hWnd = window.HWND) && (shellFolderView := window.Document)
                break
            }
        for item in shellFolderView.SelectedItems
            result.Push(item.path) ; .= (result = "" ? "" : "`n") . item.Path
        Return result
    }

    ;f;unction: copy selection to clipboard to var
    ClipToVar() {
        cliptemp := ClipboardAll() ;backup
        A_Clipboard := ""
        Send("^c")
        Errorlevel := !ClipWait(1)
        clip := A_Clipboard
        A_Clipboard := cliptemp    ;restore
        return clip
    }

    ;TODO TargetLVSelection is a sample, will remove later
    TargetLVSelection(* ){   
        /* NOTE
            (onEvent("Click" 
                passes by default 2 arguments: ctrl, info
            For Custom arguments like TargetLVSelection(ctrl,info, <Custom Arguments>) 
                we need to use onEvent("Click",TargetLVSelection.bind(,,"CustomArg")
        */
        ; MsgBox(ctr.Text)    


        RowNumber := 0  ; This causes the first loop iteration to start the search at the top of the list.
        Loop
        {
            RowNumber := oListView.GetNext(RowNumber)  ; Resume the search at the row after that found by the previous iteration.
            if not RowNumber  ; The above returned zero, so there are no more selected rows.
                break
            Text := oListView.GetText(RowNumber)
            MsgBox('The next selected row is #' RowNumber ', whose first field is "' Text '".')
        }
    }
}

;=== EVENTS 
{

    OnMouseEvent(wp, lp, msg, hwnd) { ; 

        global PreviousControl
        PreviousHwnd := 0 
        static TME_LEAVE := 0x2, onButtonHover := false
        if msg = WM_MOUSEMOVE && !onButtonHover ; && ButtonHwnd = hwnd 
        { 
            TRACKMOUSEEVENT := Buffer(8 + A_PtrSize * 2)
            NumPut('UInt', TRACKMOUSEEVENT.Size,
                'UInt', TME_LEAVE,
                'Ptr', hwnd,
                'Ptr', 10, TRACKMOUSEEVENT)
            DllCall('TrackMouseEvent', 'Ptr', TRACKMOUSEEVENT)
            
            myControl := GuiCtrlFromHwnd(hwnd)
            try{
                if InStr(myControl.Text, ".ico"){
                    if !(myControl.hwnd=PreviousHwnd){
                        ; ControlSetStyle("+0x800000", myControl, mygui.Hwnd)
                        sb.text := SB_Default "`t`tCLICK ME !"

                        ; For GuiCtrHwnd, GuiCtrlObj in MyGui
                        ;     if (InStr(GuiCtrlObj.Text, ".ico")>0) && !(hwnd=GuiCtrHwnd){
                        ;         ControlGetPos(&x, &y, &w, &h, GuiCtrlObj)
                        ;         ControlHide(GuiCtrlObj)
                        ;     }

                        PreviousControl := myControl
                        PreviousHwnd := PreviousControl.Hwnd
                    }
                }            
            }
        }
        if msg = WM_MOUSELEAVE {
            if IsSet(PreviousControl){
                ; ControlSetStyle("-0x800000", PreviousControl, mygui.Hwnd)
                sb.text := SB_Default

                ; For GuiCtrHwnd, GuiCtrlObj in MyGui
                ;     if (InStr(GuiCtrlObj.Text, ".ico")>0) {
                ;         ControlShow(GuiCtrlObj) 
                ;     }                
            }
        }
    }

    AHK_NOTIFYICON(wParam, lParam, *)
    {
        if (lParam = 0x201) ; WM_LBUTTONDOWN
        {
            mygui.show
            return 0
        }
    }

    OnDropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
        if oByDragDrop{
            ofilter.text := "*"
            for index, item in FileArray
                oListView.Add(,item)
            SaveDragDropList
        }
    }
    
    OnSize(GuiObj, MinMax, Width, Height){
        if (oListView.Visible = true)
            Anchor(oListView.Hwnd,"h")
    }

    onEscape(*){
        mygui.Minimize
    }

    onClose(*) { 
        SaveOptions      
    }

    ; LV_SortArrow(LV, Column) {
    ;     ;https://www.autohotkey.com/boards/viewtopic.php?t=94353
    ;     static SORTUP         := Chr(0x25B2)  ; (▲) Black Up-Pointing Triangle
    ;     static SORTDOWN       := Chr(0x25BC)  ; (▼) Black Down-Pointing Triangle
    ;     static HDI_TEXT       := 0x0002
    ;     static HDI_FORMAT     := 0x0004
    ;     static HDF_RIGHT      := 0x0001
    ;     static LVM_FIRST      := 0x1000
    ;     static HDM_FIRST      := 0x1200
    ;     static HDM_GETITEM    := HDM_FIRST + 11
    ;     static HDM_SETITEM    := HDM_FIRST + 12
    ;     static LVM_GETHEADER  := LVM_FIRST + 31
    ;     static HDITEM_SIZE    := 6 * 4 + 6 * A_PtrSize
    ;     static TextPtrOffset  := 2 * 4
    ;     static TextSizeOffset := 2 * 4 + 2 * A_PtrSize
    ;     static FormatOffset   := 3 * 4 + 2 * A_PtrSize
    ;     static MaxChars       := 260
    ;     static PrevColumn     := 0
    ;     Text := Buffer(2 * MaxChars)
    ;     HDITEM := Buffer(HDITEM_SIZE)
    ;     NumPut 'UInt', HDI_TEXT | HDI_FORMAT, HDITEM, 0
    ;     NumPut 'Ptr', Text.Ptr, HDITEM, TextPtrOffset
    ;     NumPut 'Int', MaxChars, HDITEM, TextSizeOffset
    ;     HDR := SendMessage(LVM_GETHEADER, 0, 0, LV)
    ;     if PrevColumn && Column != PrevColumn {
    ;         SendMessage(HDM_GETITEM, PrevColumn - 1, HDITEM, HDR)
    ;         Format := NumGet(HDITEM, FormatOffset, 'Int'), String := StrGet(Text)
    ;         StrPut Format & HDF_RIGHT ? SubStr(String, 3) : SubStr(String, 1, -2), Text
    ;         SendMessage(HDM_SETITEM, PrevColumn - 1, HDITEM, HDR)
    ;     }
    ;     PrevColumn := Column
    ;     SendMessage(HDM_GETITEM, Column - 1, HDITEM, HDR)
    ;     Format := NumGet(HDITEM, FormatOffset, 'Int'), String := StrGet(Text)
    ;     if Format & HDF_RIGHT
    ;         String := RegExMatch(String, '^(' SORTUP '|' SORTDOWN ')(.*)', &Match) ? (Match[1] == SORTDOWN ? SORTUP : SORTDOWN) Match[2] : SORTUP A_Space String
    ;     else
    ;         String := RegExMatch(String, '(.*)(' SORTUP '|' SORTDOWN ')$', &Match) ? Match[1] (Match[2] == SORTDOWN ? SORTUP : SORTDOWN) : String A_Space SORTUP
    ;     StrPut String, Text, MaxChars - 1
    ;     return SendMessage(HDM_SETITEM, Column - 1, HDITEM, HDR)
    ; }    
}

; === INI 
{
    LoadOptions(*){

        try{
            oFolderPicker.value     := IniRead(A_ScriptDir . "\config.ini", "Settings", "SelectedFolder",1)
        }catch{
            oFolderPicker.Add([A_ScriptDir "\"])
            oFolderPicker.value := 1
        }
        oFilter.text            := IniRead(A_ScriptDir . "\config.ini", "Settings", "Filter", "*")

        oByFolder.value         := IniRead(A_ScriptDir . "\config.ini", "Settings", "byFolder",True)
        oByDragDrop.value       := IniRead(A_ScriptDir . "\config.ini", "Settings", "byDragDrop",False)
        oByFileExplorer.value   := IniRead(A_ScriptDir . "\config.ini", "Settings", "byFileExplorer",False)
        
        setStatusBarMessage
        
        try
            oCommands.value            := IniRead(A_ScriptDir . "\config.ini", "Settings", "LastCommand",0)
        catch
            oCommands.value := 0

        global GuiX                    := IniRead(A_ScriptDir . "\config.ini", "Settings", "GuiX",0)
        global GuiY                    := IniRead(A_ScriptDir . "\config.ini", "Settings", "GuiY",0)
        
        if (oByFolder.Value = false)
            oFolderPicker.Enabled := false
        ;     {
        ;         ListViewContent := StrSplit(IniRead(A_ScriptDir . "\config.ini", "Settings", "DragDropList"), "|")
        ;         for each, item in ListViewContent
        ;             oListView.Add(,item)
                
        ;     }
    }

    SaveOptions(*){
        mygui.Submit(false)
        if (oFolderPicker.value>0)
            IniWrite(oFolderPicker.value,   A_ScriptDir . "\config.ini", "Settings", "SelectedFolder")

        IniWrite(oFilter.text,              A_ScriptDir . "\config.ini", "Settings", "Filter")
        
        IniWrite(oByFolder.value,           A_ScriptDir . "\config.ini", "Settings", "byFolder")
        IniWrite(oByDragDrop.value,         A_ScriptDir . "\config.ini", "Settings", "byDragDrop")
        IniWrite(oByFileExplorer.value,     A_ScriptDir . "\config.ini", "Settings", "ByFileExplorer")

        ; if (oByDragDrop.Value = true)
        ;     SaveDragDropList
        
        IniWrite(oCommands.value, A_ScriptDir . "\config.ini", "Settings", "LastCommand")

        MyGui.GetPos(&GuiX, &GuiY, &GuiW, &GuiH)
        IniWrite(GuiX, A_ScriptDir . "\config.ini", "Settings", "GuiX")
        IniWrite(GuiY, A_ScriptDir . "\config.ini", "Settings", "GuiY")

    }

    saveMyFolders(*){
        joinedList := ""
        myFolders := ControlGetItems(oFolderPicker)
        for index, item in myFolders 
            {
                joinedList .= item
                if (index < myFolders.Length)
                    joinedList .= "|"
            }            
        IniWrite(joinedList, A_ScriptDir . "\config.ini", "Settings", "myFolders")
        
    }

    SaveDragDropList(*){
        joinedList := ""
        count := oListView.GetCount()
        Loop Count
            {
                RetrievedText := oListView.GetText(A_Index)
                joinedList .= RetrievedText 
                if (A_index < count)
                    joinedList .= "|"                
            }
        
        IniWrite(joinedList, A_ScriptDir . "\config.ini", "Settings", "DragDropList")
    }
}

;  === BUTTONS 
{
    RunCommand(*){
        getTarget    
        f := %oCommands.text%, f()  ; call function by name
    }
    AddBtn_Click(*)
    {
        SelectedFolder := DirSelect("*" . A_ScriptDir, 3 , "Select a folder")
        if (SelectedFolder = "")
            return
        SelectedFolder .= "\"
        for index, item in myFolders
            if (item = SelectedFolder)
                return
        myFolders.Push(SelectedFolder)
        oFolderPicker.Add([SelectedFolder])         ; Update ComboBox items
        oFolderPicker.value := myFolders.Length     ; Set selected item = last
        mygui.Submit(false)
        saveMyFolders
        ListFiles
    }
    RemoveBtn_Click(*)
    {
        previousChoice := oFolderPicker.Value
        oFolderPicker.Delete(oFolderPicker.Value)
        try 
            oFolderPicker.value := previousChoice
        catch
            try
                oFolderPicker.value := previousChoice - 1
        saveMyFolders
        ListFiles
    }
    ClearDD(*){
        if (oByDragDrop.value = true){
            oListView.Delete
            IniWrite("", A_ScriptDir . "\config.ini", "Settings", "DragDropList")
        }
    }
    RemoveSelected(*){
        ListViewSelectedRows := []
        RowNumber := 0  ; This causes the first loop iteration to start the search at the top of the list.
        Loop
        {
            RowNumber := oListView.GetNext(RowNumber)  ; Resume the search at the row after that found by the previous iteration.
            if not RowNumber  ; The above returned zero, so there are no more selected rows.
                break
            ListViewSelectedRows.InsertAt(1,RowNumber)
        }
        For index, row in ListViewSelectedRows
            oListView.Delete(row)
        if oByDragDrop.Value=true
            SaveDragDropList
    }
}

; === HOTKEYS 
{
    #HotIf WinActive(mygui.Hwnd)     ;(oListView.Focused=true)
    ^a::{
        ControlFocus(oListView)
        global AllSelected
        if !IsSet(AllSelected)
            AllSelected:=false
        if (AllSelected=true)
            {
            oListView.Modify(0,"-Select")
            AllSelected := false
            }
        else{
            oListView.Modify(0,"+Select")
            AllSelected := true
        }
    }
    #HotIf   
}



