
/* INFO ----------------------------------------------------------------------------------------
    Author        :   Anastasiou Alex
    Email         :   AnastasiouAlex@gmail.com
    Blog          :   https://AlexOfRhodes.github.io/
    Github        :   https://github.com/AlexOfRhodes
    Youtube       :   https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
    Vk            :   https://vk.com/video/playlist/735281600_1
------------------------------------------------------------------------------------------------
    Project       :   Typing Simulator
    Description   :   Drag files to the gui to add them to the list.
                      To abort the simulation right click on the tray icona and exit app.
------------------------------------------------------------------------------------------------ 
    ChangeLog       :   Revision    yyyy-mm-dd      Note
                        1.0.0       2023-07-15      First release          
------------------------------------------------------------------------------------------------ */

#Requires AutoHotkey >=2.0

#SingleInstance Force
SetWorkingDir A_ScriptDir

;---GUI
{
    ; TRAY MENU
    {
        I_Icon := A_ScriptDir "\icons\typeSim.PNG" 
        iF FileExist(I_Icon)
            TraySetIcon(I_Icon)
        OnMessage(0x404, AHK_NOTIFYICON)
        
        Tray := A_TrayMenu
        ; tray.delete ; Delete the standard items.
        Tray.Add()  ; Creates a separator line.
        
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

    global GuiX:=0
    global GuiY:=0

    MyGui := Gui()
    mygui.SetFont("q5")

    ; FILE PICKER
    
    oFilePicker := MyGui.AddComboBox("xm section w600 ")
    LoadFileList
    if ControlGetItems(oFilePicker).Length=1
        oFilePicker.value := 1

    ; AddBtn := MyGui.AddButton("-Tabstop ys section h22 w22", "+")
    ; AddBtn.OnEvent("Click", AddBtn_Click)  
    ; AddBtn.SetFont("bold s12 q5")

    RemoveBtn := MyGui.AddButton("-Tabstop ys h22 w22", "-")
    RemoveBtn.OnEvent("Click", RemoveBtn_Click)  
    RemoveBtn.SetFont("bold s12 q5")

    mygui.OnEvent("Escape", onEscape)
    mygui.OnEvent("Close", onClose)
    myGui.OnEvent("DropFiles",OnDropFiles)

    imageGithub          := MyGui.AddPicture( "xs w24 h24 section", "icons\github.ico").OnEvent("Click",OpenLinkGithub )
    imageBlog            := MyGui.AddPicture( "ys w24 h24", "icons\blog.ico").OnEvent("Click",OpenLinkBlog )
    imageEmail           := MyGui.AddPicture( "ys w24 h24", "icons\gmail.ico").OnEvent("Click",OpenLinkEmail )
    imageBuyMeACoffee    := MyGui.AddPicture( "ys w24 h24", "icons\BuyMeACoffee.ico").OnEvent("Click",OpenLinkBuyMeACoffee )
    imageYouTube         := MyGui.AddPicture( "ys w24 h24", "icons\YouTube.ico").OnEvent("Click",OpenLinkYoutube )

    myGui.Add("Button", "ys x+50 w80 section", "Type away").OnEvent("click",SimulateTyping)
    chUseClipboard := mygui.add("CheckBox","xs","Use Clipboard")
    chUseClipboard.OnEvent("Click", chUseClipboardClicked)  

    mygui.AddText("ys+3", "delay ms")
    KeyDelay := MyGui.AddComboBox("ys+2 w50",[10,20,30,40,50,60,70,80,90,100])
    keydelay.Value := 5

    myGui.Add("Button", "ys w80", "Edit File").OnEvent("click",EditFile(*) =>run(oFilePicker.text))
    myGui.Add("Button", "ys w80", "Edit Ini").OnEvent("click",EditIni(*) =>run("config.ini"))

    myGui.Opt("Resize +AlwaysOnTop -DPIScale -MaximizeBox " )    ; +MinSize
    mygui.BackColor := "FFFFFF"

    LoadOptions

    MyGui.AddText("xm", "Drag & Drop Files on GUI appends combobox. Once the Typing simulation begins if you want to stop it, right click on tray to exit app.")

    if (GuiX > 0){
        MyGui.Show("x" GuiX "y" GuiY "autosize" )
    }else{
        mygui.show("AutoSize center")
    }

}

chUseClipboardClicked(*){
    oFilePicker.Enabled := !chUseClipboard.value
}

SimulateTyping(*){
    SendMode("event")
    SetKeyDelay KeyDelay.text ;50
    mygui.Hide

    if chUseClipboard.value = true
        txt :=  A_Clipboard
    else
        txt := Fileread(oFilePicker.Text, "`n")

    SendText(txt)
    mygui.show
}

; AddBtn_Click(*)
; {
; if you want to add files from file picker code it here
; }

RemoveBtn_Click(*)
{
    previousChoice := oFilePicker.Value
    oFilePicker.Delete(previousChoice)
    try 
        oFilePicker.value := previousChoice
    catch
        try
            oFilePicker.value := previousChoice - 1
    SaveFileList
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
    for index, item in FileArray
    {
        if (InStr(item, ".txt") > 0)
        {
            pass := true
            myFiles := ControlGetItems(oFilePicker)
            for cbindex, cbitem in myFiles 
                {
                    if cbitem = item
                    { 
                        pass := false
                        break
                    }
                }       
            if (pass = true)
                oFilePicker.Add([item])
        }
    }
    SaveFileList
}

onEscape(*){
    mygui.Minimize
}

onClose(*) { 
    SaveOptions      
}

; === INI 
{
    LoadOptions(*){

        global GuiX                    := IniRead(A_ScriptDir . "\config.ini", "Settings", "GuiX",0)
        global GuiY                    := IniRead(A_ScriptDir . "\config.ini", "Settings", "GuiY",0)
        
    }

    SaveOptions(*){
        mygui.Submit(false)

        MyGui.GetPos(&GuiX, &GuiY, &GuiW, &GuiH)
        IniWrite(GuiX, A_ScriptDir . "\config.ini", "Settings", "GuiX")
        IniWrite(GuiY, A_ScriptDir . "\config.ini", "Settings", "GuiY")

    }

    SaveFileList(*){
        joinedList := ""
        myFiles := ControlGetItems(oFilePicker)
        for index, item in myFiles 
            {
                joinedList .= item
                if (index < myFiles.Length)
                    joinedList .= "|"
            }            
        IniWrite(joinedList, A_ScriptDir . "\config.ini", "Settings", "myFiles")      
    }

    
    LoadFileList(*){
        ListboxContent := []
        oFilePicker.Delete()
        try
            ListboxContent := StrSplit(IniRead(A_ScriptDir . "\config.ini", "Settings", "MyFiles"), "|")
        if (ListboxContent.Length > 0)
            for index, item in ListboxContent
                if FileExist(item)
                    oFilePicker.Add([item])                              
    }

}


