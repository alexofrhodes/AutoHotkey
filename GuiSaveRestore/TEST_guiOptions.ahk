#SingleInstance force  
#NoENV              
SetBatchLines -1    
ListLines Off       

#Include, guiSaveRestore.ahk

BuildGui:
    Gui, Add, Button, x+5 h21 w60 , Save    
    Gui, Add, Button, x+5 h21 w60 vDELETEBUTTON, Delete 

    Gui, Add, Text, xm
    Gui, Add, Text, xm section , Edit1
    Gui, Add, Edit, ys w200 vfrmJOBTITLE, 

    Gui, Add, Text, xs section right vDropDown1, DropDown
    Gui, Add, DropDownList, x+5 vfrmCONVERT, First||Second|Third

    Gui, Add, Checkbox, vfrmRECURSE section xm vCheck1, Checkbox1
    Gui, Add, Checkbox, vfrmAUTODUMP xs vCheck2, Checkbox2

    Gui, Add, StatusBar

    guiName := "ThisGui"
    ;show but off screen to avoid flicker, guiRestore will move the gui in correct place
    gui, show, x10000 y10000,%guiName% 

    guiRestore(,guiName)
return

GuiClose:
GuiEscape:
    guiSave(,guiName) ;only for controls with vID
    gui, hide