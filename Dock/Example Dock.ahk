#NoEnv
#SingleInstance, force

Gui, +hwndGuihwnd
Gui, Font, s9				
Gui, Add, Button, gBtn, Dock to Top
Gui, Add, Button, gBtn, Dock to Bottom
Gui, Add, Button, gBtn, Dock to Right
Gui, Add, Button, gBtn, Dock to Left
Gui, Add, Button, gBtn ys, Dock to BL
Gui, Add, Button, gBtn, Dock to BR
Gui, Add, Button, gBtn, Dock to TL
Gui, Add, Button, gBtn, Dock to TR


Gui, Show, xCenter yCenter w300, class Dock Example  
Gui, +AlwaysOnTop

#Include, Class Dock.ahk
;The first argument is the host's hwnd and the second the client's hwnd
exDock := new Dock( Dock.HelperFunc.Run("notepad.exe"),Guihwnd)
exDock.Position("TL")
exDock.CloseCallback := Func("CloseCallback")

Return

CloseCallback(self)
{
	WinKill, % "ahk_id " self.hwnd.Client
	ExitApp
}

GuiClose:
	Gui, Destroy

Btn:
	exDock.Position(A_GuiControl)
Return

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; another way to get hwnd 
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; 		Run notepad.exe
; 		WinWait, ahk_class Notepad
; 		MainhWnd := WinExist("ahk_class Notepad")
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; or for active window: 
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; 		MainhWnd := WinExist("A")


; Gui, Add, Button, gAdd, Add dock
; Gui, Add, Button, gAdd, Add dock to Top
; Gui, Add, Button, gAdd, Add dock to Bottom
; Gui, Add, Button, gAdd, Add dock to Right
; Gui, Add, Button, gAdd, Add dock to Left