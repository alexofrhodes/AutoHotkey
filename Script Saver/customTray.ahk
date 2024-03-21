

SetupTray(){
    I_Icon := A_ScriptDir "\icons\snippet.png" 
    iF FileExist(I_Icon)
        TraySetIcon(I_Icon)

    Tray := A_TrayMenu
    Tray.Add()  ; Creates a separator line.
    Tray.Add("Left CTRL x2", DoNothing)
	Tray.SetIcon("Left CTRL x2", "icons\hotkey.ico")
	tray.Disable("Left CTRL x2")
	Tray.Add()
	
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

 
FollowLink(ControlName, info, *){
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