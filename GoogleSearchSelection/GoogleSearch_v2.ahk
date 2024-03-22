#Requires AutoHotkey v2.0

SendMode ("Input")

^+G:: ;--Search Google
{
    ProgID := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", "Progid")
    Browser := "iexplore.exe"
    if (ProgID = "ChromeHTML")
        Browser := "chrome.exe"
    if (ProgID = "FirefoxURL")
        Browser := "firefox.exe"
    
    Save_Clipboard := ClipboardAll()
    A_Clipboard := ""
    Send "^c"
    ClipWait()
    ErrorLevel := !ClipWait(0.5)
    if (!ErrorLevel)
        Query := A_Clipboard
    else
    {
        IB := InputBox("Google Search", "", "w200 h100")
        if (IB.Result = "OK")
            Query := IB.Value
        else
            return
    }

    Clipboard := Save_Clipboard
    Search(Browser, Query)
}

Search(Browser, Query)
{
    ; Replace special characters in the query for URL encoding
    Query := StrReplace(Query, "`r`n", " ") ; Replace newline characters with a space
    Query := StrReplace(Query, " ", "%20") 
    Query := StrReplace(Query, "#", "%23") ; Replace '#' with '%23' for URL encoding

    ; Trim the query
    Query := Trim(Query)
    
    if (Browser = "iexplore.exe")
    {
        Found_IE := false
        For wb in ComObject("Shell.Application").Windows 
            If InStr(wb.FullName, "iexplore.exe")
            {
                Found_IE := true
                break
            }
        if Found_IE
            wb.Navigate("http://google.com/search?hl=en&q=" Query, 2048) 
        else
        {
            wb := ComObject("InternetExplorer.Application")
            wb.Visible := true
            wb.Navigate("http://google.com/search?hl=en&q=" Query) 
        }
    }
    else
    {
        Run(Browser " http://www.google.com/search?hl=en&q=" Query)
    }
}
