#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^+G:: ;--Search Google

{
	RegRead, ProgID, HKEY_CURRENT_USER, Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice, Progid
	Browser := "iexplore.exe"
	if (ProgID = "ChromeHTML")
		Browser := "chrome.exe"
	if (ProgID = "FirefoxURL")
		Browser := "firefox.exe"
	
}

{
	Save_Clipboard := ClipboardAll
	Clipboard := ""
	Send ^c
	ClipWait, .5
	if !ErrorLevel
		Query := Clipboard
	else
		InputBox, Query, Google Search, , , 200, 100, , , , , %Query%
	if query !=
		Gosub Search
	
	Clipboard := Save_Clipboard
	Save_Clipboard := ""
	return
}

Search:
{
	
StringReplace, Query, Query, `r`n, %A_Space%, All 
StringReplace, Query, Query, %A_Space%, `%20, All
StringReplace, Query, Query, #, `%23, All
Query := Trim(Query)
if (Browser = "iexplore.exe")
{
	Found_IE := false
	For wb in ComObjCreate("Shell.Application").Windows 
		If InStr(wb.FullName, "iexplore.exe")
		{
			Found_IE := true
			break
		}
	if Found_IE
		wb.Navigate("http://google.com/search?hl=en&q=" Query, 2048) 
	else
	{
		wb := ComObjCreate("InternetExplorer.Application")
		wb.Visible := true
		wb.Navigate("http://google.com/search?hl=en&q=" Query) 
	}
}
else
{
	Run, %browser% http://www.google.com/search?hl=en&q=%Query% 
}
return
}


