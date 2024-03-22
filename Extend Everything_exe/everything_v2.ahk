#Requires Autohotkey v2.0


^+F::Run("C:\Program Files (x86)\Everything\Everything.exe")

^!+F::
{
	ClipSaved := A_Clipboard
	A_Clipboard := ""
	Sleep(500)
	Send("^c")
	Errorlevel := !ClipWait(.5)
	if !ErrorLevel
		Query := A_Clipboard
	if (query != "")
	{
		Run("C:\Program Files (x86)\Everything\Everything.exe")		
		Sleep(500)
		Send("^F")
		Sleep(500)
		Send("^V")
	}else{
		A_Clipboard :=ClipSaved
	}	
	Query := ""
	return
}