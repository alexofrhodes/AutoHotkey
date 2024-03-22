
^+F::Run,  "C:\Program Files (x86)\Everything\Everything.exe" 

^!+F::
{
	ClipSaved := Clipboard
	Clipboard := ""
	sleep 500
	Send ^c
	ClipWait, .5
	if !ErrorLevel
		Query := Clipboard
	; else
	; 	return
		; InputBox, Query, Google Search, , , 200, 100, , , , , %Query%
	Run,  "C:\Program Files (x86)\Everything\Everything.exe"
	if query !=
	{
		Sleep 500
		send ^F
		sleep 500
		send ^V
		; sleep 500
		; send {Tab}
	}else{
		Clipboard :=ClipSaved
	}	
	Query := ""
	return
}