;duplicate urrent line if not empty else line above

^Down::
{
Clipboard := ""
Send {END}+{HOME 2}^c
ClipWait, 0
stringLen x, Clipboard

If (x=0) {
	;msgbox, , , "empty clip" %Clipboard%
	Send {UP}{END}+{HOME 2}^c
	ClipWait, 0
	stringLen x, Clipboard
	If (x>0) {
		Send {End}{Enter}{Home}^v
	}
}Else{
	;msgbox, , , "not empty clip" %Clipboard%
	Send {End}{Enter}{Home}^v
}	
return
}
