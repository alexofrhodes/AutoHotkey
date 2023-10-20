#IfWinActive ahk_exe explorer.exe

clip := ClipToVar()
Loop, parse, clip, `n, `r
{
	firstfile := A_LoopField
	SplitPath, firstfile, , folder, , filename_noext
	filefolder := folder "\" filename_noext "\"
	FileCreateDir, % filefolder
	FileMove, % firstfile, % filefolder
}

ClipToVar() {
  cliptemp := clipboardall ;backup
  clipboard = 
  send ^c
  clipwait, 1
  clip := clipboard
  clipboard := cliptemp    ;restore
  return clip
}