;Smarter unzip sketch
;by nod5  181114
 
;note: only a sketch/prototype so far
;todo: 
;- add error checks/handling: 
;   1 unzip errors (paths to long, or whatever)
;   2 file/folder move errors: 
;     - script currently silently avoids overwrite
;     - need to add code to prompt user: overwrite yes/no? and react accordingly
;- must test how reliable the unzip function performs if the zip file is huge, etcetera
;- notify user while processing (window, tooltip, etcetera)
;- prevent repeat commands while processing
;- replace crude F5 refresh with something better, maybe also reselect zip file after refresh
 
#IfWinActive ahk_exe explorer.exe

;#W::
 
;get all selected files in Explorer and verify that it is .zip
clip := ClipToVar()

;MsgBox % clip
sleep, 200

Loop, parse, clip, `n, `r
{
  firstfile := A_LoopField

 ; MsgBox % firstfile
sleep, 200

if !FileExist(firstfile) or (SubStr(firstfile,-3) != ".zip")
  return
 
;unzip to temporary folder
SplitPath, firstfile, , folder, , zipname_noext
tempfolder := folder "\temp_" A_Now
FileCreateDir, % tempfolder 
Unzip(firstfile, tempfolder)
 
;count unzipped files/folders recursively
Loop, Files, % tempfolder "\*.*", FDR
{
  allcount := A_Index
  onlyitem := A_LoopFilePath
}
 
;count unzipped files/folders at first level only
Loop, Files, % tempfolder "\*.*", FD
{
  firstlevelcount := A_Index
  firstlevelitem  := A_LoopFilePath
}
 
;case1: only one file/folder in whole zip
if (allcount = 1)
{
  if InStr( FileExist(onlyitem), "D")
  {
    SplitPath, onlyitem, onlyfoldername
    FileMoveDir, % onlyitem, % folder "\" onlyfoldername
  }
  else
    FileMoveDir % tempfolder, % folder "\" zipname_noext ;;;
    FileMove   , % onlyitem, % folder
}
 
;case2: only one folder (and no files) at the first level in zip
else if (firstlevelcount = 1) and InStr(FileExist(firstlevelitem), "D")
{
  SplitPath, firstlevelitem, firstlevelfoldername
  FileMoveDir, % firstlevelitem, % folder "\" firstlevelfoldername
}
 
;case3: multiple files/folders at the first level in zip
else
{
  FileMoveDir % tempfolder, % folder "\" zipname_noext
}
 
;cleanup temp folder
FileRemoveDir, % tempfolder, 1
 
}

;refresh Explorer to show results
If WinActive("ahk_exe Explorer.exe")
  Send {F5}
return
 
 
;function: copy selection to clipboard to var
ClipToVar() {
  cliptemp := clipboardall ;backup
  clipboard = 
  send ^c
  clipwait, 1
  clip := clipboard
  clipboard := cliptemp    ;restore
  return clip
}
 
 
;function: unzip files to already existing folder
;zip file can have subfolders
Unzip(zipfile, folder)
{
  psh := ComObjCreate("Shell.Application")
  psh.Namespace(folder).CopyHere( psh.Namespace(zipfile).items, 4|16)
}
