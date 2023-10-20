# The ahk shell menu

Displays a popup menu when you 
extended-right-click 
in file explorer.   
(remember to select ALL target files with left click before calling the popup)  

each 1st level folder in (A_ScriptDir \ Scripts \ ) serves as a submenu (category) and each ahk or exe in those folders (recursive) becomes a submenu button

while each 1st level ahk or exe becomes a menu button

**eg folder/file structure**
>- Scripts\  
>     - Txt Scripts\  
>        - Action Set 1\  
>            - Cat1\
>                - a.ahk
>            - b.exe
>     - Zip Scripts\  
>        - c.ahk  
>    - d.ahk  
>    - e.exe  


**result**

>- Txt Scripts **>**  
>    - a.ahk  
>    - b.exe
>- Zip Scripts **>**  
>    - c.ahk  
>- d.ahk
>- e.exe

**One way for the ahk scripts loop the selected files is**  


```ahk
GoSub, Main

Main() {
clip := ClipToVar()
Loop, parse, clip, `n, `r
{
  targetFile := A_LoopField
  ;YOUR CODE HERE
}

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
```