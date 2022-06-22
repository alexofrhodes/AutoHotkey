When you create vbe window menus they need event handler which breaks if code execution is stopped.  
(i have writen code to format the codepane content or inset snippets etc for which i wanted to have a menu inside the vbe window,  
have a look at my excel menu builder on github.  
This problem doesn't occur on excel ribbon ui elements because they don't depend on event handler.  
Ugly solution would be to have a button in excel ribbon to recreate the event handler.  
But it is sooo easy to create popup menus with AutoHotkey to run your excel macros!  

replace workbookname variable with your own and make a list of your macros in the vba.menu file.  
can be modified to hold macros from different workbooks if interested

<img src="https://user-images.githubusercontent.com/62287665/172789520-b56c74f3-b5e8-4e83-9401-a54d3db82e8c.jpg" width="400"> <img src="https://user-images.githubusercontent.com/62287665/172789524-cf018151-86bf-4a24-8f01-bf20ec09b6c7.jpg" width="400">
<img src="https://user-images.githubusercontent.com/62287665/174991681-3a9ae151-0858-490e-beee-7c9b1bce5e99.jpg" width="400" height="">

