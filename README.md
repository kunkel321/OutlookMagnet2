##OutlookMagnet2
This is an old AutoHotkey v1 project that has been remade to AHK v2. Many thanks to the folks who have helped (with both versions!).

The AutoHotkey forum thread is here:
https://www.autohotkey.com/boards/viewtopic.php?f=83&t=130938


The (new) GitHub repository is here:
[url]https://github.com/kunkel321/OutlookMagnet2[/url]
Newer versions will be added to GitHub.

Select some text, then activate the code.
[img]https://i.imgur.com/8zzWiT7.png[/img]

As you might guess, the tool is designed to work with MS Outlook. It won't work with Web/Cloud Outlook.  It only works with the "installed" version on your local machine.  If OL is not installed, you can still get the Preview MsgBox seen in the screenshot, but the tool wouldn't be very useful.   

[u]Notes:[/u]
-This script won't convert things like "day after tomorrow" into actual dates.  Luckily, OL has an excellent parsing feature and will convert the verbiage for you.  I tried to make the tool capture date verbiage that will get correctly parsed. 
-There are three supporting txt files that contain:
--a list of common US names
--a list of meeting/appointment locations
--a list of meeting types
you might guess that I work in the public education world, with special needs kids, so that meeting types, etc, are tailored toward that.  Change that information to suit your work situation.  Different names can be used for the files.  File names are assigned near the top of the ahk code.  If no files are present, sample files will be made.
-If you do have Outlook installed, but the install path is different than my own, change the path that is near the top of the ahk code. 
-This tool gets called via another script, so there is no embedded hotkey.  
-Hard-coded in the ahk code is time TIME verbiage like "before school," which gets converted to "7:15."  Change these to suit your needs. 
-On GitHub are 'OutlookMagnet2.ahk' and 'OutlookMagnet2.exe.'  Keep both!  The exe is not at compiled version of the ahk.  It is a renamed copy of 'AutoHotkey.exe' (v2).  Both (the ahk and the exe) have to be in the same folder.  This allows the ahk file to run as a portable app. 


