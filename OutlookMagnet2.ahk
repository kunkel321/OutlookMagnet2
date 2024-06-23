#SingleInstance
#Requires AutoHotkey v2+
;#################################
; App: OUTLOOK MAGNET for AHK v2.
; By: kunkel321 (with much help from others)
; Date: 6-21-2024
;#################################
; Requires that MS Outlook app is installed on local machine.
; (Won't work with Web/Cloud Outlook.)
; Also uses three different text files, see below.  If not present, sample files will be created.
; There's no hotkey.  This script should get 'Run' from another script, link, etc.  
; Select some text (which contains new appt verbiage) then activate.

;###### A FEW TEST STRINGS ########### 
; There is a conference for Jimmy Bob after school next Wed in the Resource Room.
; Let's do Monday in 2 weeks in Rm 101.
; The Feb 20th transition meeting for Jenny is at noon in the Commons.
; Nov 17 before school for Billy's thing.
; Day after tomorrow in the morning in your office.

SetWorkingDir(A_ScriptDir)
PathToOutlook := "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
LibraryOfLocations :=  "LocationList.txt"
LibraryOfNames := "2000MostCommonUSNames.txt"
LibraryOfMeetingTypes := "MeetingTypes.txt"

; ============
; Incase the files don't exist... create them and put a couple sample library items.
If not FileExist(LibraryOfLocations)
	FileAppend("Resource Room`r`nRm|Room\s(\d){,3}`r`nCommons", LibraryOfLocations)
If not FileExist(LibraryOfNames)
	FileAppend("Jimmy Bob`r`nBilly`r`nJenny", LibraryOfNames)
If not FileExist(LibraryOfMeetingTypes)
	FileAppend("staffing`r`nconference`r`ntransition", LibraryOfMeetingTypes)
; ============

ClipboardOld := ClipboardAll() ; Save clipboard to restore below.
A_Clipboard := ""  ; Must start off blank for detection to work.
Send("^c")
Errorlevel := !ClipWait(1)
if ErrorLevel  ; ClipWait timed out.
{	MsgBox("No text selected(?)`n`n   (Script exiting.)")
    ExitApp()
}

;######## GET LOCATION #####################
myLocList := Fileread(LibraryOfLocations)
myLocList := StrReplace(myLocList, "`r`n", "|")
reLocation := StrReplace(myLocList, myLocList, "(" myLocList ")")
reLocPre := "in\s(the\s)?"  ; don't capture
;MsgBox '---myLocList---`n' myLocList '`n`n---reLocation---`n' reLocation '`n`n---reLocPre---`n' reLocPre

If RegExMatch(A_Clipboard, "i)" reLocPre "(\K)" reLocation, &myLocation)
	MyLocation := myLocation[0]
Else If RegExMatch(A_Clipboard, "i)" reLocation, &myLocation) 
	MyLocation := myLocation[0]
Else
	MyLocation := ""

;######## GET TIME #########################
reTime := "\b(10|11|12|[0-9])(:[0-5][0-9])(am|pm|a|p)?|(10|11|12|[0-9])(am|pm|a|p)|((before\s|after\s)school|noon|in the morning)\b"
reTimPre := "(at|@|time:)"

If RegExMatch(A_Clipboard, "i)" reTimPre "\s\K" reTime, &myTime)
    myTime := myTime[0]
Else If RegExMatch(A_Clipboard, "i)" reTime, &myTime)
    myTime := myTime[0]
else
	myTime := ""

If (myTime = "after school")
    myTime := "2:15 pm" ; <--- These get changed from time to time.  Put your own items, based on your needs. 
Else If (myTime = "before school" ) || (myTime = "in the morning")
    myTime := "7:00 am"

;msgbox 'time: ' myTime
;######### GET DATE ######################## 
; ---- okay if captured alone ----
reNumDate := "\b(([0-9]|10|11|12)[-.\/]([0-2]?[0-9])[-.\/](20)?[0-9][0-9])"
reWDay := "\b(((Mon|Tue(s)?|Wed(nes|s)?|Thu(rs|r)?|Fri|Sat(ur)?|Sun))(day|d)?|(day\safter\s)?tomorrow)\b"
reMonth := "\b((Jan|Feb)(r?u(ary)?)?|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|(Sep(tem)?|Oct(o)?|Nov|Dec)(em)?(ber)?)\b"
re1st31st := "\b(?<!:)([0-2]?[0-9]|30|31)(st|nd|rd|th)?(?!:)\b"
; ---- only to be captured if paired with above string ----
re2WeeksPre := "(((2|two)\sweeks\s)(from\s)?|((week\s)?after next\s(on\s)?))"
re2WeeksSuf := "(((in\s)?(2|two) weeks)|((week )?after next))" 
re1WeekPre := "((next\s?(week\s?)?(on\s)?)|(week\s(from|after)))"
re1WeekSuf := "(next( week)?)"

;A_Clipboard := StrReplace(A_Clipboard, myTime)        ; Removes myTime, so digits are not interpreted as dates.
If RegExMatch(A_Clipboard, "i)(" reNumDate ")", &myDate) ; e.g. 01-01-2016
	myDate := myDate[0]
else If RegExMatch(A_Clipboard, "i)(" reMonth "\s" re1st31st "|" re1st31st "\sof\s" reMonth  ")", &myDate) ; e.g. Jan 1st | 1st of Jan
	myDate := myDate[0]
else If RegExMatch(A_Clipboard, "i)(" re2WeeksPre "\s" reWDay "|"  reWDay "\s" re2WeeksSuf ")", &myDate) ; e.g. Mon in two weeks
	myDate := myDate[0]
else If RegExMatch(A_Clipboard, "i)(" re1WeekPre "?\s" reWDay "|"  reWDay "\s" re1WeekSuf "?)", &myDate) ; e.g. next Mon
	myDate := myDate[0]
Else
	myDate := ""

;######### GET SUBJECT #########################
; A_Clipboard := StrReplace(A_Clipboard, myLocation)    ; removes myLocation so that it's not mistaken as the Subject.
; A_Clipboard := StrReplace(A_Clipboard, myDate)        ; removes myDate

myNameList := Fileread(LibraryOfNames)
myNameList := StrReplace(myNameList, "`r`n", "|")
reNames := StrReplace(myNameList, myNameList, "\b(" myNameList ")\b")

myTypeList := Fileread(LibraryOfMeetingTypes)
myTypeList := StrReplace(myTypeList, "`r`n", "|")
reTypes := StrReplace(myTypeList, myTypeList, "(" myTypeList ")")
;MsgBox '---Names---`n' reNames
;MsgBox '---Types---`n' reTypes

If RegExMatch(A_Clipboard, reNames, &myName)
	myName := myName[0]
Else 
	myName := ""

If RegExMatch(A_Clipboard, "i)" reTypes, &myType)
	myType := myType[0]
else
	myType := ""

if (myName) AND (myType)
{    mySubject := myName . " - " . myType
    ; msgBox myName AND myType box `n myName = %myName%, mytype = %myType% `n mySubject is %mySubject%
}
else
{	mySubject := myName . "" . myType
	; msgBox not myName AND myType  `n myName = %myName%, mytype = %myType% `n mySubject is %mySubject%
}

;#################### SUMMARIZE #########################
decision := MsgBox(
	"`nDate:`t`t" myDate 
	"`nTime:`t`t" myTime 
	"`nLocation:`t`t" myLocation 
	"`nSubject:`t`t" mySubject
	"`n`nSend to Outlook?"
	,"PREVIEW", 4096 " OC icon?")
If decision = "Cancel"
	EscFunction() 

If not FileExist(PathToOutlook)
{	MsgBox("This message means that MS Outlook was not found at the expected location on your local disc.  Please see the variable `"PathToOutlook`" near the top of the code.`n`nNow exiting app.")
	ExitApp()
}

;############### OUTLOOK PART ########################
Run("`"" PathToOutlook "`" /c ipm.appointment")
If WinWaitActive("Untitled - Appointment", , 10)
{	SendInput("!u" mySubject) ; Subject  now called 'title': Alt+l
	SendInput("!t" myDate) ; Shortcut keys to jump to different fields. 
	SendInput("!t{Tab}{Tab}" myTime)
	SendInput("!i" myLocation)
	Sleep 1500
	If mySubject = "" ; Now go to first blank field, and leave cursor there. 
		SendInput("!u")
	Else If myDate = ""
		SendInput("!t")
	Else If myTime = ""
		SendInput("!t{Tab}{Tab}")
	Else If myLocation = ""
		SendInput("!i")
}
else
	MsgBox("The new appointment dialog never appeared.`n`nNow restoring previous clipboard and exiting.")

EscFunction() 

Esc::
EscFunction(*) 
{ 	A_Clipboard := ClipboardOld  ; Restore previous contents of clipboard. 
	ExitApp()
}

/*
Navigation of Outlook Calendar Item Form:   (updated 10-10-2018)
/c ipm.appointment
Subject !u
Location !i
Start date !t
Start time !t{Tab}{Tab}
End Date !d
End time !d{Tab}{Tab}
Notes !d{Tab}{Tab}{Tab}{Tab}  Developer Tab must be hidden.
Save and Close !s
*/


