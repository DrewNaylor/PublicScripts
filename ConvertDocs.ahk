;
; ConvertDocs Version: 2023.03-1
; Language:       English
; Platform:       AutoHotKey 1.1, Windows 7, Adobe Acrobat DC, Microsoft PowerPoint 2010.
; Author:         Drew Naylor <drewnaylor_apps AT outlook DOT com>
; Copyright:      Copyright 2017, 2021, 2023 Drew Naylor
; License:        Licensed under the MIT License. See "LICENSE" in this repo for more details.
; Website:        https://drew-naylor.com
;
; Script Function:
;	Opens PDF files in Adobe Acrobat, deletes pages requested by the user,
;   opens the file in Microsoft PowerPoint, and prints off a black-and-white
;   version in PDF format. If you want to use this, you may need to tweak the
;   delays a bit so things don't go too fast for your computer if you're not using
;   an SSD and/or if you're using a different version of Windows and/or PowerPoint.
;   Your PDF printer may also be at a different part of the printer list in PowerPoint,
;   so you may have to change that as well. Please review this script's code so you know
;   what it does, as it's a bit complicated and requires user input at parts.
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; First, we have to open Adobe Acrobat Reader.
Run, Acrobat.exe *.pdf

; Next, we sleep for 2000 milliseconds/2 seconds to wait for Acrobat Reader.
sleep 2000

; Now we open the Tools area.
SendInput !em

; We have to delay pressing Tab for Acrobat to catch up.
sleep 200
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {ENTER}

; Ask the user which pages they want to delete.
InputBox, pagesDeleteThese, Which page(s) would you like to delete?, Please type the page(s) you wish to delete into the textbox below separated with a comma. A range of numbers can also be entered with the hyphen.
If ErrorLevel
	Exit
else

; Before moving the mouse, save the current position for the user.
MouseGetPos, mousevarx, mousevary

; After opening the PDF editing area, move the mouse into
; the page range field after a short delay.
sleep 200
MouseMove, 302, 157

; Now we click the mouse.
MouseClick, left

; We don't want anything here yet.
SendInput {Backspace}

; We can now input what the user typed earlier regarding the
; pages they want to delete.
SendInput %pagesDeleteThese%
sleep 100
SendInput {ENTER}

; Now move the mouse back to where it was.
MouseMove, %mousevarx%, %mousevary%

; Ask the user if they are sure that they want to delete the
; pages.
MsgBox, 4, Delete selected page(s)?, Are you sure you want to delete the selected page(s)? Once your selected pages are deleted, Acrobat Reader will change the page numbers for all other pages. This script will exit if you click "No."
IfMsgBox No
	Exit
else
	SendInput {DELETE}
	SendInput {ENTER}

; After page deletion is finished, we start the PowerPoint export.
sleep 500
SendInput !ftt
sleep 1000
Send {HOME}
Send `TEMP `
Send {ENTER}
sleep 8000
Send #{LEFT}
sleep 200

; Ask the user how many pages the presentation has now.
InputBox, pagesTotalNumber, How many page(s) are in this presentation?, Please type the total number of page(s) this presentation contains currently. You can look at the bottom-left corner of the PowerPoint window to get the number.
If ErrorLevel
	Exit
else


; Now that PowerPoint is open, we can change the color of everything.
sleep 2000
Loop, %pagesTotalNumber%
{
	sleep 100
	; Set color mode to black and white.
	SendInput !w
	sleep 100
	SendInput b
	sleep 100
	; Select all things on this page and change their color to
	; Black with Grayscale Fill.
	SendInput ^a
	sleep 100
	SendInput !b
	sleep 100
	SendInput l
	sleep 100
	; Go to next page.
	SendInput {PgDn}
}


; Once the Loop is finished, we want to print out the
; black-and-white document as a PDF.
sleep 2000
; Open the Print option in the File Menu.
SendInput !f
SendInput p
sleep 500
; Open the printer dropdown box.
SendInput {TAB 3}
sleep 200
; Select the Adobe PDF printer.
SendInput {ENTER}
sleep 200
SendInput {HOME}
SendInput {ENTER}
sleep 500
; Change document color to "Pure Black and White."
SendInput {END}
SendInput {Up 2}
sleep 200
SendInput {END}
SendInput {ENTER}
sleep 500
; Open the "Print" window.
SendInput {SHIFT down}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {TAB}
SendInput {SHIFT up}
sleep 200
; Click the "Print" button.
SendInput {ENTER}


; Tell the user that the script has finished running.
MsgBox, 0, Finished, This script has finished running. Please type the filename you want to use for this document. Once you're finished, make sure to delete any PDF files in the "original_pdf" folder because this script won't run properly if there are two or more PDF documents in that folder at one time.
