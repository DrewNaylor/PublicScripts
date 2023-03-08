;
; MoveMouse Version: 2.1
; Language:       English (but all this matters for is the code and its comments)
; Platform:       AutoHotKey v1.1, Windows
; Author:         Drew Naylor <drewnaylor_apps AT outlook DOT com>
; Copyright:      2019, 2021, 2023 Drew Naylor
; License:        MIT License. Please view the file "LICENSE" in this repo for more details.
; Website:        https://drew-naylor.com
;
; Script Function:
;	Moves the mouse to a point on screen. I use this when I'm done working on my computer
;       and I want to go to sleep, as there's a place I usually move my pointer to before
;       turning my mouse and monitor off (it's part of the swish in the bottom-rightish part of the
;       default Windows 7 wallpaper when in widescreen/16:9, if you're wondering).
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



; First, we have to move the pointer. That's it.
; Be sure to modify these values according to your needs.
; You can use I think it's called AutoIt WindowSpy or
; something to know where a point on the screen is.
; Single monitor use at 1920x1080:
 MouseMove, 1365, 784
; Dual monitor use with a smaller 1366x768 monitor on the left side:
;MouseMove, 2645, 784
