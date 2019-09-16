Option Explicit
Dim objShell, title, i

Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "C:\Program Files (x86)\KHCONF"
objShell.Run "khconfkh.exe"
title = "KHCONF KH"

Wscript.Sleep 2000

for i = 1 to 3
 objShell.AppActivate title
 WScript.Sleep 100
 objShell.SendKeys "{TAB}"
next

WScript.Sleep 100
objShell.SendKeys " "
Wscript.Sleep 500

for i = 1 to 5
 objShell.AppActivate title
 WScript.Sleep 100
 objShell.SendKeys "{UP}"
next

WScript.Quit
