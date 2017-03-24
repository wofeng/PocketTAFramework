#SingleInstance force
#NoTrayIcon
#NoEnv
#include ../../Common/scriptheader.ahk
/*
Gosub, RunAllCases
ExitApp
*/

Gosub, OpenZoomClient
Gosub, LoginClient
Gosub, LogoutClient
ExitApp

OpenZoomClient:
paths := "C:\Users\" . A_UserName . "\AppData\Roaming\Zoom\bin\zoom.exe"
run, %paths%
WinWaitActive, ahk_class ZPFTEWndClass,,%@timeout%/1000
return

LoginClient:
xy1 := waitForImage("loginclient.png")
ClickElement(xy1[1] + 80 . "," . xy1[2] + 10,,"W")
xy2 := waitForImage("google.png")
ClickElement(xy2[1] - 150 . "," . xy2[2] - 60,,"W")
_InputVal("xxx@mailinator.com")
ClickElement("0, 60")
_InputVal("xxx")
Debug(xy2[1] . "," . xy2[2])
ClickElement(xy2[1] - 80 . "," . xy2[2] + 40,,"W")
WinWaitActive, ahk_class ZPPTMainFrmWndClass,,%@timeout%/1000
return

LogoutClient:
WinWaitActive, ahk_class ZPPTMainFrmWndClass,,%@timeout%/1000
Sleep, 500
Send, {Tab}{Enter}
Sleep, 500
Send, {UP}{Enter}
WinWaitClose, ahk_class ZPPTMainFrmWndClass,,%@timeout%/1000
return

_InputVal(val) 
{
 Send, ^a{DEL}  
 Send, %val%
}

#include ../../Common/scriptfooter.ahk
