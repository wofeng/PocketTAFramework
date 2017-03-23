#SingleInstance force
#NoTrayIcon
#NoEnv
#include ../../Common/scriptheader.ahk

Gosub, RunAllCases
ExitApp
/*
Gosub, initVars
Gosub, SignInPage
PrtScreen("test.png")
Gosub, SignIn
Gosub, SignOut
ExitApp
*/
initVars:
global @brexe = "firefox.exe"
global @site_url = "https://dev.zoom.us"
return

SignInPage:
OpenSiteWithBrowser()
WaitForImage("login.png")
xy:= WaitForImage("logo.png")
ClickElement(xy[1] + 1200 . "," . xy[2] + 8,,"W")  ;sig in link
return

SignIn:
WaitForImage("facebook.png")
ClickElement("-500, 170")
Send xxx@mailinator.com
Send, {Tab} xxx {Enter}
return

SignOut:
WaitForImage("profile.png")
xy:= WaitForImage("logo.png")
ClickElement(xy[1] + 1850 . "," . xy[2] + 8,,"W") 
return
#include ../../Common/scriptfooter.ahk
