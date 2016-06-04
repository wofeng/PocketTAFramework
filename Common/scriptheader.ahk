#SingleInstance force
#NoTrayIcon
#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Pixel, Relative
; initial global variable and hotkeys
if (!@_has_initialGlobalVars)
{
	
	Gosub, initialGlobalVars
}

