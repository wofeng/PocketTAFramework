#SingleInstance force
#NoTrayIcon
#NoEnv
SetWorkingDir %A_ScriptDir%
/*
 03/23/2017: 
  - Fix fail to modify case description issue
  - GuiSize issue
*/
; Allow the user to maximize or drag-resize the window:
Gui, +Resize 
Gui, Add, Button, x640 y10 vSwbtn, Switch View
Gui, Add, ListView, xm r5 w710 h200 Grid -Multi +Report vMyListView gMyListView, Name|Folder|Size (KB)|Description
LV_ModifyCol(1, 150)
LV_ModifyCol(2, 100) 
LV_ModifyCol(3, 80) 
LV_ModifyCol(4, 350) 
ImageListID1 := IL_Create(10)			 ; Create an ImageList so that the ListView can display some icons:
ImageListID2 := IL_Create(10, 10, true)  ; A list of large icons to go with the small ones.
LV_SetImageList(ImageListID1)			; Attach the ImageLists to the ListView so that it can later display the icons:
LV_SetImageList(ImageListID2)
Gui, Add, StatusBar, vMyStatusBar

LoadMyListView()
SetStatusBar()
Gui, Show,, Testing Assistant
return

GuiSize:
   If A_EventInfo = 1
      Return
	  
   GuiControl,1:Move,Swbtn, % "X" . (A_GuiWidth - 90) 
   GuiControl,1:Move,MyListView, % "X" . (5) . " Y" . (35) . " W" . (A_GuiWidth - 10) . " H" . (A_GuiHeight - 55)
;   WinSet,Redraw
   LV_ModifyCol( 2, "Sort" )
Return

SetStatusBar(){
	SB_SetParts(650, 80)
	SB_SetText("author: Feng Wu", 1)
	SB_SetText("v0.2", 2)
}

OpenFile:
sel_rownum := LV_GetNext("Focused") 
LV_GetText(file_name, sel_rownum, 1)
LV_GetText(dir_name, sel_rownum, 2)
Run, %dir_name%/%file_name%,,UseErrorLevel 
if ErrorLevel
	MsgBox Could not perform requested action on [%dir_name%\%file_name%].
Gui,1: -AlwaysOnTop
return
  
FileSettings:		
Gui,1: +Disabled
sel_rownum := LV_GetNext("Focused") 
LV_GetText(dir_name, sel_rownum, 2)
FileRead, desc, %dir_name%/description.txt
; Change Settings GUI
Gui,2: +Resize +AlwaysOnTop +OwnDialogs
Gui,2: Add, Edit, x22 y30 w440 h300 vMyEditConfig, %desc%
Gui,2: Add, Button, x182 y350 w120 h50 vMySaveConfig gSaveConfig, Save
Gui,2: Show, w492 h415, Change Settings
return

2GuiSize:
Anchor("MyEditConfig", "wh")
Anchor("MySaveConfig", "xy")
return 

SaveConfig:
Gui,2: Submit
FileCopy,%dir_name%/description.txt, %dir_name%/description.bak, 1   
FileDelete, %dir_name%/description.txt
; No submit then GuiControlGet, MyEditConfig
FileAppend, %MyEditConfig%, %dir_name%/description.txt
Gui,2:Destroy

; Update the row
Gui,1: +AlwaysOnTop -Disabled 
Gui, 1: Default
sel_rownum := LV_GetNext("Focused") 
LV_Modify(sel_rownum, "Col4", MyEditConfig)
return


MyListView:
if A_GuiEvent = DoubleClick 
{
    LV_GetText(file_name, A_EventInfo, 1)
    LV_GetText(dir_name, A_EventInfo, 2) 
	Run, %dir_name%/%file_name%,,UseErrorLevel 
	if ErrorLevel
		MsgBox Could not perform requested action on [%dir_name%\%file_name%].
}
return

LoadMyListView() {
	
	GuiControl, -Redraw, MyListView  ; Improve performance by disabling redrawing during load.
	arr := []

	Loop, *.exe,0, 1 
	{
		if(!RegExMatch(A_LoopFileDir, "\\")){		; The .Exe file should be at the root of sub-folder
			icon_num := _GetFileIconNum(A_LoopFileLongPath)
			RegexMatch(A_LoopFileFullPath, "(.+)\\(.+)", subpat)  ; filename = subpat2, dirname = subpat1
			FileEncoding, UTF-8
			FileRead, desc, %A_LoopFileDir%\description.txt
			LV_Add("Icon" . icon_num, subpat2, subpat1, A_LoopFileSizeKB, desc)
		}
	}
	GuiControl, +Redraw, MyListView  ; Improve performance by disabling redrawing during load.
}

_GetFileIconNum(file_full_name) {
	global ImageListID1, ImageListID2
	; Calculate buffer size required for SHFILEINFO structure.
	sfi_size := A_PtrSize + 8 + (A_IsUnicode ? 680 : 340)
	VarSetCapacity(sfi, sfi_size)
	
	if not DllCall("Shell32\SHGetFileInfo" . (A_IsUnicode ? "W":"A"), "str", file_full_name, "uint", 0, "ptr", &sfi, "uint", sfi_size, "uint", 0x101)  
	{
		icon_num = 99999 			; 0x101 is SHGFI_ICON+SHGFI_SMALLICON
	}
	else 
	{
		hIcon := NumGet(sfi, 0)		; Extract the hIcon member from the structure:
		icon_num := DllCall("ImageList_ReplaceIcon", "ptr", ImageListID1, "int", -1, "ptr", hIcon) + 1  ; Add the HICON directly to the small-icon and large-icon lists.
																									; Below uses +1 to convert the returned index from zero-based to one-based:
		DllCall("ImageList_ReplaceIcon", "ptr", ImageListID2, "int", -1, "ptr", hIcon)
		DllCall("DestroyIcon", "ptr", hIcon)													  ; Now that it's been copied into the ImageLists, the original should be destroyed:
		}
	return icon_num
}

ButtonSwitchView:
if not IconView
	GuiControl, +Icon, MyListView  	  ; Switch back to details view.
else
	GuiControl, +Report, MyListView    ; Switch to icon view.
IconView := not IconView              ; Invert in preparation for next time.
return

GuiContextMenu:  				
; Create a popup menu to be used as the context menu:
Menu, MyContextMenu, Add, Open, OpenFile
Menu, MyContextMenu, Add, Settings, FileSettings
Menu, MyContextMenu, Default, Open  ; Make "Open" a bold font to indicate that double-click does the same thing.
; Launched in response to a right-click or press of the Apps key.
if A_GuiControl <> MyListView  		 ; Display the menu only for clicks inside the ListView.
    return
Menu, MyContextMenu, Show, %A_GuiX%, %A_GuiY% 	; Show the menu at the provided coordinates, A_GuiX and A_GuiY.  These should be used
												; because they provide correct coordinates even if the user pressed the Apps key:
sel_rownum := LV_GetNext("Focused") 
return


GuiClose:  
ExitApp

2GuiClose:
2GuiEscape:
Gui,1: -AlwaysOnTop
Gui,2: -AlwaysOnTop
MsgBox, 4,, Do you want to save changes?
IfMsgBox Yes
{
   Gosub, SaveConfig
}   
Gui,2:Destroy
Gui,1: -Disabled +AlwaysOnTop
return