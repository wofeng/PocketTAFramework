initialGlobalVars:
; parameters starts with "@" is looked as global variable, start with "@_" is looked as inner global variable
@error_msg =
@cfgobj 	:= GetAllKVFromFile("../../common/sysconfig.ini")
@timeout 	:= @cfgobj.timeout
@site_url 	:= RegExReplace(@cfgobj.site_url, "/","",ps,1, 0) ; remove "/" at the last of site_url
@image_dir 	:= @cfgobj.image_dir
@capture_screen_path 	:= @cfgobj.capture_screen_path
@error_image_path		:= @cfgobj.error_image_path
@brexe 		:= @cfgobj.brexe
@reg_global	:= @cfgobj.reg_global
@curr_acc	:= @cfgobj.curr_acc
;TODO log object
@isDebug 	:= @cfgobj.isDebug
; parameters or label starts with "_" is looked as private parameter or methods which invoked by framework
@case_dependency := @cfgobj.case_dependency
@caseLevel := @cfgobj.caseLevel
@_has_initialGlobalVars = true
return


LoadAllCasesFromIniFiles:
; @_ini_caselevel_cases = {k1:"label1", k2:"label2", k3: "label3"}
; @_inis_caselevel_cases = [{k1:"label1", k2:"label2", k3: "label3"}, ...,{kn:"label1n", kn:"label2n", kn: "label3n"}]
@_ini_files  := ""
@_inis_caselevel_cases := []
@_modulenum := 0
@_allcases := 0
Debug("<========= [" . A_ThisLabel . "] =========>")
IfNotExist, common\*.ini
{
	@error_msg := "Err: Fail to read ini files from the 'common' directory, there are no any ini files."
	Gosub, ErrorLabel
}
	
Loop, common\*.ini															; get all ini files
	;@_ini_files:= @_ini_files . "common\" . A_LoopFileName . "`n"
	@_ini_files .= "common\" . A_LoopFileName . "`n"

Sort, @_ini_files, CL														; Arranges a variable's contents in alphabetical order.

Loop, parse, @_ini_files, `n												; get every ini file
{
	if (Trim(A_LoopField))
	{
		Debug("[" . A_LoopField . "] is loading")
		@_ini_caselevel_cases := GetSectionKVFromFile(A_LoopField, @caseLevel)
		if (!!@_ini_caselevel_cases)
		{
			@_inis_caselevel_cases.Insert(@_ini_caselevel_cases)
			Debug("Done")
		}
	}
	
}	


@_modulenum := @_inis_caselevel_cases.MaxIndex()
Loop % @_modulenum
{
	for @_key, @_val in @_inis_caselevel_cases[A_Index]
	{
		++@_allcases
	}
}

@_isCasesReady := true
return

RunAllCases:
@_casenum := 0 
if (!@_isCasesReady)
{
	Gosub, LoadAllCasesFromIniFiles
}

if (@_inis_caselevel_cases[1])
{
  @_ini_caselevel_cases := @_inis_caselevel_cases[1]
  Gosub RunCasesOneIni
}
return

RunCasesOneIni: 
Debug("<========= [" . A_ThisLabel . "] =========>")
@_obj := @_ini_caselevel_cases.clone()
for @_key, @_val in @_obj
{
	if(IsLabel(@_val))
	{
		++@_casenum
		@isError := SubStr(@_val, 1, 1) != "_" ? 0 : @isError 						; Reset isError for parent case 
		if(@isError and SubStr(@_val, 1, 1) == "_" and @case_dependency) 	; Skip sub-label execution if parent case error occurs and case_dependency = 1 in sysconfig.ini
		{
			Debug("Skip [" . @_key . " :" . @_val . "] execution")
			@_ini_caselevel_cases.remove(@_key) 									; Remove the skip case, use new loop to go through each case	
			Goto RunCasesOneIni 													; similiar to recurisive function
		}

		Debug("Start executing [" . @_key . " : " . @_val . "]")
		Gosub, %@_val%
		Debug("Done [" . @_key . " : " . @_val . "]")
	}
	else
	{
		Debug("Err [" . @_key . " : " . @_val . "] has NOT been implemented in the script.")
	}
	@_ini_caselevel_cases.remove(@_key)				; Remove the finished, unimplemented case, use new loop to go through each case
}
@_inis_caselevel_cases.remove(1)					; Remove finished cases in the ini file and make next cases object in next ini file as current object to be handled
Goto RunAllCases
return

NextCase:
Debug("<========= [" . A_ThisLabel . "] =========>")
@_ini_caselevel_cases.remove(@_key)												; remove the current run case
Debug("Stop executing [" . @_key . " : " .  @_val . "] case")
Gosub, RunCasesOneIni
return


ErrorLabel:
@isError=1
;PrtScreen(error_image_path . "/err_" . A_Now . ".png")
OutputDebug, "<========= [" . A_ThisLabel . "] =========>"
OutputDebug % @error_msg
Gosub, NextCase
return

WarringLabel:
OutputDebug, "<========= [" . A_ThisLabel . "] =========>"
;PrtScreen(error_image_path . "/war_" . A_Now . ".png")
OutputDebug % @error_msg
return

Pause::
Pause
Suspend
ToolTip, Pause/Resume..........
Sleep, 1000
ToolTip
return

PrintScreen::
PrtScreen(@capture_screen_path . "/" . A_Now . ".png")
return

Esc::
ToolTip, Stop the script execution 
Sleep, 1000
ToolTip
ExitApp
