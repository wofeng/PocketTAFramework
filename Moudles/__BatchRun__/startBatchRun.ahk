#include  ../../common/scriptheader.ahk

Gosub, BatchRunCases
ExitApp

BatchRunCases:
@_moudles_keys := GetSectionKVFromFile("../../Common/sysconfig.ini", "Modules") 
for @_k, @_v in @_moudles_keys
{
	try
	{
		RunWait, TestSuite.ahk, ..\%@_v%
	}
	catch e
	{
		OutputDebug, Does the [%@_v%] exist in [Modules] directory?
	}

}
return
#include  ../../common/scriptfooter.ahk