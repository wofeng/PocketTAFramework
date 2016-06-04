#include  ../../Common/scriptheader.ahk

Gosub, RunAllCases
ExitApp

f1:
OutputDebug, --------------------f1-----------------------
@isError = 1
return


__f2:
OutputDebug,--------------------f2-----------------------
;goto, NextCase

return


f3:
OutputDebug,--------------------f3-----------------------
return

#include  ../../Common/scriptfooter.ahk
#Include  ../../Lib/WX11/WX11Lib.ahk