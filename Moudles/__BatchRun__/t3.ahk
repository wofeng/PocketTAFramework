gosub, L1
gosub, L2
ExitApp

L1:
OutputDebug, x=1
example_func(false)         ; Stop execute the scripts below
OutputDebug, y=2
return

ErrorLabel:
exit

L2:
OutputDebug, z=3
return

example_func(param)
{
  if !param
   gosub, ErrorLabel

}