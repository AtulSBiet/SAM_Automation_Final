Option Explicit
Call SetSACMajorVersionInUFT()'Setting SACMajorVersion Environment Value in UFT

Dim Iterator
For Iterator = 1 To 10
	Call fn_Open_SAC()
	Call fn_Get_Certificate_SerialNo_from_SAC(1)
	Call fn_SAC_Close()
	Call UnlockTokenSAMManage("Tokens by user","t","temp123#")
Next


 @@ hightlight id_;_656836_;_script infofile_;_ZIP::ssf8.xml_;_
