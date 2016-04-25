Option Explicit
Call fn_Open_SAC()
Call GetCertificateSerialNoSAC10("1", "My Token")
Call fn_SAC_Close()
Call UnlockTokenSAMManage("Tokens by user","t","temp123#")

Call GetSACVersionFromRegistry()


Dim WshShell,regvalue,registryPath1
	Set WshShell = CreateObject("WScript.Shell")
	
	'10.0.43.0
	'HKLM\SOFTWARE\SafeNet\Authentication\SAC\RevisionID
	'Read the value of key from the registry
	regvalue = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\SafeNet\Authentication\SAC")

registryPath1="HKEY_LOCAL_MACHINE\SOFTWARE\SafeNet\Authentication\SAC\RevisionID"
'Here i am creating an instance of DotNet Factory computer object
Set objDFComputer=DotNetFactory("Microsoft.VisualBasic.Devices.Computer","Microsoft.VisualBasic")
'Here i am creating an instance of Registry Object
Set objRegistry=objDFComputer.Registry
'The below will retrieve the "Default" Property
value = objRegistry.GetValue(registryPath1,"","")
print objRegistry.GetValue(registryPath1,"","")

'The below retrieves the property of the regsitry
print objRegistry.GetValue(registryPath1,"RevisionID","")
'




'Call CopyUnlockCodeSAC10()
'
'Dim objCB
'Set objCB= CreateObject("Mercury.Clipboard")
'Dim responsecode:responsecode = objCB.GetText
'Call PasteUnlockCodeSAC10(responsecode ,"temp123#")

MsgBox "End Of SAC Functions"
 @@ hightlight id_;_656836_;_script infofile_;_ZIP::ssf8.xml_;_
