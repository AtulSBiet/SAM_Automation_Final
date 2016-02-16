Option Explicit

Dim result
result = OpenIEWithURL(samManageUrl)

If result = "PASS" Then
	'loginSamManage
	result = LoginSamManage(samUserName,samPassword)
Else
	Call fn_ExecutionLog("LoginSamManage", "Not Started")
	ExitAction'TODO:OR should we use ExitTest
End If

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call DisableTokenSAMManage("Tokens by user",enrollmentUserName)
Call EnableTokenSAMManage("Tokens by user",enrollmentUserName)
Call CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName)
Call RevokeTokenSAMManage("Damaged")'Revocation Reason is: Damaged
'Teardown for Token Revocation
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

'Setup for Token Revocation
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Lost")'Revocation Reason is: Lost
'Teardown for Token Revocation
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

'Setup for Token Revocation
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Upgrade")'Revocation Reason is: Upgrade
'Teardown for Token Revocation
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Damaged")'Replace Reason is: Damaged
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Lost")'Replace Reason is: Lost
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Upgrade")'Replace Reason is: Upgrade
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call GenerateTempLogonPassword("Tokens by user", enrollmentUserName)
Call UnlockTokenSAMManage("Tokens by user",enrollmentUserName,newPinToUnlockSAMManage)
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

'-----------------Close browser
Call CloseIEBrowser()
 @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf279.xml_;_
