Option Explicit
result = "FAIL"

If result = "PASS" Then
	'loginSamManage
	result = LoginSamManage(samUserName,samPassword)
Else
	Call fn_ExecutionLog("LoginSamManage", "Not Started")
	ExitAction'TODO:OR should we use ExitTest
End If

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call UnlockTokenSAMManage("Tokens by user",enrollmentUserName,newPinToUnlockSAMManage)
Call CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName)
Call EnableTokenSAMManage("Tokens by user",enrollmentUserName)
Call DisableTokenSAMManage("Tokens by user",enrollmentUserName)
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")

'Setup for Token Revocation
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Damaged")'Revocation Reason is: Damaged
'Teardown for Token Revocation
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")

'Setup for Token Revocation
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Lost")'Revocation Reason is: Lost
'Teardown for Token Revocation
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")

'Setup for Token Revocation
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Upgrade")'Revocation Reason is: Upgrade
'Teardown for Token Revocation
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Damaged")'Replace Reason is: Damaged
Call RemoveTokenFromInventory("Connected tokens")

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Lost")'Replace Reason is: Lost
Call RemoveTokenFromInventory("Connected tokens")

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Upgrade")'Replace Reason is: Upgrade
Call RemoveTokenFromInventory("Connected tokens")

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call GenerateTempLogonPassword("Tokens by user", enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")

MsgBox "Stop Test"
'-----------------Close browser
Call CloseIEBrowser()
