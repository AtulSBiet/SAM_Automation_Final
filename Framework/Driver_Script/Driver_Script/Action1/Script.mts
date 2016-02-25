Option Explicit
Dim result
Dim Iterator

Call fn_Close_DIGI_Dialog()'Closing it if there is already process running otherwise opening digihub does not work
Call fn_DIGI_Open()
Call fn_DIGI_Connect_New_Token(FirstPort)
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

'Setup for Token Revocationi
Call fn_DIGI_Disconnect_Connected_Token()
Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Lost")'Revocation Reason is: Lost
'Teardown for Token Revocation
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

'Setup for Token Revocation
Call fn_DIGI_Disconnect_Connected_Token()
Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call RevokeTokenSAMManage("Upgrade")'Revocation Reason is: Upgrade
'Teardown for Token Revocation
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call fn_DIGI_Disconnect_Connected_Token()
Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Damaged")'Replace Reason is: Damaged
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call fn_DIGI_Disconnect_Connected_Token()
Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Lost")'Replace Reason is: Lost
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call ReplaceTokenSAMManage("Upgrade")'Replace Reason is: Upgrade
Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)

Call fn_DIGI_Disconnect_Connected_Token()
Call fn_DIGI_Connect_New_Token(FirstPort)
Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call GenerateTempLogonPassword("Tokens by user", enrollmentUserName)
Call UnlockTokenSAMManage("Tokens by user",enrollmentUserName,newPinToUnlockSAMManage)
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
Call RemoveTokenFromInventory("Connected tokens")
'Call RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)
Call fn_DIGI_Disconnect_Connected_Token()
Call fn_Close_DIGI_Dialog()
'-----------------Close browser
Call CloseIEBrowser() @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf279.xml_;_
