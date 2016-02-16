Option Explicit
Dim res
res="PASS"'res=fn_CreateNewCertificate("TRUE",CertName)           '"Create New MSCA Certificate for Enrollment from SAMManage & SAMService".
'If res="PASS" Then
'	res=fn_Initialize_by_SACUsingSimplePassword("My Token",defaultTokenPin)
'End if
'If res="PASS" Then
'	res=OpenIEWithURL(samServiceUrl)
'End if
'If res="PASS" Then
'	res=LoginSamService(samUserName,samPassword)
'End if
'If res="PASS" Then
'	res=EnrollUSBTokenSAMService(newTokenPin,defaultTokenPin)
'End if
'If res="PASS" Then
''"Update Content of Enrolled token"
'	res= fn_UpdateTokenContent_SamService(defaultTokenPinWithoutEncoding,NewPinForUpdate,samUserName,samPassword )
'End If
'If res="PASS" Then
'	res = fn_ResetORChangeTokenPassword_SamService(NewPinForReset) '"Reset Token Password without Current Pin"
'End If
If res="PASS" Then
	res = fn_ResetORChangeTokenPassword_SamService_WithCurrentPin(defaultTokenPinWithoutEncoding,LatestPinForReset) '"Reset Token Password with current PIN"
End If
If res="PASS" Then
	res = fn_UnlockToken_SamService(NewPinForUnlock,NewPinForUnlock ) '"Unlock Token from SAC"
End If
If res="PASS" Then
	res = fn_DisableAndEnableToken_Temp_SamService ()   '"Disable and then Enable Token"
End If
If res="PASS" Then
	res = ReplaceOrUpgradeTheTokenSAMService("Lost",defaultTokenPin,"2")
End If
If res="PASS" Then
	res = CloseIEBrowser()
End If
If res="PASS" Then
	res = CleanFromSAMManage()
End If

'TODO: Make following 6 calls as a function
If res="PASS" Then
	res=OpenIEWithURL(samServiceUrl)
End if
If res="PASS" Then
	res=LoginSamService(samUserName,samPassword)
End if
If res="PASS" Then
	res=EnrollUSBTokenSAMService(newTokenPin,defaultTokenPin)
End if
If res="PASS" Then
	res = ReplaceOrUpgradeTheTokenSAMService("Damaged",defaultTokenPin,"2")
End If
If res="PASS" Then
	res = CloseIEBrowser()
End If
If res="PASS" Then
	res = CleanFromSAMManage()
End If

If res="PASS" Then
	res=OpenIEWithURL(samServiceUrl)
End if
If res="PASS" Then
	res=LoginSamService(samUserName,samPassword)
End if
If res="PASS" Then
	res=EnrollUSBTokenSAMService(newTokenPin,defaultTokenPin)
End if
If res="PASS" Then
	res = ReplaceOrUpgradeTheTokenSAMService("Upgrade",defaultTokenPin,"2")
End If
If res="PASS" Then
	res = CloseIEBrowser()
End If
If res="PASS" Then
	res = CleanFromSAMManage()
End If


'************************************
Call OpenIEWithURL(samServiceUrl)
Call LoginSamService(samUserName,samPassword)
Call EnrollUSBTokenSAMService(newTokenPin,defaultTokenPin)
Call fn_UpdateTokenContent_SamService(defaultTokenPin,defaultTokenPin,samUserName,samPassword )'Call fn_UpdateTokenContent_SamService(defaultTokenPin,NewPinForUpdate,samUserName,samPassword )
Call fn_ResetORChangeTokenPassword_SamService(newPinForReset) '"Reset Token Password without Current Pin"
Call fn_ResetORChangeTokenPassword_SamService_WithCurrentPin(currentPinForReset,latestPinForReset) '"Reset Token Password with current PIN"
Call fn_UnlockToken_SamService(newPinForUnlock,newPinForUnlock ) '"Unlock Token from SAC"
Call fn_DisableAndEnableToken_Temp_SamService ()   '"Disable and then Enable Token"
Call ReplaceOrUpgradeTheTokenSAMService("Lost",defaultTokenPin,"2")
Call CloseIEBrowser()
Call  CleanFromSAMManage()

