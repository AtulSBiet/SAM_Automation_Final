'call fn_ExecutionLog("fn_UpdateTokenContent_SamService", "Started")
'For Iterator = 1 To 10
Dim res
res="PASS" 
	Call fn_Close_DIGI_Dialog()'Closing it if there is already process running otherwise opening digihub does not work
	Call fn_DIGI_Open()
	Call fn_DIGI_Connect_New_Token(FirstPort)
'	
		'res=fn_CreateNewCertificate("TRUE",CertName)   '"Create New MSCA Certificate for Enrollment from SAMManage & SAMService".
If res="PASS" Then
	res=OpenIEWithURL(samServiceUrl)
End if
If res="PASS" Then
	res=LoginSamService(SamUser1Name,SamUser1Password)
End if
If res="PASS" Then
	
	res=EnrollUSBTokenSAMService(NewTokenpinForEnroll)
End if
If res="PASS" Then
		'"Update Content of Enrolled token"
	res= fn_UpdateTokenContent_SamService(TokenPasswordForUpdate,samUser1Name,samUser1Password )
End If
If res="PASS" Then
	res = fn_ResetORChangeTokenPassword_SamService(NewPinForReset) '"Reset Token Password without Current Pin"
End If
If res="PASS" Then
	res = fn_ResetORChangeTokenPassword_SamService_WithCurrentPin(currentPinForReset,LatestPinForReset) '"Reset Token Password with current PIN"
End If
res="PASS"
If res="PASS" Then
	res = fn_UnlockToken_SamService(NewPinForUnlock,NewPinForUnlock ) '"Unlock Token from SAC"
End If
If res="PASS" Then
	res = fn_DisableAndEnableToken_Temp_SamService ()   '"Disable and then Enable Token"
End If
If res="PASS" Then
	res = ReplaceOrUpgradeTheTokenSAMService("Lost",NewPinForRevoke,SecondPort)
	Call fn_DIGI_Disconnect_Connected_Token
	Call CloseIEBrowser ()
	call CleanFromSAMManage()
End If
res="PASS"
If res="PASS" Then
	Call fn_DIGI_Connect_New_Token(FirstPort)
	Call OpenIEWithURL(samServiceUrl)
	Call LoginSamService(SamUser1Name,SamUser1Password)
	Call EnrollUSBTokenSAMService(NewTokenpinForEnroll)
	res = ReplaceOrUpgradeTheTokenSAMService("Damage",NewPinForRevoke,SecondPort)
	Call fn_DIGI_Disconnect_Connected_Token
	Call CloseIEBrowser ()
	call CleanFromSAMManage()
End If

If res="PASS" Then
	Call fn_DIGI_Connect_New_Token(FirstPort)
	Call OpenIEWithURL(samServiceUrl)
	Call LoginSamService(SamUser1Name,SamUser1Password)
	Call EnrollUSBTokenSAMService(NewTokenpinForEnroll)
	res = ReplaceOrUpgradeTheTokenSAMService("Upgrade",NewPinForRevoke,SecondPort)
	Call CloseIEBrowser ()
	call CleanFromSAMManage()
End If
	Call fn_DIGI_Disconnect_Connected_Token()
	Call fn_Close_DIGI_Dialog()
	
'Next
