

'pre-requisite of this action -> first call SAMService_Enrollment action

'=============================================================================================

    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Operations Started---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

 
Call fn_Close_DIGI_Dialog()'Closing it if there is already process running otherwise opening digihub does not work
	
	
'Opening DIGI Hub Dialog
If fn_DIGI_Open()="PASS" then
     Call fn_ExecutionLog("fn_DIGI_Open has been passed ", "PASS")
 Else 
      Call fn_ExecutionLog("fn_DIGI_Open has been passed ", "FAIL") 
      Call fn_CleanUp()
	  ExitAction
End If


' Connect USB Token on the lab port
If fn_DIGI_Connect_New_Token(FirstPort) = "PASS" Then
	Call fn_ExecutionLog("fn_DIGI_Connect_New_Token has been passed ", "PASS")	
Else
	Call fn_ExecutionLog("fn_DIGI_Connect_New_Token---> failed ", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

'Open The IE Brwoser		
If OpenIEWithURL(samServiceUrl)="PASS" Then
	Call fn_ExecutionLog("Function OpenIEWithURL has been passed", "PASS")
Else 
	Call fn_ExecutionLog("Function OpenIEWithURL -------->. Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End if


'SAMService login with valid user credentials
If LoginSamService(SamUser1Name,SamUser1Password)="PASS" Then
	Call fn_ExecutionLog("Function LoginSamService has been passed", "PASS")
 Else 
	Call fn_ExecutionLog("Function LoginSamService-------> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End if
'
''Unlock Token from SAMService
'If fn_UnlockToken_SamService(NewPinForUnlock,NewPinForUnlock)="PASS" Then
'	call fn_ExecutionLog("Function fn_UnlockToken_SamService has been passed", "PASS")
'   Else
'Call fn_ExecutionLog("Function fn_UnlockToken_SamService--------->>Failed", "FAIL")
'	Call fn_CleanUp()
'	ExitAction
'End If
'
'Update Token Content from SAMService

If fn_UpdateTokenContent_SamService(TokenPasswordForUpdate,samUser1Name,samUser1Password )="PASS" Then
     Call fn_ExecutionLog("Function fn_UpdateTokenContent_SamService has been passed", "PASS")
Else
	 Call fn_ExecutionLog("Function fn_UpdateTokenContent_SamService ------>Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'Reset Toke Password from SAMService without current pin
If fn_ResetORChangeTokenPassword_SamService(NewPinForReset)="PASS" Then
	Call fn_ExecutionLog("Function fn_ResetORChangeTokenPassword_SamService has been passed", "PASS")
Else
	Call fn_ExecutionLog("Function fn_ResetORChangeTokenPassword_SamService------> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'Reset Toke Password from SAMService without current pin
If fn_ResetORChangeTokenPassword_SamService_WithCurrentPin(currentPinForReset,LatestPinForReset)="PASS" Then
	 Call fn_ExecutionLog("Function fn_ResetORChangeTokenPassword_SamService_WithCurrentPinhas been passed", "PASS")'"Reset Token Password with current PIN"
Else
	Call fn_ExecutionLog("Function fn_ResetORChangeTokenPassword_SamService_WithCurrentPinhas------>> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If



'Disable & ENable the Token
If fn_DisableAndEnableToken_Temp_SamService()="PASS" Then
	Call fn_ExecutionLog("Function fn_DisableAndEnableToken_Temp_SamService has been passed", "PASS")
   Else
    Call fn_ExecutionLog("Function fn_DisableAndEnableToken_Temp_SamService----------->>> FAiled", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

 Call fn_CleanUp()


  
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Operations Finished---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

'====================================================================================
