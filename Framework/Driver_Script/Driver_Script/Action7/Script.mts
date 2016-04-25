  
  
'===================================================================================
'===================================================================================


  
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService ReplaceToken Started---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")


Dim Token_Replace_Reason
'String the Value in 'Token_Replace_Reason' variable
Token_Replace_Reason=parameter("Replace_Reason")

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

 
 'Replace or Upgrade The Token
 If ReplaceOrUpgradeTheTokenSAMService(Token_Replace_Reason,NewPinForRevoke,SecondPort) = "PASS" Then
     Call fn_ExecutionLog("Function ReplaceOrUpgradeTheTokenSAMService has been passed", "PASS")
 Else
	Call fn_ExecutionLog("Function ReplaceOrUpgradeTheTokenSAMService---------->> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'To close the IE Brwoser
Call fn_CleanUp()


If OpenIEWithURL(samManageUrl)="PASS" Then
	Call fn_ExecutionLog("Function OpenIEWithURL(url) has been passed ", "PASS")
Else
	Call fn_ExecutionLog("Function OpenIEWithURL(url) ---> failed ", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'SAMManage login with admin credentials 
If LoginSamManage(samUserName,samPassword) = "PASS" Then
	Call fn_ExecutionLog("Function LoginSamManage has been passed", "PASS")	 
Else
	Call fn_ExecutionLog("Function LoginSamManage ------>> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If



'Compare the certificate Serial from SAC to SAM
If CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName)="PASS" Then
 	Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam has been passed", "PASS")
 Else	
   Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam ---------->>Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'Remove Token from Inventory
If RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)="PASS" Then
	Call fn_ExecutionLog("Function RemoveAllTokenFromInventory has been passed", "PASS")
 Else	
   Call fn_ExecutionLog("Function RemoveAllTokenFromInventory ----------->>>> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

'To Close the IE Browsers & SAC Tools
Call fn_CleanUp()
	
	
	Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService ReplaceToken Finished--->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
'=======================================================================









