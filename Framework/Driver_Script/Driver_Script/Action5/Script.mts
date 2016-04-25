
'=================================================================================================================
'=================================================================================================================

    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Enrollment Started---->>> ", "PASS")
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
	Call fn_ExecutionLog("Function OpenIEWithURL -------> ","FAIL")
	Call fn_CleanUp()
	ExitAction
End if


'SAMService login with valid user credentials
If LoginSamService(SamUser1Name,SamUser1Password)="PASS" Then
	Call fn_ExecutionLog("Function LoginSamService has been passed", "PASS")
 Else 
	Call fn_ExecutionLog("Function LoginSamService-------->> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End if


'Enroll the Token
If EnrollUSBTokenSAMService(NewTokenpinForEnroll)="PASS" Then
	Call fn_ExecutionLog("Function EnrollUSBTokenSAMService has been passed", "PASS")
    wait(5)	
Else 
	Call fn_ExecutionLog("Function EnrollUSBTokenSAMService--------->Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End if


 'To close the IE Browser
 Call fn_CleanUp()
 
 
 ' Opening IE Brwoser with SAMManage URL
 If OpenIEWithURL(samManageUrl) = "PASS" Then
    Call fn_ExecutionLog("Function OpenIEWithURL has been passed", "PASS")
 Else
	Call fn_ExecutionLog("Function OpenIEWithURL--------> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If




'SAMManage login with admin credentials 
If LoginSamManage(samUserName,samPassword) = "PASS" Then
	Call fn_ExecutionLog("Function LoginSamManage has been passed", "PASS")	
Else
	Call fn_ExecutionLog("Function LoginSamManage----------->>. Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

'Compare the certificate Serial from SAC to SAM
If CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName)="PASS" Then
    Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam has been passed", "PASS")	
Else
	 Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam -------> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

'To close the IE Browsers
 Call fn_CleanUp()
 
 
 
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Enrollment Finished--->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
'==========================================================================================================
'==========================================================================================================
