'
' If Browser("name:=SAM Self Service Center").Page("title:=SAM Self Service Center").Exist(4) Then
'	msgbox "TRUE"
'  Else 
'  msgbox "false"
'End If
'If Browser("name:=SAM Self Service Center").Page("title:=SAM Self Service Center").WebElement("innertext:=Token successfully disabled").Exist(30) then
'msgbox "TRUE"
'else
'msgbox "FALSE"
'End IF


'==========================================================================================================
'==========================================================================================================
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Enrollment strated---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

'Connecting Token to DIGI Hub

Call fn_Close_DIGI_Dialog()'Closing it if there is already process running otherwise opening digihub does not work

'Opening DIGI Hub Dialog
If fn_DIGI_Open()="PASS" then
     Call fn_ExecutionLog("fn_DIGI_Open has been passed ", "PASS")
 Else 
      Call fn_ExecutionLog("fn_DIGI_Open has been passed ", "FAIL") 
End If


' Connect USB Token on the lab port
If fn_DIGI_Connect_New_Token(FirstPort) = "PASS" Then
	Call fn_ExecutionLog("fn_DIGI_Connect_New_Token has been passed ", "PASS")	
Else
	Call fn_ExecutionLog("fn_DIGI_Connect_New_Token---> failed ", "FAIL")
	Call fn_CleanUp()
	ExitAction'TODO:OR should we use ExitTest
End If
 
 ' Opening IE Brwoser with SAMManage URL
If OpenIEWithURL(samManageUrl)="PASS" Then
	Call fn_ExecutionLog("Function OpenIEWithURL(url) has been passed ", "PASS")
Else
	Call fn_ExecutionLog("Function OpenIEWithURL(url) ---> failed ", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'SAMManage login with admin credentials 
If LoginSamManage(samUserName,samPassword) = "PASS" Then
	Call fn_ExecutionLog("LoginSamManage has been passed", "PASS")	
Else
	Call fn_ExecutionLog("LoginSamManage function ---> failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

    Call fn_ExecutionLog("==============================", "PASS")
    Call fn_ExecutionLog("Function  <<--RemoveAllTokenFromInventory--->>> started ", "PASS")
    Call fn_ExecutionLog("============================== ","PASS")

'Remove Token from Inventory
If RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName)="PASS" Then
	Call fn_ExecutionLog("Function RemoveAllTokenFromInventory has been passed ", "PASS")
Else
    Call fn_ExecutionLog("Function RemoveAllTokenFromInventory has been passed ", "PASS")
	Call fn_CleanUp()
	ExitAction
End If

    Call fn_ExecutionLog("==============================", "PASS")
    Call fn_ExecutionLog("Function  <<--RemoveAllTokenFromInventory--- Finished>>> ", "PASS")
    Call fn_ExecutionLog("============================== ","PASS")
  
'Enrolling Token
If EnrollUSBTokenSAMManage("Users by username", enrollmentUserName) = "PASS" Then
	Call fn_ExecutionLog("Function EnrollUSBTokenSAMManage has been passed ", "PASS")
Else  
    Call fn_ExecutionLog("Function EnrollUSBTokenSAMManage---> failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


    Call fn_ExecutionLog("==============================", "PASS")
    Call fn_ExecutionLog("Function  <<---CompareCertSerNoInSacAndSam--->>> started ", "PASS")
    Call fn_ExecutionLog("============================== ","PASS")


'Compare certificate Serial number from SAC To SAM
If CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName) = "PASS" Then
	call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam has been passed", "PASS")
Else
	Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam ---> failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

 
    Call fn_ExecutionLog("==============================", "PASS")
    Call fn_ExecutionLog("Function  <<---CompareCertSerNoInSacAndSam--->>> Finished ", "PASS")
    Call fn_ExecutionLog("============================== ","PASS") 
    
    
    
    
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Enrollment Finished---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
'===================================================================================================



