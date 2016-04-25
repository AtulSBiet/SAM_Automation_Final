'===============================================================
'================================================================
'Pre-requisite to run this action -> call SAMManage_Enrollment 


    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Operations strated---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

If Browser("name:=SAM Management Center").Exist Then
	Browser("name:=SAM Management Center").Sync
 else
    Call fn_ExecutionLog("NO SAMManage page found  ---> Failed", "FAIL")
    Call fn_CleanUp()
	ExitAction 	
End If
'Disable Token

If DisableTokenSAMManage("Tokens by user",enrollmentUserName)="PASS" Then
	Call fn_ExecutionLog("Function DisableTokenSAMManage has been passed", "PASS")
Else 
    Call fn_ExecutionLog("Function DisableTokenSAMManage ------> Failed", "FAIL")
    Call fn_CleanUp()
	ExitAction
End If	 


'Enable Token
If EnableTokenSAMManage("Tokens by user",enrollmentUserName) = "PASS" Then
	Call fn_ExecutionLog("Function EnableTokenSAMManage has been passed", "PASS")
Else
	Call fn_ExecutionLog("Function EnableTokenSAMManage ----> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'Unlock Token
If UnlockTokenSAMManage("Tokens by user",enrollmentUserName,newPinToUnlockSAMManage) = "PASS" Then	
      Call fn_ExecutionLog("Function UnlockTokenSAMManage has been passed", "PASS")
Else
	Call fn_ExecutionLog("Function UnlockTokenSAMManage -----> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If
 
 
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Operations Finished---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
'============================================================================

'add more SAMManage Operations

'================================================================



