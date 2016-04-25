

'=============================================================================
'=============================================================================
'Pre-requisite to call this action -> call SAMManage Enrollment

    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage ReplaceToken strated---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

If Browser("name:=SAM Management Center").Exist Then
	Browser("name:=SAM Management Center").Sync
 else
    Call fn_ExecutionLog("NO SAMManage page found  ---> Failed", "FAIL")
    Call fn_CleanUp()
	ExitAction 	
End If

'Option Explicit
Dim Token_Replace_Reason1

'Passing Parameter
Token_Replace_Reason1=parameter("Token_replace_Value")

'Replace or Revoke the Token
If ReplaceTokenSAMManage(Token_Replace_Reason1) = "PASS" Then
	Call fn_ExecutionLog("Function ReplaceTokenSAMManage", "PASS")
Else
	Call fn_ExecutionLog("Function  ReplaceTokenSAMManage ---> Failed ", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If


'Compare certificate Serial number from SAC To SAM
If CompareCertSerNoInSacAndSam("Tokens by user",enrollmentUserName) = "PASS" Then
	
	Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam has been passed", "PASS")
Else
	Call fn_ExecutionLog("Function CompareCertSerNoInSacAndSam ------> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If



'Login SAMManage with admin credential
If RemoveAllTokenFromInventory("Tokens by user",enrollmentUserName) = "PASS" Then
	Call fn_ExecutionLog("Function LoginSamManage has been Passed", "PASS") 
Else
	Call fn_ExecutionLog("Function LoginSamManage   ----> Failed", "FAIL")
	Call fn_CleanUp()
	ExitAction
End If

'Close all the Browsers & windows before next scenario starts
Call fn_CleanUp()
	
	
	Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage ReplaceToken Finished---->>> ", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
'===============================================================
'================================================================
