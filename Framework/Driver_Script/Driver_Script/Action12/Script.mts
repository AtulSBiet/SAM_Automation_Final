'==========================================
'SAMService Scenarios
'==========================================
  
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Scenarios strated---->>> ", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
  
  
'To test 'Lost Token ' scenarios from SAM Service portal

RunAction "SAMService_Enrollment", oneIteration

RunAction "SAMService_Operations", oneIteration

RunAction "SAMService_ReplaceToken", oneIteration ,Lost



'To test 'Lost Token ' scenarios from SAM Service portal

RunAction "SAMService_Enrollment", oneIteration

RunAction "SAMService_ReplaceToken", oneIteration ,"Damaged"



'To test 'Lost Token ' scenarios from SAM Service portal

RunAction "SAMService_Enrollment", oneIteration

RunAction "SAMService_ReplaceToken", oneIteration ,"Upgrade"




    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMService Scenarios Finished---->>> ", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
