'=========================================================================
'SAMManage Scenarios
'=========================================================================


    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Scenarios strated---->>> ", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")

'Scenario to test 'Lost Token' functionality


RunAction "SAMManage_Enrollment", oneIteration

RunAction "SAMManage_Operations", oneIteration

RunAction "SAMManage_ReplaceToken", oneIteration ,"Damaged"



' Scenario to test 'Upgrade Token' Functionality

RunAction "SAMManage_Enrollment", oneIteration

RunAction "SAMManage_ReplaceToken", oneIteration ,"Lost"


'Scenario to test 'Upgrade Token Functionality

RunAction "SAMManage_Enrollment", oneIteration

RunAction "SAMManage_ReplaceToken", oneIteration ,"Upgrade"






    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("Action <<--SAMManage Scenarios Finished---->>> ", "PASS")
    Call fn_ExecutionLog("======================================================", "PASS")
    Call fn_ExecutionLog("====================================================== ","PASS")
    Call fn_ExecutionLog("======================================================", "PASS")



