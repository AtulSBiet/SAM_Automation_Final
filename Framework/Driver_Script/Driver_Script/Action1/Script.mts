'LoginSamService
Systemutil.Run ieExecutableLocation, samServiceUrl
Call LoginSamService(samUserName,samPassword)
Call EnrollUSBTokenSAMService()	
'-----------------Close browser
SystemUtil.CloseProcessByName "iexplore.exe" @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").Link("Back to main menu")_;_script infofile_;_ZIP::ssf101.xml_;_
'loginSamManage
Systemutil.Run ieExecutableLocation, samManageUrl
'Call LoginSamManage(samUserName,samPassword)
'With Browser("Browser").Page("AdamLoginPage")
'		.WebEdit("UserName").Set samADAMUserName
'		.WebEdit("UserPass").SetSecure samPassword
'		.WebButton("Log On").Click
'End With

Call EnrollUSBTokenSAMManage("Users by username", enrollmentUserName)
Call CompareCertSerNoInSacAndSam()
Call UnassignTokenSAMManage("Tokens by user",enrollmentUserName)
'Call UnlockTokenSAMManage()
'Browser("Browser").Page("SAM Management Center_3").Link("Helpdesk").Click
'Browser("Browser").Page("SAM Management Center_2").WebList("ddlSearchField1").Select "Tokens by user"
'Browser("Browser").Page("SAM Management Center_2").WebButton("Go").Click
'Browser("Browser").Page("SAM Management Center_3").WebEdit("txtSearchValue1").Set "a"
'Browser("Browser").Page("SAM Management Center_3").WebButton("Go").Click
'Browser("Browser").Page("SAM Management Center_3").WebButton("Unlock").Click
'Window("Window_2").WinToolbar("Overflow Notification").Press "SafeNet Authentication Client"
'Window("SafeNet Authentication").Restore
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinButton("Button").Click
'Browser("Browser").Page("SAM Management Center_3").Link("Paste").Click
'Browser("Browser").Dialog("Internet Explorer").WinButton("Allow access").Click
'Browser("Browser").Page("SAM Management Center_3").WebButton("Run").Click
'Browser("Browser").Page("SAM Management Center_3").Link("Copy").Click
'Browser("Browser").Page("SAM Management Center_3").WebButton("Done").Click
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinEdit("Response Code:").WinMenu("ContextMenu").Select "Paste"
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinEdit("New Password:").SetSecure "56023cc584ed65abc1cb07874b7aed2ba2eb4c97e19e"
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinEdit("New Password:").Type  micTab 
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinEdit("Confirm Password:").SetSecure "56023ccc748eb0c4ead9a7887db4e32c51eacf4d5ffb"
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").WinButton("OK").Click
'Window("SafeNet Authentication").Dialog("Unlock Token: My Token").Dialog("Unlock Token: My Token").WinButton("OK").Click

Call DisableTokenSAMManage()
Call EnableTokenSAMMamange()

Browser("Browser").Page("SAM Management Center_5").WebList("ddlHelpDeskButtons").Select "Enable" @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 5").WebList("ddlHelpDeskButtons")_;_script infofile_;_ZIP::ssf207.xml_;_
Browser("Browser").Page("SAM Management Center_5").WebButton("Run").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 5").WebButton("Run")_;_script infofile_;_ZIP::ssf208.xml_;_
Browser("Browser").Page("SAM Management Center_5").WebButton("Done").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 5").WebButton("Done")_;_script infofile_;_ZIP::ssf209.xml_;_


Call RemoveTokenFromInventory("Connected tokens")


'UnassignFromSamMAnage

'Browser("Browser").Page("Page").Sync
'Option Explicit

'=======================================================================
      ' Author Name         : Ashish Kesarwani
      'Script Name          : Driver Script 
      'Script Description   : Driver Script to controlling the Execution of Test cases 
      'Last Updated Date    : < 21/08/2015> 
      'Last Update By       : Ashish
'=======================================================================

'Declaring Variables as Dim

Dim s_count
'Public s_RegSuite_Path,s_Execution_Log


 '--------------Initializing the Global Variables using Functional Library--------------
 
   ExecuteFile "C:\SAM_Automation_Final\Framework\Config_File\Config.txt"
   
 '---------------To capture Error if Config file not loaded------------------------------- 
   
   If err.number<>0 Then
   	 call fn_ExecutionLog("unable to associate config file, error_number-"&err.number&"Error Description -"&err.description,Fail)
   	 Reporter.ReportEvent micFail,"Config File Not Loaded" ,"Config File Not Loaded"
   	 ExitTest
   	 
   End If
   
   ' ----------------storing the Regression Suit Path to 's_RegSuite_Path'  Variable-------------
   
 's_RegSuite_Path = Folder_Path & Regression_Suit_Path
 
  'msgbox s_RegSuite_Path
 
  '-------------------Storing the Execution Log sheet path to 's_Execution_Log' variable---------------
 's_Execution_Log = Folder_Path & Execution_Log_Path 
   
' msgbox s_Execution_Log
 
 
 'Deleting previous log sheet 
'	Set delete_Log_sheet = createObject("Scripting.FileSystemObject")
'	
'	If delete_Log_sheet.FileExists(s_Execution_Log) Then
'	    delete_Log_sheet.CopyFile s_Execution_Log, Folder_Path & "Execution_Log\Old_Logs\",True
'	    'delete_Log_sheet.MoveFile m_Folder_Path & "Logs\Old_Logs\Log.xlsx",m_Folder_Path & "Logs\Old_Logs\Log"&cstr(now)&".xlsx"
'		delete_Log_sheet.deleteFile(s_Execution_Log)
'	End If
'	Set delete_Log_sheet=nothing
	
'Creating new Execution Log file	
	
	Set new_Log_Excel = createObject("Excel.application")
	new_Log_Excel.Workbooks.Add
	'new_Log_Excel.ActiveWorkbook.SaveAs (m_Folder_Path & "Logs\Log.xlsx")
	Set new_log_sheet1=new_Log_Excel.ActiveWorkbook.Worksheets("sheet1")
	new_log_sheet1.cells(1,1).value="S No"
	new_log_sheet1.cells(1,2).value="Time Stamp"
	new_log_sheet1.cells(1,3).value="Step Description"
	new_log_sheet1.cells(1,4).value="Status"
	new_Log_Excel.ActiveWorkbook.SaveAs (s_Execution_Log)
	new_Log_Excel.ActiveWorkbook.Close
	new_Log_Excel.application.Quit
	set new_Log_Excel=nothing
	Set new_log_sheet1=nothing
	SystemUtil.CloseProcessByName "Excel.exe *32"
	'call fn_ExecutionLog("Execution Log file Created Successfully","PASS")
	
' -----------------Adding Sheet1 for Batch Sheet--------------
	
  DataTable.AddSheet("sheet1")
	
'-------------------Execution of test cases as mentioned in Global Sheet in QTP Datatable	/Regression suite which is in 'InProgress Status--------
  While fn_check_Remaining_Execution<>0
   
       '-------------------Import test Regression suit to Global DataTable
        
 
         Datatable.Importsheet s_RegSuite_Path,"sheet1","sheet1"
 
       '-----------------Error handling to Verify Regression_Suite excel file is imported into QTP Datatable successfully or not-----		
		If err.number<>0 Then
		    Call fn_CaptureScreenshot("Regression_Suite")
			'reporter.ReportEvent micFail,"Regression Suite File Import:","Regression suite file is not imported - " & err.number &":" & err.description
			Call fn_ExecutionLog(err.number, err.description)
            ExitTest
   		End If 
  
       '--------------Get Rows Count from Global Sheet in Datatable------------------
	    s_Row_Count = Datatable.GetSheet("sheet1").GetRowCount
	    'msgbox s_Row_Count
   	
   	  For s_count = 1 To s_Row_Count Step 1
   	
   	    Datatable.SetCurrentRow s_count
   	    
   	     s_Test_Name = Datatable("Script_ID","sheet1")
        ' msgbox s_Test_Name
         s_IP_Address_Regression=DataTable("IP_Address","sheet1")
         'msgbox s_IP_Address_Regression
	     s_Indicator = Datatable("Indicator","sheet1")
	    ' msgbox s_Indicator
	     s_Status = Datatable("Status","sheet1")
	     'msgbox s_Status
   	        
   	         If Trim(s_IP_Address_Regression)=Trim(IP_Address) Then
 	
 		        	If  s_Test_Name <>"" and  (Trim(s_Indicator)=Trim("Y") or Trim(s_Indicator)=Trim("Y") ) and UCase(trim(s_Status))="INPROGRESS"  Then
		       '---------------------Calling Test Case Function to execute the Test Case Function------------------
		               call fn_ExecutionLog("Execution of"&s_Test_Name&"is going to start ","PASS")
		               
		        	    'msgbox s_Test_Name
		        	   call fn_ExecutionLog("Calling "&s_Test_Name,"PASS")
						strTestCase_ExecutionStatus=eval(s_Test_Name)
						'msgbox strTestCase_ExecutionStatus
						
				'----------------------Checking Status of Test Case-----------------------------------
					     If  UCase(Trim(strTestCase_ExecutionStatus))="PASS" Then
							call fn_ExecutionLog(s_Test_name&" executed Successfully" ,"PASS")
							
							DataTable("Execution_Status","sheet1")="PASS"
							
							'reporter.ReportEvent micPass,Datatable("Script_ID",Global) & " Execution Status:","Test Script execution is successful"
					     else
					        call fn_ExecutionLog(s_Test_name&" Not executed Successfully" ,"FAIL")
							'reporter.ReportEvent micFail,Datatable("Script_ID",Global) & " Execution Status:","Test Script execution is unsuccessful"	
				    	     DataTable("Execution_Status","sheet1")="FAIL"
				    	     
							
							Datatable.SetCurrentRow s_count
				    	     
				    	
				    	End If	
				    	
				'-------------  comment-----------------------  	
						     DataTable("Status","sheet1")="Done"
						    DataTable("Indicator","sheet1")="N"
						    Datatable.SetCurrentRow s_count+1
							DataTable("Indicator","sheet1")="Y"
							Datatable.SetCurrentRow s_count
						   
		        	    
		          End If  
		        
           End If
           
            
   	
           
                   If err.number<>0 Then
		           Call fn_CaptureScreenshot("Regression_Suite")
			      'reporter.ReportEvent micFail,"Regression Suite File Import:","Regression suite file is not imported - " & err.number &":" & err.description
			      Call fn_ExecutionLog(err.number, err.description)
                  ExitTest
   		          End If 
   	     
   	  '-------------- Exporting  the Global Sheet to Regression suit------------------
     Datatable.Export s_RegSuite_Path
   	
     Next
     
     
                If err.number<>0 Then
		           Call fn_CaptureScreenshot("Regression_Suite")
			      'reporter.ReportEvent micFail,"Regression Suite File Import:","Regression suite file is not imported - " & err.number &":" & err.description
			      Call fn_ExecutionLog(err.number, err.description)
                  ExitTest
   		          End If 
   		          
    ' -------------------Giving Wait time to check client execution is completed------------------------
    wait(1)
   call fn_ExecutionLog("Waiting for "&IP_Address&"to execute the remaining Test cases at their end","PASS" )
   	
 Wend
 
 '--------------------End Of Driver Script---------------------------------------
 
 
