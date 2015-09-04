'LoginSamService
Systemutil.Run ieExecutableLocation, samServiceUrl
Call LoginSamService(samUserName,samPassword)
With Browser("Browser")
	With .Dialog("Windows Security")
		.WinEdit("WinEdit").Set "samdemo\administrator" @@ hightlight id_;_1049404_;_script infofile_;_ZIP::ssf80.xml_;_
		.WinEdit("WinEdit").Type  micTab @@ hightlight id_;_1049404_;_script infofile_;_ZIP::ssf81.xml_;_
		.WinEdit("WinEdit_2").SetSecure "55e7e04057b2a1bc2def2b40b4918ad9d007e9c22f3e" @@ hightlight id_;_721484_;_script infofile_;_ZIP::ssf82.xml_;_
		.WinEdit("WinEdit_2").Type  micReturn @@ hightlight id_;_721484_;_script infofile_;_ZIP::ssf83.xml_;_
	End With
	'Enroll Token From SamService
	With .Page("SAM Self Service Center")
		.Link("Enroll a new smartcard").Click @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center").Link("Enroll a new smartcard")_;_script infofile_;_ZIP::ssf84.xml_;_
		.Link("Start").Click
		.Link("Yes. Continue with the").Click @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").Link("Yes. Continue with the")_;_script infofile_;_ZIP::ssf93.xml_;_
		.WebEdit("ctl00$main$txtTokenPin").SetSecure "55e7e14e627e62aa6414a9281ae4a046cb32c4e59a6d2446d622" @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").WebEdit("ctl00$main$txtTokenPin")_;_script infofile_;_ZIP::ssf94.xml_;_
		.WebButton("Submit").Click @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").WebButton("Submit")_;_script infofile_;_ZIP::ssf95.xml_;_
		.WebButton("Submit").Click @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").WebButton("Submit")_;_script infofile_;_ZIP::ssf100.xml_;_
	End With
	'-----------------Close browser
	SystemUtil.CloseProcessByName "iexplore.exe" @@ hightlight id_;_Browser("Browser").Page("SAM Self Service Center 2").Link("Back to main menu")_;_script infofile_;_ZIP::ssf101.xml_;_

	'loginSamManage
	Systemutil.Run ieExecutableLocation, samManageUrl

	With .Dialog("Windows Security")
		.WinEdit("WinEdit").Set "samdemo\administrator" @@ hightlight id_;_459988_;_script infofile_;_ZIP::ssf102.xml_;_
		.WinEdit("WinEdit").Type  micTab @@ hightlight id_;_459988_;_script infofile_;_ZIP::ssf103.xml_;_
		.WinEdit("WinEdit_2").SetSecure "55e7f0d3c38eebef9f3bffbd4b3f4be6a4b8739c2d1c" @@ hightlight id_;_918594_;_script infofile_;_ZIP::ssf104.xml_;_
		.WinButton("OK").Click @@ hightlight id_;_2491338_;_script infofile_;_ZIP::ssf105.xml_;_
	End With
	With .Page("SAM Management Center")
		.WebList("ddlSearchField1").Select "Tokens by user" @@ hightlight id_;_Browser("Browser").Page("SAM Management Center").WebList("ddlSearchField1")_;_script infofile_;_ZIP::ssf106.xml_;_
		.WebEdit("txtSearchValue1").Set "a" @@ hightlight id_;_Browser("Browser").Page("SAM Management Center").WebEdit("txtSearchValue1")_;_script infofile_;_ZIP::ssf107.xml_;_
		.WebButton("Go").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center").WebButton("Go")_;_script infofile_;_ZIP::ssf108.xml_;_
		.WebButton("Unassign").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 2").WebButton("Unassign")_;_script infofile_;_ZIP::ssf109.xml_;_
		.WebButton("Run").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 2").WebButton("Run")_;_script infofile_;_ZIP::ssf110.xml_;_
		wait(5)
		.WebButton("Done").Click @@ hightlight id_;_Browser("Browser").Page("SAM Management Center 2").WebButton("Done")_;_script infofile_;_ZIP::ssf111.xml_;_
	End With
End With


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
 	
 		        	If  s_Test_Name <>"" and  (Trim(s_Indicator)=Trim("Y") or Trim(s_Indicator)=Trim("y") ) and UCase(trim(s_Status))="INPROGRESS"  Then
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
 
 
