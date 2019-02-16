Option Compare Database
Public PUBLIC_CostCenter_Workbooks_InvalidPath As String
Public PUBLIC_PUSHACTUALS_Exit_wMessage As Boolean
Public PUBLIC_PUSHACTUALS_DBLocked As Boolean
Public PUBLIC_PUSHACTUALS_FileNotFound As Boolean
Public PUBLIC_PUSHACTUALS_FileNotFound_Name As String
Public PUBLIC_PUSHACTUALS_FileUpdated_Name As String
Public Function Macro_Push_DBdatasets_to_CostCenterWorksheets()


    fn01_PushDB_Datasets TempVars("tmpV_SLT_ColumnSelect"), _
                        TempVars("tmpV_SLT"), _
                        TempVars("tmpV_FiscalYear"), _
                        TempVars("tmpV_CostCenterWorkbookName"), _
                        True, "NONE", "Active"

End Function
Private Sub Run_PushDB_Datasets()


  fn01_PushDB_Datasets "SLT1", "PES", 2020, "_CCWorkbook.xlsm"


End Sub
Private Function fn01_PushDB_Datasets(sSLT_ColumnName As String, _
                                        sSLT_Value As String, _
                                        lFiscalYear As Long, _
                                        sCostCenter_WorkbookName_woPreFix As String, _
                                        Optional bDefault_FilePath As Boolean = True, _
                                        Optional sWorkbookCopy_FilePath As String = "NONE", _
                                        Optional sCostCenter_Status As String = "Active")

 'Local Var
  Dim sMessage As String

 'Set Public Variables default values
  PUBLIC_PUSHACTUALS_Exit_wMessage = True
  PUBLIC_PUSHACTUALS_DBLocked = False
  PUBLIC_PUSHACTUALS_FileNotFound = False
  PUBLIC_PUSHACTUALS_FileNotFound_Name = ""
  PUBLIC_PUSHACTUALS_FileUpdated_Name = ""

 
 'Push the workboods to the folders
  If fn02_PushDB_Datasets_to_CostCenterWorkbooks(sSLT_ColumnName, _
                                    sSLT_Value, _
                                    lFiscalYear, _
                                    sCostCenter_WorkbookName_woPreFix, _
                                    bDefault_FilePath, _
                                    sWorkbookCopy_FilePath, _
                                    sCostCenter_Status) Then
  
       'Hide message if false
        If PUBLIC_PUSHACTUALS_Exit_wMessage Then
        
         'Start Time
          Debug.Print vbCrLf & "Compleded Time " & Now() & vbCrLf
          
          Debug.Print "Succeed to create Template!!!"
          MsgBox "Success - Pushed Forecast Cost Center Worksheets to folders", vbInformation + vbOKOnly, "Macro Push Template to Folders"
          
        Else
        
           'Show message if filese were not found
            If PUBLIC_PUSHACTUALS_FileNotFound Then
            
               sMessage = "The Files were NOT found" & vbCrLf & _
                         PUBLIC_PUSHACTUALS_FileNotFound_Name & vbCrLf
                         
               If Len(PUBLIC_PUSHACTUALS_FileUpdated_Name) > 0 Then
               
                  sMessage = sMessage & "The files were updated" & vbCrLf & _
                              PUBLIC_PUSHACTUALS_FileUpdated_Name
                
                End If
                
               'List workbooks that were NOT updated and those there were updated
                MsgBox sMessage, vbOKOnly, "List Workbooks that were not updated"
           
            End If
    
        End If
    
  Else
  
      'Hide message if false
       If PUBLIC_PUSHACTUALS_Exit_wMessage Then
       
           MsgBox "Failed to create Forecast Worksheet", vbExclamation + vbOKOnly, "FPA Database"
         
       Else
           
          'IF there was an error related to Locked DB
           If PUBLIC_PUSHACTUALS_DBLocked Then
          
               Debug.Print "Locked = " & PUBLIC_PUSHACTUALS_DBLocked & vbCrLf & " Name Range Not Found " & PUBLIC_PUSHACTUALS_FileNotFound
       
           End If
         
       End If
  
  End If


End Function
Private Function fn02_PushDB_Datasets_to_CostCenterWorkbooks(sSLT_ColumnName As String, _
                                        sSLT_Value As String, _
                                        lFiscalYear As Long, _
                                        sCostCenter_WorkbookName_woPreFix As String, _
                                        bDefault_FilePath As Boolean, _
                                        sWorkbookCopy_FilePath As String, _
                                        sCostCenter_Status As String, _
                                        Optional sCostCenter_ColumnName As String = "CC", _
                                        Optional sWrkSht_Control As String = "Control", _
                                        Optional sCellAddress_CostCenter As String = "D5", _
                                        Optional sCellAddress_FiscalYear As String = "D10", _
                                        Optional sWrkSht_Update_Run_Data As String = "Update Run Data", _
                                        Optional sNameRange_DBdata_FinishRunTime As String = "vbaPull_DBdata_FinishRunTime", _
                                        Optional sNameRange_QueryCount As String = "vbaQueryCount") As Boolean


 'DAO Objects
  Dim qd As DAO.QueryDef
  Dim objRS_DAO As DAO.Recordset
  
 'Excel Objects
  Dim appXL_ExcelInstance As Excel.Application
  Dim xlWrkBk_CostCenterForecast As Excel.Workbook
  Dim xlWrkSht_Control As Excel.Worksheet
  Dim xlWrkSht_UpdateRunData As Excel.Worksheet
  Dim Rng_CostCenter As Excel.Range
  Dim Rng_FiscalYear As Excel.Range
  Dim Rng_QueryCount As Excel.Range
  Dim Rng_RunUpdateMacro_YesNo As Excel.Range
  Dim Rng_FinishRunTime As Excel.Range

 'Local Variables
  Dim sWorksheetForcast_Name As String
  Dim sFilePath As String
  Dim lFilePath_StringLength As Long
  Dim sForecastWrkSt_FilePath_Name As String
  Dim lCostCenter As Long
  Dim i As Integer
  
 'Set the function to true
  fn02_PushDB_Datasets_to_CostCenterWorkbooks = True
  bError_FileNotFound = False
  
 On Error GoTo ProcErr
  
  DoCmd.Hourglass True

 '  *** Create LIST OF COST CENTERS TO BE PULLED from the workbooks ***
 '
    sSQL = "SELECT tbl_Hierarchy.[" & sCostCenter_ColumnName & "],tbl_Hierarchy.[CC Name], tbl_Hierarchy.CostCenter_FilePath" & vbCrLf
    sSQL = sSQL & ", tbl_Hierarchy.[" & sSLT_ColumnName & "] , tbl_Hierarchy.CCStatus" & vbCrLf
    sSQL = sSQL & " FROM tbl_Hierarchy" & vbCrLf
  
   'If SLT value = ALL then enter LOAD all the cost centers
    If sSLT_Value = "ALL" Then
    
      sSQL = sSQL & " WHERE (((tbl_Hierarchy.CCStatus)='" & sCostCenter_Status & "'))" & vbCrLf
    
    Else
    
      sSQL = sSQL & " WHERE (((tbl_Hierarchy." & sSLT_ColumnName & ")= '" & sSLT_Value & "') " & vbCrLf
      sSQL = sSQL & " AND ((tbl_Hierarchy.CCStatus)='" & sCostCenter_Status & "'))" & vbCrLf
      
    End If

    sSQL = sSQL & " ORDER BY tbl_Hierarchy.[" & sCostCenter_ColumnName & "]"

'  Debug.Print CurrentDb.Name
'  Debug.Print vbCrLf & sSQL & vbCrLf
    
 'NOTE I use QueryDef when referencing tables in a MS Access Database
  Set objRS_DAO = CurrentDb.CreateQueryDef("", sSQL).OpenRecordset

  Debug.Print "Number of files to be copied " & objRS_DAO.RecordCount
  objRS_DAO.MoveFirst
  
  '----------------------------------------------------------------------------------------
  '                 >>>>>> Create a new Excel instance   <<<
  
  'If Excel is not running then Create a new instance of the Excel application.
   Set appXL_ExcelInstance = New Excel.Application
   
  'If Excel is nothing then MS Excel is not installed in your system true.
   If appXL_ExcelInstance Is Nothing Then
     
     fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
     PUBLIC_PUSHACTUALS_Exit_wMessage = False
     
     MsgBox "MS Excel is not installed on your computer", vbCritical + vbOKOnly, "Error handled by FPA Database"
     Resume ProcExit
     
   End If
   
  'Turn of updates
   With appXL_ExcelInstance
  
      .Visible = False
      .AskToUpdateLinks = False
      .DisplayAlerts = False
      .EnableEvents = False
      .ScreenUpdating = False
  
   End With
  
 '*********************************************************************************************************
 '          >>>> Push data from DB to Cost Center Forecast Workbooks  <<<<
 '
 '          Update: CostCenter,FiscalYear and FilePath
 '

 'Set int
  i = 0
  
 '** COST CENTER LOOP - List of cost center workbooks to be imported **
  Do While Not objRS_DAO.EOF
      
    'Use the file path from the CostCenter_FilePath field from tbl_Hierarchy
     If bDefault_FilePath Then
                                       
      'Get FilePath and remove the SLT categories
       sFilePath = objRS_DAO.Fields("CostCenter_FilePath").Value
      
     Else
     
      'Use file path passed by the argument sWorkbookCopy_FilePath and add it the the STL folder
      'NOTE: STL value name needs
       sFilePath = sWorkbookCopy_FilePath & Replace(objRS_DAO.Fields(sSLT_ColumnName).Value, "&", "_") & "\"

     End If
     
    'Get cost center
     lCostCenter = objRS_DAO.Fields(sCostCenter_ColumnName).Value
     
    ' >> Set the Excel Forecast FilePath Name with Worksheet Name <<
     sWorksheetForcast_Name = lCostCenter & "_" & "FY" & Right(lFiscalYear, 2) & sCostCenter_WorkbookName_woPreFix
     sForecastWrkSt_FilePath_Name = sFilePath & sWorksheetForcast_Name
     
    ' FILE PATH NAME
     Debug.Print "File Path Name " & sForecastWrkSt_FilePath_Name
     Debug.Print "Start Time Workbook Open " & Now()

       
    ' *********** OPEN WORKBOOK  **************
     Set xlWrkBk_CostCenterForecast = fn_OpenWorkbook(appXL_ExcelInstance, sFilePath, sWorksheetForcast_Name)
     
    'If the file is NOT found = TRUE then go to the next workbook to update
     If Not PUBLIC_PUSHACTUALS_FileNotFound Then
     
        'Set Range for Finish Run Time
         Set xlWrkSht_UpdateRunData = xlWrkBk_CostCenterForecast.Worksheets(sWrkSht_Update_Run_Data)
         Set Rng_FinishRunTime = xlWrkSht_UpdateRunData.Range(sNameRange_DBdata_FinishRunTime)
        
        
        'Set the range for the cost center
         Set xlWrkSht_Control = xlWrkBk_CostCenterForecast.Worksheets(sWrkSht_Control)
         Set Rng_CostCenter = xlWrkSht_Control.Range(sCellAddress_CostCenter)
         Set Rng_FiscalYear = xlWrkSht_Control.Range(sCellAddress_FiscalYear)
         Set Rng_QueryCount = xlWrkSht_Control.Range(sNameRange_QueryCount)
         
         
        'Set the Cost Center in the Control worksheet cell D5
         Rng_CostCenter.Value = lCostCenter
         Rng_FiscalYear.Value = lFiscalYear
          
          
        '------------------------------------------------------------------------------------------------
        '    >>>>>>>>>>>> Update the template Data Tabs per Query Tabs SQL statment <<<<<<<<<<
        '
         
            If Not fn03_CreateDataSet_fromQueryTabSQL(xlWrkBk_CostCenterForecast, CInt(Rng_QueryCount.Value)) Then
            
              fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
              GoTo ProcExit
              
            End If
     
       'Update Data Macro Update (that is set to Finish Run Time)
        Rng_FinishRunTime.Value = Now()
    
       'Recalculate Worksheet control
        xlWrkSht_Control.Calculate
        
       'Increment int
        i = i + 1
        Debug.Print Format(i, "00") & " Excel Forecast FilePath Name " & sForecastWrkSt_FilePath_Name & "   Time " & Now()
    
       ' >>> SAVE COST CENTER FORECAST WORKBOOK<<<
        With xlWrkBk_CostCenterForecast
            .Save
            .Close
        End With
    
    End If

   'Set string to null
    sForecastWrkSt_FilePath_Name = vbNullString

   'Go to the next record in the recordset
    objRS_DAO.MoveNext
    
  Loop

 'Show the spreadsheet if successful
  appXL_ExcelInstance.Visible = True
  
ProcExit:

  DoCmd.Hourglass False
  
 'Set Alerts back on
  With appXL_ExcelInstance
      .ScreenUpdating = True
      .DisplayAlerts = True
      .EnableEvents = True
      .AskToUpdateLinks = True
  End With
    
 'Quit Excel
  appXL_ExcelInstance.Quit

 'Close DAO
  objRS_DAO.Close
  Set objRS_DAO = Nothing

  Exit Function

ProcErr:

  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit

  Case 16, 91
    'Debug.Print "fn02_PushDB_datasets_to_CostCenterWorkbooks"
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 94 'Object not found
    fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
    PUBLIC_PUSHACTUALS_FileNotFound = True
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Import Table name, Name Range or other parameter was not provided", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case 424   'Object not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 1004 'Workbook is open
    PUBLIC_PUSHACTUALS_Exit_wMessage = False
    PUBLIC_PUSHACTUALS_FileNotFound = True
'    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
'
'    MsgBox "Workbook is already open or was not found! " & sWorksheetForcast_Name & vbCrLf & vbCrLf & _
'    "Take action on workbook and close it before running program", vbInformation + vbOKOnly, "Notice FPA Database"
    Resume Next
    
  Case 3021 'SLT Not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical

    If MsgBox("SLT " & sSLT_ColumnName & " =" & sSLT_Value & " Was not found", vbOKCancel + vbExclamation, "Notice FPA Database") = vbCancel Then
          PUBLIC_PUSHACTUALS_Exit_wMessage = False
          Resume ProcExit
    Else
          Resume Next
    End If
    
  Case 3061, 3075
    fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
    PUBLIC_PUSHACTUALS_Exit_wMessage = False
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "Incorrect information was Entered!." & vbCrLf & vbCrLf & "Check for the right SLT value was entered", vbInformation + vbOKOnly, "Notice FPA Database"
    Resume ProcExit
    
  Case 3704 'Recordset is already closed
    Resume Next

  Case 3709 'Connection object isnt open
    fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
    
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit

  Case -2147467259 'Invalid Path
    fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
    PUBLIC_PUSHACTUALS_Exit_wMessage = False
    PUBLIC_PUSHACTUALS_DBLocked = True
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit

  Case Else
    fn02_PushDB_Datasets_to_CostCenterWorkbooks = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
 

End Function
Private Function fn_OpenWorkbook(appXL_ExcelInstance As Excel.Application, _
                                sFilePath As String, _
                                sWorksheetForcast_Name As String) As Excel.Workbook


 'Excel Variables
  Dim xlWrkBk As Excel.Workbook
  
 'Local Variable
  Dim sForecastWrkSt_FilePath_Name As String
  Dim sWorkbookList As String
  Dim iCount_OpenWorkbook As Integer
  Dim iCount_Updated As Integer
  Dim iCount_NotUpated As Integer

  
 On Error GoTo ProcErr
 
 'Set Default value
  iCount_OpenWorkbook = 0
  PUBLIC_PUSHACTUALS_FileNotFound = False

' 'If there is an Excel app running check to see if it is the template
'  If Not appXL_ExcelInstance Is Nothing Then
'
'   'Check all open workbook and list them
'    For Each xlWrkBk In appXL_ExcelInstance.Workbooks
'
'       'Get workbook names
'        iCount_OpenWorkbook = iCount_OpenWorkbook + 1
'        sWorkbookList = sWorkbookList & xlWrkBk.Name & vbCrLf
'
'       'Check to see if the template is already open
'        If xlWrkBk.FullName = sForecastWrkSt_FilePath_Name Then
'
'           'If the template is open then let the user the ability to stop the program
'            If MsgBox("The Cost Center Forecast Workbook is already open!" & vbCrLf & vbCrLf & _
'                    "Do you want to continue to continue with Creating the worksheets from the template?", _
'                    vbExclamation + vbYesNo + vbDefaultButton2, "FPA Database Notification") = vbNo Then
'
'              'If the user does nto want to continue then exit the program
'               PUBLIC_PUSHACTUALS_Exit_wMessage = False
'               GoTo ProcExit
'
'            End If
'
'        End If
'
'      Next
'
'   'Message if multiple workbooks are open
'    If iCount_OpenWorkbook > 0 Then
'
'      Debug.Print iCount_OpenWorkbook & "Workbooks that are Open on your computer. When running Push Template to Folder!" & vbCrLf & sWorkbookList
'
'    End If
'
' End If

'----------------------------------------------------------------------------------------------------------
'                       *********** OPEN WORKBOOK  **************
'
  sForecastWrkSt_FilePath_Name = sFilePath & sWorksheetForcast_Name
  Debug.Print sForecastWrkSt_FilePath_Name
  Set fn_OpenWorkbook = appXL_ExcelInstance.Workbooks.Open("" & sForecastWrkSt_FilePath_Name & "", False, False)
  
  
  
 'If Template not found then exit procedure
  If PUBLIC_PUSHACTUALS_FileNotFound Then
  
    PUBLIC_PUSHACTUALS_FileNotFound_Name = PUBLIC_PUSHACTUALS_FileNotFound_Name & sWorksheetForcast_Name & vbCrLf
    Debug.Print "Template file name not found "
    GoTo ProcExit
    
  Else
    
    iCount_Updated = iCount_Updated + 1
    PUBLIC_PUSHACTUALS_FileUpdated_Name = PUBLIC_PUSHACTUALS_FileUpdated_Name & sWorksheetForcast_Name & vbCrLf
    
  End If


ProcExit:


Exit Function

ProcErr:

  Select Case Err.Number
  Case 91
    'Debug.Print fn_OpenWorkbook
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 424, 429    'File Not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 462
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
  
  Case 1004
    PUBLIC_PUSHACTUALS_Exit_wMessage = False
    PUBLIC_PUSHACTUALS_FileNotFound = True
  
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3704 'Recordset is already closed
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
  
  Case 50290 'Recordset is already closed
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case -2147417848 'The object invoked has disconnected from its clients
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case -2147467259 'Invalid Path
    PUBLIC_PUSHACTUALS_FileNotFound = True
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    'MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
  
  Case Else
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
Resume ProcExit
  

End Function
Private Function fn03_CreateDataSet_fromQueryTabSQL(xlWrkBk_CostCenterForecast As Excel.Workbook, iQueryCount As Integer) As Boolean

 'ADO Objects
  Dim Cmd As ADODB.Command
  Dim Conn As ADODB.Connection
  Dim rsDbDataSet As ADODB.Recordset

 'Excel Objects
  Dim xlWrkSht_Query As Excel.Worksheet
  Dim xlWrkSht_Data As Excel.Worksheet
  
 'Local variable
  Dim sSQL As String
  Dim iCount As Integer
  Dim sWorksheetQuery_Name As String
  Dim sWorksheetData_Name As String
  Dim sCellAddress As String
  Dim lRowCount_RecordSet As Long
  Dim sConnectionString As String
 
 
 'Set the function to true
  fn03_CreateDataSet_fromQueryTabSQL = True
  
 On Error GoTo ProcErr
 
 'Instantiate Objects
  Set Cmd = New ADODB.Command
  Set Conn = New ADODB.Connection
  Set rsDbDataSet = New ADODB.Recordset

'Set connection string to current project (ie the database you are in)
 Set Conn = CurrentProject.Connection
 
 With Conn
 
  .CursorLocation = adUseClient
  
 End With
  
 '* Get the SQL statment from each Query Tab and copy a recordset from that SQL statment into each corresponding Data tab
  For Each xlWrkSht_Query In xlWrkBk_CostCenterForecast.Worksheets
  
      If Left(xlWrkSht_Query.Name, 5) = "Query" Then
      
         'Count queries run
          iCount = iCount + 1
          
        'Get Query worksheet name
         sWorksheetQuery_Name = xlWrkSht_Query.Name
         
        'Get SQL from Query worksheet
         sSQL = xlWrkSht_Query.Range("C2").Value
         
        'Get Worksheet Data
         sWorksheetData_Name = xlWrkSht_Query.Range("C5").Value
         Set xlWrkSht_Data = xlWrkBk_CostCenterForecast.Worksheets(sWorksheetData_Name)
         
         'Debug.Print "Source " & sWorksheetQuery_Name & vbCrLf & _
                     "Target " & sWorksheetData_Name
              
         '***************************** Create Recordset ***********************
        
         'Set command object
          With Cmd
    
            .ActiveConnection = Conn
            .CommandText = sSQL
            .CommandType = adCmdText
            .CommandTimeout = 6  'NOTE: if user isn't logged into PreSale DB ie. have Forecast Tool website open. Then app will timeout
    
          End With
          
         '-------------------------------------------------------------
         '   Execute SQL statment and set the returned dataset to recordset rsDbDataSet
          Set rsDbDataSet = Cmd.Execute
          
         'Initialize to zero value
          lRowCount_RecordSet = 0
        
         'Count the number of rows and columns in the record set
          Do While Not rsDbDataSet.EOF
    
              lRowCount_RecordSet = lRowCount_RecordSet + 1
              rsDbDataSet.MoveNext
          Loop
    
         'Move to the first record
          rsDbDataSet.MoveFirst
        
          Debug.Print sWorksheetData_Name & " Record Count = " & lRowCount_RecordSet
          
         'If there are no records returned from the query then skip CopyRecordset
          If lRowCount_RecordSet > 0 Then
          
             ' ***** Append rsDbDataSet to rsPUBLIC_Pivot ******
              If fn_CopyDataSet_to_WorksheetDataTabs(xlWrkSht_Data, rsDbDataSet, 3, 1, True, 500, 100) = False Then
        
                Debug.Print "**** Error Exited CopyRecordset Sub ******"
                GoTo ProcExit
        
              End If
              
           End If
           
          sSQL = vbNullString
          
          
         '--------------------------------------------------------------------------------------------------------------
         '        >>>>>>>>>>>>>>>>>> Exit For Loop <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
         '
         '      NOTE: If the number of queries that were run equal were to the Number of Queries value from the Control tab
         '            then ProcExit
         
         'If sWorksheetQuery_Name = "Query3" Then
          If iCount = iQueryCount Then

            Debug.Print "Count of Queries run " & iCount
            Exit For

          End If

      End If
                   
  Next


ProcExit:
  
 'Close Connection object
  Conn.Close
  Set Conn = Nothing
  
  rsDbDataSet.Close
  Set rsDbDataSet = Nothing
  
  
Exit Function


ProcErr:

  Select Case Err.Number
  
  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    'Debug.Print "fn03_CreateDataSet_fromQueryTabSQL"
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3021 'No record time has not been entered (SQL returned no records)
    Debug.Print "No Records Returned for SQL query below"
    Debug.Print "Source " & sWorksheetQuery_Name

    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset empty End program to stop more errors
    Resume Next
    
  Case -2147217913, -2147217900, -2147217904, -2147467259 'Error with the Criteria of the expression or SQL statement
    fn03_CreateDataSet_fromQueryTabSQL = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print sSQL
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbInformation, "Error handled by FPA Database"
'    MsgBox "Send email to ITOPursuitsites@atos.net stating there is an error with Forecast Tools Reports" & vbCrLf & " With error # " & Err.Number & vbCrLf & "Send email to ITOPursuitsites@atos.net", vbExclamation + vbOKOnly
    Stop
'    Resume Next
    Resume ProcExit
    
  Case -2147217908
    fn03_CreateDataSet_fromQueryTabSQL = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbInformation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case Else
    fn03_CreateDataSet_fromQueryTabSQL = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    'Resume Next
    Resume ProcExit
    
  End Select

End Function

Private Function fn_CopyDataSet_to_WorksheetDataTabs(xlWrkSht_TARGET As Excel.Worksheet, _
                                                rsSpeadsheet As ADODB.Recordset, _
                                                lStartRow As Long, _
                                                lStartColumn As Long, _
                                                Optional bAddHeader As Boolean = False, _
                                                Optional lClearData_Rows As Long = 1, _
                                                Optional lClearData_Columns As Long = 1) As Boolean


'-------------------------------------------------------------------------------------------------------------
'

' CopyRecordset to Spreadsheet function
'                 ADODB recordset (contains data to be written to worksheet via copyrecordset method)
'                 Excel worksheet object
'                    (worksheet object receives the recordset from recordset copy method)
'                 String (name of worksheet that received the recordset)
'
' Arguments Passed
'
' 1) Write recordset data to worksheet
'              Arguments:
'                   a) Set worksheet: Excel workbook object xlWrkSht_TARGET
'                   b) Take recordset: rsSpeadsheet
'
' 2) Select first cell to Copyto recordset data
'    NOTE: The cell is the firts column and row of recordset data
'              Arguments:
'               Selection of cell to copy recordset from the 2 long variables (lStartRow and lStartColumn)
'
'
' Optional arguments
'
' 1) Create Header (default false)
'         Write header data if argument bAddHeard is TRUE or don't write header FALSE. FALSE is default value
'         Arguments: boolean (Write a header from the copied recordset TRUE=write header FALSE= do not write)

' 2) Clear data and formatting in worksheet
'        Set number of rows and columns to be cleared
'                Select end selection cell from the numeric arguments passed to (lClearData_Rows and lClearData_Columns
'                      Long (2 arguments to select cells in worksheet to clear data and formatting in copied to worksheet)

 
'************************************************************************************************
'*
'* NOTE: DO NOT USE MS EXCEL's  "Selection" for a substitute for a range of cells
'*          Excel will explicitly instantiate the "Selection" if you do use it
'*            You can't close the instantiate selection
'*


 '------- Local Variables ----------
  Dim lNumberFields As Long
  Dim lRowCount_RecordSet As Long
  Dim lColumnCount As Long
  Dim iCol As Integer
  Dim lHeaderRows As Long
  Dim sWorksheetQuery_Name As String
 
 '------- Excel Objects -------

  
 '---------Range Objects--------
  Dim objCell_1 As Excel.Range
  Dim objCell_2 As Excel.Range
  Dim objRange As Excel.Range
  
  
 'Set CopyRecordset to Spreadsheet to TRUE
  fn_CopyDataSet_to_WorksheetDataTabs = True
  
 '---------Constants------------
  Const ADDITIONAL_FIELD_INCREMENT = 2 'This is required to get he last field in a recordset. Recordset field starts with ZEREO
  Const HEADER_ROWS_ADDED = 1
 
On Error GoTo ProcErr

 'Get the name of the worksheet
  WorksheetQuery_Name = xlWrkSht_TARGET.Name

 '*** CLEAR ALL THE DATA (Clear Contents) IN THE WORKSHEET ***
 'Check that row count and column count are not 1
  If lClearData_Rows <> 1 And lClearData_Columns <> 1 Then
   
      Set objCell_1 = xlWrkSht_TARGET.Cells(lStartRow, lStartColumn)
      Set objCell_2 = xlWrkSht_TARGET.Cells(lStartRow + lClearData_Rows, lStartColumn + lClearData_Columns)
   
      Set objRange = xlWrkSht_TARGET.Range(objCell_1, objCell_2)
      'Debug.Print objRange.Worksheet.Name & " Clear Data Range = " & objRange.Address
      
      objRange.ClearContents
  
  End If

 '   >>>> Check to see if headers are to be added: AddHeader True/False  <<<<
  If bAddHeader Then

       '*** Copy Field Header into the spreadsheet using recordset fieldnames***
       'NOTE: you need to add 1 to get all fields. Since recordset field start with 0
         
       'Count number of fields
        lNumberFields = rsSpeadsheet.Fields.Count ' + ADDITIONAL_FIELD_INCREMENT

         
       'Copy field names to the first row of the worksheet
        For iCol = lStartColumn To lNumberFields + lStartColumn - 1
        
           'Debug.Print rsSpeadsheet.Fields(iCol - lStartColumn).Name
           xlWrkSht_TARGET.Cells(lStartRow, iCol).Value = rsSpeadsheet.Fields(iCol - lStartColumn).Name
               
        Next
      
        lHeaderRows = HEADER_ROWS_ADDED
        
    Else
      
       'Zero rows are added when FALSE
        lHeaderRows = 0

  End If
  
 
 '--------------------------------------------------------------------------------------------
 '            >>>>>>>>>    COPY RECORDSET TO SPREADSHEET    <<<<<<<<
 '
 rsSpeadsheet.MoveFirst
  xlWrkSht_TARGET.Cells(lStartRow + lHeaderRows, lStartColumn).CopyFromRecordset rsSpeadsheet

 'Recalculate worksheet
  xlWrkSht_TARGET.Calculate
   

ProcExit:

'Close recordset
 rsSpeadsheet.Close
 Set rsSpeadsheet = Nothing
 
Exit Function

ProcErr:

  Select Case Err.Number
  
  Case 5 'Recordset error
   fn_CopyDataSet_to_WorksheetDataTabs = False
   
   MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
   Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume ProcExit
  
  Case 9 'Worksheet not found
   fn_CopyDataSet_to_WorksheetDataTabs = False
   
   MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
   Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume ProcExit
    
  Case 91, 424 'Hourglass Comand
  'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume Next
    
  Case 1004
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3021 'No record time has not been entered (SQL returned no records)
    'Debug.Print "fn_CopyDataSet_to_WorksheetDataTabs"
    'Debug.Print "No Records Returned" & sWorksheetQuery_Name

    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3704 'Recordset is already closed
    Resume Next

  Case 3265 'Description of Item Can NOT be found in the recordset
   'If error then set fn_CopyDataSet_to_WorksheetDataTabs = False
    fn_CopyDataSet_to_WorksheetDataTabs = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit

  Case -2147467259 'Steam Object can't be read because it is empty
   'If error then set fn_CopyDataSet_to_WorksheetDataTabs = False
    fn_CopyDataSet_to_WorksheetDataTabs = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit
    
  Case Else
    fn_CopyDataSet_to_WorksheetDataTabs = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
  Resume ProcExit

End Function



