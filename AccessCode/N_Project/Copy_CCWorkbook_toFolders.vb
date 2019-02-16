Option Compare Database
Public PUBLIC_CREATEWB_Exit_wMessage As Boolean
Public PUBLIC_CREATEWB_DBLocked As Boolean
Public PUBLIC_CREATEWB_FileNotFound As Boolean
Public PUBLIC_CREATEWB_FileNotFound_Name As String
Public Function Macro_Run_PushTemplate_toFolder(sSLT_ColumnName As String, _
                                        sSLT_Value As String)


    fn01_Push_Template_toFolders TempVars("tmpV_Template_FilePath").Value, _
                                TempVars("tmpV_TemplateName").Value, _
                                TempVars("tmpV_WorkbookCopy_FilePath").Value, _
                                sSLT_ColumnName, sSLT_Value, _
                                TempVars("tmpV_FiscalYear").Value, TempVars("tmpV_WorkbookName").Value

End Function
Private Sub RunPush_Template()

    '\\NKE-WIN-NAS-P22\GOT_Budget\Fin Planning\Template_Production\
    
    '\\NKE-WIN-NAS-P22\GOT_Budget\@ACL\FY20 Financial Workbooks\AccessDB\
    '\\NKE-WIN-NAS-P22\GOT_Budget\@ACL\FY20 Financial Workbooks\
    
    'C:\Users\PSeie2\Desktop\wip - Financial Forecasting Process Build\Template_Production\
    '\\NKE-WIN-NAS-P22\GOT_Budget\Fin Planning\FY2019\
    

'   '>>>>>>>>>>>>>>> TEST <<<<<<<<<<<<<<<<<<<
'    fn01_Push_Template_toFolders "C:\Users\PSeie2\Desktop\wip - Financial Forecasting Process Build\Template_Production\", _
'                              "100000 Template Cost Center Workbook v0122.xlsm", _
'                              "\\NKE-WIN-NAS-P22\GOT_Budget\Fin Planning\FY2019\", _
'                               "SLT5", "PES", _
'                               2019, "Forecast.xlsm"


   ' >>>>>>>>>>>>>>> PRODUCTION <<<<<<<<<<<<<<<<<<<
    fn01_Push_Template_toFolders "C:\Users\PSeie2\Desktop\wip - Financial Forecasting Process Build\Template_Production\", _
                              "100000 Template Cost Center Workbook v0122.xlsm", _
                              "\\NKE-WIN-NAS-P22\GOT_Budget\@ACL\FY20 Financial Workbooks\", _
                               "SLT1", "TOPS", _
                               2020, "_CCWorkbook.xlsm"


End Sub
Private Function fn01_Push_Template_toFolders(sTemplate_FilePath As String, _
                                        sTemplate_Name As String, _
                                        sWorkbookCopy_FilePath As String, _
                                        sSLT_ColumnName As String, _
                                        sSLT_Value As String, _
                                        lFiscalYear As Long, _
                                        sSet_WorkbookName As String, _
                                        Optional sRunMacro_YesNo As String = "NO", _
                                        Optional sCostCenter_ColumnName As String = "CC")


 'Set Public Variables default values
  PUBLIC_CREATEWB_Exit_wMessage = True
  PUBLIC_CREATEWB_DBLocked = False
  PUBLIC_CREATEWB_FileNotFound = False
  PUBLIC_CREATEWB_FileNotFound_Name = ""

 'Local Variables
  Dim sTemplate_FilePathName As String
  
 'Set FilePathName
  sTemplate_FilePathName = sTemplate_FilePath & sTemplate_Name

 
 'Push the workboods to the folders
  If fn02_Create_Workbook_fromTemplate(sTemplate_FilePathName, _
                                    sWorkbookCopy_FilePath, _
                                    sSLT_ColumnName, _
                                    sSLT_Value, _
                                    lFiscalYear, _
                                    sSet_WorkbookName, _
                                    sRunMacro_YesNo, _
                                    sCostCenter_ColumnName) Then
  
     'Hide message if false
      If PUBLIC_CREATEWB_Exit_wMessage Then
      
       'Start Time
        Debug.Print vbCrLf & "Compleded Time " & Now() & vbCrLf
        
        Debug.Print "Succeed to create Template!!!"
        MsgBox "Success - Pushed Forecast Cost Center Worksheets to folders", vbInformation + vbOKOnly, "Macro Push Template to Folders"
  
      End If
    
  Else
  
     'Hide message if false
      If PUBLIC_CREATEWB_Exit_wMessage Then
      
        MsgBox "Failed to create Forecast Worksheet", vbExclamation + vbOKOnly, "FPA Database"
        
      Else
      
       'IF there was an error related to Locked DB
        If PUBLIC_CREATEWB_DBLocked Or Not PUBLIC_CREATEWB_FileNotFound Then
       

            Debug.Print "Locked = " & PUBLIC_CREATEWB_DBLocked & vbCrLf & " Name Range Not Found " & PUBLIC_CREATEWB_FileNotFound
    
        End If
        
      End If
  
  End If


End Function
Private Function fn02_Create_Workbook_fromTemplate(sTemplate_FilePathName As String, _
                                        sWorkbookCopy_FilePath As String, _
                                        sSLT_ColumnName As String, _
                                        sSLT_Value As String, _
                                        lFiscalYear As Long, _
                                        sSet_WorkbookName As String, _
                                        Optional sRunMacro_YesNo As String = "NO", _
                                        Optional sCostCenter_ColumnName As String = "CC") As Boolean


 'Excel Variables
  Dim appXL As Excel.Application
  Dim appXL_PushTemplate As Excel.Application
  Dim xlWrkBk As Excel.Workbook
  Dim xlWrkBk_Template As Excel.Workbook
  Dim xlWrkBk_Target As Excel.Workbook
  
 'Local Variable
  Dim sWorkbookList As String
  Dim iCount_OpenWorkbook As Integer

  
 On Error GoTo ProcErr
 
 'Set Default value
  fn02_Create_Workbook_fromTemplate = True
  iCount_OpenWorkbook = 0
  
 'Attempt to reference Excel which is already running.
  Set appXL = GetObject(, "Excel.Application")

 'If there is an Excel app running check to see if it is the template
  If Not appXL Is Nothing Then
  
   'Check all open workbook and list them
    For Each xlWrkBk In appXL.Workbooks
    
       'Get workbook names
        iCount_OpenWorkbook = iCount_OpenWorkbook + 1
        sWorkbookList = sWorkbookList & xlWrkBk.Name & vbCrLf
        
       'Check to see if the template is already open
        If xlWrkBk.FullName = sTemplate_FilePathName Then
        
           'If the template is open then let the user the ability to stop the program
            If MsgBox("Template is already open!" & vbCrLf & vbCrLf & _
                    "Do you want to continue to continue with Creating the worksheets from the template?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "FPA Database Notification") = vbNo Then
                    
              'If the user does nto want to continue then exit the program
               PUBLIC_CREATEWB_Exit_wMessage = False
               GoTo ProcExit
                    
            End If
          
        End If
        
      Next
    
   'Message if multiple workbooks are open
    If iCount_OpenWorkbook > 0 Then
    
      Debug.Print iCount_OpenWorkbook & "Workbooks that are Open on your computer. When running Push Template to Folder!" & vbCrLf & sWorkbookList
      
    End If
  
  End If
          
            
 ' >>>>>> Create a new Excel instance   <<<
 'If Excel is not running then Create a new instance of the Excel application.
  Set appXL_PushTemplate = New Excel.Application
 
 'If Excel is nothing then MS Excel is not installed in your system true.
  If appXL_PushTemplate Is Nothing Then
     
     fn02_Create_Workbook_fromTemplate = False
     MsgBox "MS Excel is not installed on your computer", vbCritical + vbOKOnly, "Error handled by FPA Database"
     Resume ProcExit
     
  End If

  With appXL_PushTemplate

      .Visible = False
      .AskToUpdateLinks = False
      .DisplayAlerts = False
      .EnableEvents = False
      .ScreenUpdating = False
  
  End With
  
' 'Set Template path file name
'  Debug.Print "Template file path name " & sTemplate_FilePathName & vbCrLf
  
 'Start Time
  Debug.Print "Start Time " & Now()
  
 'Update workbook
  Set xlWrkBk_Template = appXL_PushTemplate.Workbooks.Open("" & sTemplate_FilePathName & "", False, False)
  
  Debug.Print "Template Open " & xlWrkBk_Template.FullName
  
 'If Template not found then exit procedure
  If PUBLIC_CREATEWB_FileNotFound Then
  
    PUBLIC_CREATEWB_FileNotFound_Name = sTemplate_FilePathName
    Debug.Print "Template file name not found "
    GoTo ProcExit
    
  End If
  
 
 '*** MAKE COPIES OF THE TEMPLATE AND PUT THEM INTO THE CORRECT FOLDER ***
  If Not fn03_CopyTemplate_toFolder(xlWrkBk_Template, _
                                  sWorkbookCopy_FilePath, _
                                  sSLT_ColumnName, _
                                  sSLT_Value, _
                                  lFiscalYear, _
                                  sSet_WorkbookName, _
                                  sRunMacro_YesNo, _
                                  sCostCenter_ColumnName) Then
  
      fn02_Create_Workbook_fromTemplate = False
      GoTo ProcExit
  
  End If
  
  Debug.Print "Files Copied !!!"
     
ProcExit:
  

  With appXL_PushTemplate
      .ScreenUpdating = True
      .DisplayAlerts = True
      .EnableEvents = True
      .AskToUpdateLinks = True
  End With
    
 'App Excel Quit
  xlWrkBk_Template.Close

' 'Show the spreadsheet
'  appXL_PushTemplate.Visible = True

 'Quit Excel
  appXL_PushTemplate.Quit


Exit Function

ProcErr:

  Select Case Err.Number
  Case 91
    'Debug.Print fn02_Create_Workbook_fromTemplate
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 424, 429    'File Not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 462
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
  
  Case 1004
    fn02_Create_Workbook_fromTemplate = False
    PUBLIC_CREATEWB_FileNotFound = True
  
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
    fn02_Create_Workbook_fromTemplate = False
    PUBLIC_CREATEWB_Exit_wMessage = False
    PUBLIC_CREATEWB_DBLocked = True
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
  
  Case Else
    fn02_Create_Workbook_fromTemplate = False
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
Resume ProcExit
  

End Function
Private Function fn03_CopyTemplate_toFolder(xlWrkBk_Template As Excel.Workbook, _
                                        sWorkbookCopy_FilePath As String, _
                                        sSLT_ColumnName As String, _
                                        sSLT_Value As String, _
                                        lFiscalYear As Long, _
                                        sSet_WorkbookName As String, _
                                        Optional sRunMacro_YesNo As String = "NO", _
                                        Optional sCostCenter_ColumnName As String = "CC", _
                                        Optional sCostCenter_Status As String = "Active", _
                                        Optional sWrkSht_Control As String = "Control", _
                                        Optional sCellAddress_CostCenter As String = "D5", _
                                        Optional sCellAddress_FiscalYear As String = "D10", _
                                        Optional sWrkSht_Update_Run_Data As String = "Update Run Data", _
                                        Optional sNameRange_DBdata_FinishRunTime As String = "vbaPull_DBdata_FinishRunTime", _
                                        Optional sNameRange_RunUpdateMacro_YesNo As String = "vbaRunUpdate_BeforeSave", _
                                        Optional sNameRange_QueryCount As String = "vbaQueryCount") As Boolean


 'DAO Objects
  Dim qd As DAO.QueryDef
  Dim objRS_DAO As DAO.Recordset
  
 'Excel Variables
  Dim xlWrkSht_Control As Excel.Worksheet
  Dim xlWrkSht_UpdateRunData As Excel.Worksheet
  Dim Rng_CostCenter As Excel.Range
  Dim Rng_RunUpdateMacro_YesNo As Excel.Range
  Dim Rng_FinishRunTime As Excel.Range
  Dim Rng_FiscalYear As Excel.Range
  Dim Rng_QueryCount As Excel.Range

 'Local Variables
  Dim sWorksheetForcast_Name As String
  Dim sFilePath As String
  Dim sForecastWrkSt_FilePath_Name As String
  Dim lCostCenter As Long
  Dim i As Integer
  
 'Set the function to true
  fn03_CopyTemplate_toFolder = True
  
 On Error GoTo ProcErr
  
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

  'Debug.Print vbCrLf & sSQL & vbCrLf
    
 'NOTE I use QueryDef when referencing tables in a MS Access Database
  Set objRS_DAO = CurrentDb.CreateQueryDef("", sSQL).OpenRecordset

  Debug.Print "Number of files to be copied " & objRS_DAO.RecordCount
  
  objRS_DAO.MoveFirst
  

'  **** Update Template Workbook ***

 '*********************************************************************************************************
 '          >>>> Push the template ForecastWorksheet  <<<<
 '
 '          Update: CostCenter,FiscalYear and FilePath
 '

 'Set Range for Finish Run Time
  Set xlWrkSht_UpdateRunData = xlWrkBk_Template.Worksheets(sWrkSht_Update_Run_Data)
  Set Rng_FinishRunTime = xlWrkSht_UpdateRunData.Range(sNameRange_DBdata_FinishRunTime)


 'Set the range for the cost center
  Set xlWrkSht_Control = xlWrkBk_Template.Worksheets(sWrkSht_Control)
  Set Rng_CostCenter = xlWrkSht_Control.Range(sCellAddress_CostCenter)
  Set Rng_FiscalYear = xlWrkSht_Control.Range(sCellAddress_FiscalYear)
  Set Rng_QueryCount = xlWrkSht_Control.Range(sNameRange_QueryCount)


  ''Set the range for RunUpdateMacro_YesNo
  ' Set Rng_RunUpdateMacro_YesNo = xlWrkSht_UpdateRunData.Range(sNameRange_RunUpdateMacro_YesNo)
  '
  ' 'Set the Update Before save value to in cell D22 to yes
  '  Rng_RunUpdateMacro_YesNo.Value = sRunMacro_YesNo

  'Set int
  i = 0
  
 '** COST CENTER LOOP - List of cost center workbooks to be imported **
  Do While Not objRS_DAO.EOF
      
     'Get the file path and Cost Center from table Hierarchy
      If sWorkbookCopy_FilePath = "NONE" Then
      
        sFilePath = objRS_DAO.Fields("CostCenter_FilePath").Value
        
      Else
       
       'Concatenate FilePath with STL value
        sFilePath = sWorkbookCopy_FilePath & Replace(objRS_DAO.Fields(sSLT_ColumnName).Value, "&", "_") & "\"
        
      End If
    
      lCostCenter = objRS_DAO.Fields(sCostCenter_ColumnName).Value
      
     'Set the Cost Center in the Control worksheet cell D5
      Rng_CostCenter.Value = lCostCenter
      Rng_FiscalYear.Value = lFiscalYear
      
      
     '    >>>>>>>>>>>> Update the template Data Tabs per Query Tabs SQL statment <<<<<<<<<<
     '
     
          If Not fn04_Pull_QueryTabsSQL_into_DataTabs(xlWrkBk_Template, CInt(Rng_QueryCount.Value)) Then
          
            fn03_CopyTemplate_toFolder = False
            GoTo ProcExit
            
          End If

      
     'Update Data Macro Update (that is set to Finish Run Time)
      Rng_FinishRunTime.Value = Now()
      
     'Recalculate Worksheet control
      xlWrkSht_Control.Calculate
      
      
     ' >>> SAVE TEMPLATE <<<
      xlWrkBk_Template.Save
        
     ' >> Set the Excel Forecast FilePath Name with Worksheet Name <<
      sWorksheetForcast_Name = lCostCenter & "_" & "FY" & Right(lFiscalYear, 2) & sSet_WorkbookName
      sForecastWrkSt_FilePath_Name = sFilePath & sWorksheetForcast_Name
      
    'Increment int
      i = i + 1
      Debug.Print Format(i, "00") & " Excel Forecast FilePath Name " & sForecastWrkSt_FilePath_Name & "   Time " & Now()

     ' >>>> COPY TEMPLATE TO FOLDER AS A NEW FORECAST WORKSHEET <<<<
      xlWrkBk_Template.SaveCopyAs sForecastWrkSt_FilePath_Name
  
  
     'Set string to null
      sForecastWrkSt_FilePath_Name = vbNullString
      
     'Go to the next record in the recordset
      objRS_DAO.MoveNext
    
  Loop


ProcExit:

 'Close DAO
  objRS_DAO.Close
  Set objRS_DAO = Nothing

  Exit Function

ProcErr:

  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit

  Case 91
    Debug.Print "fn03_CopyTemplate_toFolder"
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 94 'Object not found
    fn03_CopyTemplate_toFolder = False
    PUBLIC_CREATEWB_FileNotFound = True
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Import Table name, Name Range or other parameter was not provided", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case 424   'Object not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 1004 'Workbook is open
    fn03_CopyTemplate_toFolder = False
    PUBLIC_CREATEWB_Exit_wMessage = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Workbook is already open! Exiting Program!" & sWorksheetForcast_Name & vbCrLf & vbCrLf & _
    "Take action on workbook and close it before running program", vbInformation + vbOKOnly, "Notice FPA Database"
    Resume ProcExit
    
  Case 3021 'SLT Not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical

    If MsgBox("SLT " & sSLT_ColumnName & " =" & sSLT_Value & " Was not found", vbOKCancel + vbExclamation, "Notice FPA Database") = vbCancel Then
          PUBLIC_CREATEWB_Exit_wMessage = False
          Resume ProcExit
    Else
          Resume Next
    End If

  Case 3704 'Recordset is already closed
    Resume Next

  Case 3709 'Connection object isnt open
    fn03_CopyTemplate_toFolder = False
    
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit

  Case -2147467259 'Invalid Path
    fn03_CopyTemplate_toFolder = False
    PUBLIC_CREATEWB_Exit_wMessage = False
    PUBLIC_CREATEWB_DBLocked = True
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit

  Case Else
    fn03_CopyTemplate_toFolder = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
 

End Function
Private Function fn04_Pull_QueryTabsSQL_into_DataTabs(xlWrkBk_Template As Excel.Workbook, iQueryCount As Integer) As Boolean

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
  Dim iNumberOfQueries As Integer
  Dim sWorksheetQuery_Name As String
  Dim sWorksheetData_Name As String
  Dim sCellAddress As String
  Dim lRowCount_RecordSet As Long
  Dim sConnectionString As String
 
 
 'Set the function to true
  fn04_Pull_QueryTabsSQL_into_DataTabs = True
  
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
  For Each xlWrkSht_Query In xlWrkBk_Template.Worksheets
  
      If Left(xlWrkSht_Query.Name, 5) = "Query" Then
      
        'Count queries run
         iCount = iCount + 1
      
        'Get Query worksheet name
         sWorksheetQuery_Name = xlWrkSht_Query.Name
         
        'Get SQL from Query worksheet
         sSQL = xlWrkSht_Query.Range("C2").Value
         
        'Get Worksheet Data
         sWorksheetData_Name = xlWrkSht_Query.Range("C5").Value
         Set xlWrkSht_Data = xlWrkBk_Template.Worksheets(sWorksheetData_Name)
         
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
          
         'Execute command object
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
        
         'Check Data Actual row count
          If sWorksheetData_Name = "Data_Actuals" Then
         
            Debug.Print sWorksheetData_Name & " Record Count = " & lRowCount_RecordSet
          
          End If
          
         'If there are no records returned from the query then skip CopyRecordset
         ' If lRowCount_RecordSet > 0 Then
          
             ' ***** Append rsDbDataSet to rsPUBLIC_Pivot ******
              If fn05_CopyRecordset_to_Spreadsheet(xlWrkSht_Data, rsDbDataSet, 3, 1, True, 500, 100) = False Then
        
                Debug.Print "**** Error Exited CopyRecordset Sub ******"
                GoTo ProcExit
        
              End If
              
         '  End If
           
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
    'Debug.Print "fn04_Pull_QueryTabsSQL_into_DataTabs"
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3021 'No record time has not been entered (SQL returned no records)
    'Debug.Print "No Records Returned for SQL query below"
    'Debug.Print "Source " & sWorksheetQuery_Name

    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset empty End program to stop more errors
    Resume Next
    
  Case -2147217913, -2147217900, -2147217904, -2147467259 'Error with the Criteria of the expression or SQL statement
    fn04_Pull_QueryTabsSQL_into_DataTabs = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print sSQL
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbInformation, "Error handled by FPA Database"
'    MsgBox "Send email to ITOPursuitsites@atos.net stating there is an error with Forecast Tools Reports" & vbCrLf & " With error # " & Err.Number & vbCrLf & "Send email to ITOPursuitsites@atos.net", vbExclamation + vbOKOnly
    Stop
'    Resume Next
    Resume ProcExit
    
  Case -2147217908
    fn04_Pull_QueryTabsSQL_into_DataTabs = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbInformation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case Else
    fn04_Pull_QueryTabsSQL_into_DataTabs = False
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    'Resume Next
    Resume ProcExit
    
  End Select

End Function

Private Function fn05_CopyRecordset_to_Spreadsheet(xlWrkSht_TARGET As Excel.Worksheet, _
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
  fn05_CopyRecordset_to_Spreadsheet = True
  
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
   fn05_CopyRecordset_to_Spreadsheet = False
   
   MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
   Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume ProcExit
  
  Case 9 'Worksheet not found
   fn05_CopyRecordset_to_Spreadsheet = False
   
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
    'Debug.Print "fn05_CopyRecordset_to_Spreadsheet"
    'Debug.Print "No Records Returned" & sWorksheetQuery_Name

    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3704 'Recordset is already closed
    Resume Next

  Case 3265 'Description of Item Can NOT be found in the recordset
   'If error then set fn05_CopyRecordset_to_Spreadsheet = False
    fn05_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit

  Case -2147467259 'Steam Object can't be read because it is empty
   'If error then set fn05_CopyRecordset_to_Spreadsheet = False
    fn05_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit
    
  Case Else
    fn05_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
  Resume ProcExit

End Function
Public Function Mod_CreateUpdateWrkbk_List_TempVars() As String

 Dim TempVar As TempVar
 Dim sString As String
 Dim sValue As String
 Dim sTempVar_Name As String

 TempVars("tmpV_List_TempVars").Value = ""
 
'Application.CurrentProject.AllMacros.
 lCountTempVar = TempVars.Count
 
 Debug.Print TempVars.Count
 
'List Variables
 For Each TempVar In TempVars
 
 ' sString = sString & "TempVar " & TempVar.Name & " = " & TempVar.Value & " Type is " & TypeName(TempVar.Value) & vbCrLf
  
  If Len(Trim(TempVar.Value)) <> 0 Then
  
    sTempVar_Name = sTempVar_Name & TempVar.Name & vbCrLf
    sValue = sValue & TempVar.Value & vbCrLf
    
  End If
  
 Next
 
 Debug.Print sTempVar_Name
 Debug.Print sValue
 
 TempVars("tmpV_List_TempVars") = sValue

End Function
Private Sub Run_StrConnectDB()

  Dim Conn As ADODB.Connection
  Set Conn = New ADODB.Connection
    

'Set connection string to current project (ie the database you are in)
 Set Conn = CurrentProject.Connection
 
 With Conn
 
  .CursorLocation = adUseClient
  
 End With
  
 Conn.Close
  'Debug.Print CurrentProject.Connection.ConnectionString
   'Debug.Print StrConnectDB(False) & vbCrLf
    
End Sub
Public Function StrConnectDB(sDB_FilePathName As String, Optional bConnectionStringCom As Boolean = True) As String
'This connects to the local database using OLEDB connection provider
'https://www.connectionstrings.com/access/

'OLEDB connection string to Access's Jet DB
  If bConnectionStringCom Then
  
'  'OLEDB connection string to Access's Jet DB
'     StrConnectDB = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                    "Data Source=" & CurrentDb.Name & ";" & _
'                    "User Id=admin;" & _
'                    "Password="
  
     StrConnectDB = "Microsoft.ACE.OLEDB.12.0;"
     StrConnectDB = StrConnectDB & "Data Source=" & sDB_FilePathName & ";"
     StrConnectDB = StrConnectDB & "User Id=Admin;" & vbCrLf
     StrConnectDB = StrConnectDB & "Persist Security Info=False;"
                    
   Else
   
    'Connect Project Connection string
      StrConnectDB = "Microsoft.ACE.OLEDB.12.0;" & vbCrLf
      StrConnectDB = StrConnectDB & "User Id=Admin;" & vbCrLf
      StrConnectDB = StrConnectDB & "Data Source=" & sDB_FilePathName & ";" & vbCrLf
      StrConnectDB = StrConnectDB & "Mode=Share Deny None;" & vbCrLf
      StrConnectDB = StrConnectDB & "Extended Properties="""";"
                    
    End If
End Function
Sub exampleIsProcessRunning()
    Debug.Print IsProcessRunning("MyProgram.EXE")
    Debug.Print IsProcessRunning("NOT RUNNING.EXE")

End Sub
Function IsProcessRunning(process As String)
    Dim objList As Object

    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If

    
End Function
Private Sub PowerShllCmd()


pscmd = "PowerShell -Command ""{Get-ScheduledTask -TaskName 'My Task' -CimSession MYLAPTOP}"""



End Sub
