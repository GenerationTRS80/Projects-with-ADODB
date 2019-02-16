Option Compare Database
Public PUBLIC_Exit_wMessage As Boolean
Public PUBLIC_DBLocked As Boolean
Public PUBLIC_Table_NotFound As Boolean
Public PUBLIC_FileNotFound As Boolean
Public PUBLIC_FileNotFound_Name As String
Public PUBLIC_NameRangeNotFound As Boolean
Public PUBLIC_NameRangeNotFound_Name As String
Public Function Macro_Pull_CostCenter_Workbook_GreenTabs()

        aa_Main_Pull_CostCenter_Workbook_GreenTabs TempVars("tmpV_SLT_ColumnSelect"), _
                        TempVars("tmpV_SLT"), _
                        TempVars("tmpV_FiscalYear"), _
                        TempVars("tmpV_CostCenterWorkbookName"), _
                        False, "DBPull_Forecast", _
                        True, "NONE", "Active"
                      
                        
End Function

Private Sub Run_Pull_CCWorkbooks()

                        
   'Debug.Print TempVars("tmpV_WorkbookName").Value
    aa_Main_Pull_CostCenter_Workbook_GreenTabs "SLT1", "SCEF", 2020, "_CCWorkbook.xlsm", _
                                              False, "DBPull_Forecast", _
                                              True, "NONE", "Active"
    
            
End Sub
Public Function Macro_Pull_Forecast_CostCenter_Workbook()

        aa_Main_Pull_CostCenter_Workbook_GreenTabs TempVars("tmpV_SLT_ColumnSelect"), _
                        TempVars("tmpV_SLT"), _
                        TempVars("tmpV_FiscalYear"), _
                        TempVars("tmpV_CostCenterWorkbookName"), _
                        True, "DBPull_Forecast", _
                        True, "NONE", "Active"
                      
                        
End Function
Private Sub Run_Pull_Forecast_CCWorkbook()

                        
   'Debug.Print TempVars("tmpV_WorkbookName").Value
    aa_Main_Pull_CostCenter_Workbook_GreenTabs "SLT1", "TOPS", 2020, "_Bak_CCWorkbook.xlsm", _
                                              True, "DBPull_Forecast", _
                                              True, "NONE", "Active"
    
            
End Sub
Public Function aa_Main_Pull_CostCenter_Workbook_GreenTabs( _
                                                  sSLT_ColumnName As String, _
                                                  sSLT_Value As String, _
                                                  lFiscalYear As Long, _
                                                  sCostCenter_WorkbookName_woPreFix As String, _
                                                  Optional bImport_byNameRange_YesNo As Boolean = False, _
                                                  Optional sImport_NameRange_Name As String = "DBPull_Forecast", _
                                                  Optional bDefault_FilePath As Boolean = True, _
                                                  Optional sWorkbookPull_FilePath As String = "NONE", _
                                                  Optional sCostCenter_Status As String = "Active", _
                                                  Optional sCostCenter_ColumnName As String = "CC", _
                                                  Optional lForecastMonth_ID As Long = 308, _
                                                  Optional bImport_ForecastWorksheet As Boolean = False)

  'Set Public Variables default values
   PUBLIC_Exit_wMessage = True
   PUBLIC_DBLocked = False
   PUBLIC_Table_NotFound = False
   PUBLIC_FileNotFound = False
   PUBLIC_FileNotFound_Name = ""
   PUBLIC_NameRangeNotFound = False
   PUBLIC_NameRangeNotFound_Name = ""
  
  'Run function
   If ab_Import_CostCenterWorksheets_per_tblHierarchy( _
                                                      sSLT_ColumnName, _
                                                      sSLT_Value, _
                                                      lFiscalYear, _
                                                      sCostCenter_WorkbookName_woPreFix, _
                                                      bImport_byNameRange_YesNo, _
                                                      sImport_NameRange_Name, _
                                                      bDefault_FilePath, _
                                                      sWorkbookPull_FilePath, _
                                                      sCostCenter_Status, _
                                                      sCostCenter_ColumnName, _
                                                      lForecastMonth_ID) Then
                                                      
        'List any Excel files that were not imported
         If PUBLIC_FileNotFound Or PUBLIC_NameRangeNotFound Then
         
            If PUBLIC_FileNotFound Then
            
              MsgBox "The file(s) NOT found" & vbCrLf & PUBLIC_FileNotFound_Name, vbOKOnly + vbInformation, "Error FPA Database"
          
            End If
            
            If PUBLIC_NameRangeNotFound Then
            
              MsgBox "The Name Range not found " & vbCrLf & PUBLIC_NameRangeNotFound_Name, vbOKOnly + vbInformation, "Error FPA Database"
          
            End If
           
           
         Else
         
             MsgBox "Macro Success - Completed Import of the Forecast Cost Center Worksheets", vbInformation + vbOKOnly, "Macro Commpleted FPA Database"
         
         End If
    
    Else
    
       'Exit WITHOUT am error message if false
        If PUBLIC_Exit_wMessage Then
        
           'IF there was an error related to Locked DB or Table NameRange not found then don't show messagebox below
            If Not PUBLIC_DBLocked And Not PUBLIC_Table_NotFound Then
           
                MsgBox "Failed to import Cost Center Worksheets into table ", vbExclamation + vbOKOnly, "FPA Database"
             
            Else
           
                Debug.Print "Locked = " & PUBLIC_DBLocked & vbCrLf & " Name Range Not Found " & PUBLIC_Table_NotFound
        
            End If
            
        End If
      
   End If


End Function
Private Function ab_Import_CostCenterWorksheets_per_tblHierarchy( _
                                                      sSLT_ColumnName As String, _
                                                      sSLT_Value As String, _
                                                      lFiscalYear As Long, _
                                                      sCostCenter_WorkbookName_woPreFix As String, _
                                                      bImport_byNameRange_YesNo As Boolean, _
                                                      sImport_NameRange_Name As String, _
                                                      bDefault_FilePath As Boolean, _
                                                      sWorkbookPull_FilePath As String, _
                                                      sCostCenter_Status As String, _
                                                      sCostCenter_ColumnName As String, _
                                                      lForecastMonth_ID As Long) As Boolean


 'Local Variables
  Dim sSQL As String
  Dim sFilePath As String
  Dim sImport_CostCenter_WorkbookName As String
  Dim sDBPull_NameRange As String
  Dim sAppendTo_Table As String
  Dim sConnectionString As String
  Dim bUpdate_SLT_TRUE_FALSE As Boolean
  Dim lCostCenter As Long

 'ADO Recordset
  Dim rsHierarchy As ADODB.Recordset
  Dim rsNameRange_List As ADODB.Recordset
  Dim Conn_CurrentDb As ADODB.Connection
   
 'Set default value of function
  ab_Import_CostCenterWorksheets_per_tblHierarchy = True
  
  
 On Error GoTo ProcErr


 '------Instantiate objects----------
  Set Conn_CurrentDb = New ADODB.Connection
  Set rsHierarchy = New ADODB.Recordset
  Set rsNameRange_List = New ADODB.Recordset

 'OLEDB connection string to Access's Jet DB
  sConnectionString = CurrentProject.Connection.ConnectionString

 'Connection open
  Conn_CurrentDb.Open sConnectionString
  
  
 ' ** Create LIST OF COST CENTERS TO BE PULLED from the workbooks **
  sSQL = "SELECT tbl_Hierarchy.[" & sCostCenter_ColumnName & "],tbl_Hierarchy.[CC Name], tbl_Hierarchy.CostCenter_FilePath, tbl_Hierarchy.[" & sSLT_ColumnName & "] , tbl_Hierarchy.CCStatus" & vbCrLf
  sSQL = sSQL & " FROM tbl_Hierarchy" & vbCrLf
  
 'If SLT value = ALL then enter LOAD all the cost centers
  If sSLT_Value = "ALL" Then
  
    sSQL = sSQL & " WHERE (((tbl_Hierarchy.CCStatus)='" & sCostCenter_Status & "'))" & vbCrLf
  
  Else
  
    sSQL = sSQL & " WHERE (((tbl_Hierarchy." & sSLT_ColumnName & ")= '" & sSLT_Value & "') " & vbCrLf
    sSQL = sSQL & " AND ((tbl_Hierarchy.CCStatus)='" & sCostCenter_Status & "'))" & vbCrLf
    
  End If
  
  sSQL = sSQL & " ORDER BY tbl_Hierarchy.[" & sCostCenter_ColumnName & "]"
  
 'Open Recordset - Get tblHierarchy data filtered by STL value
  rsHierarchy.Open sSQL, Conn_CurrentDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
  
   
 'Create a list of Name Ranges in the CostCenter workbook
  sSQL = "SELECT tblNameRangeToTable_ImportList.NameRange_Name, tblNameRangeToTable_ImportList.ImportWorksheet_TableName, tblNameRangeToTable_ImportList.NameRange_Active" & vbCrLf
  sSQL = sSQL & " FROM tblNameRangeToTable_ImportList" & vbCrLf
 
 'NOTE: If bImport_byNameRange_YesNo is True then Import ONLY the Forecast Tab
  If bImport_byNameRange_YesNo = True Then

      sSQL = sSQL & " WHERE tblNameRangeToTable_ImportList.NameRange_Name='" & sImport_NameRange_Name & "'" & vbCrLf
      sSQL = sSQL & " ORDER BY tblNameRangeToTable_ImportList.NameRange_Name"
      
  Else
  
      sSQL = sSQL & " WHERE (((tblNameRangeToTable_ImportList.NameRange_Active)=True))" & vbCrLf
      sSQL = sSQL & " ORDER BY tblNameRangeToTable_ImportList.NameRange_Name"
  
  End If
  
  Debug.Print sSQL
  
 'Open recordset - Get tblNameRangeToTable_ImportList
  rsNameRange_List.Open sSQL, Conn_CurrentDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

 'Move to first record
  rsHierarchy.MoveFirst
  rsNameRange_List.MoveFirst
  
  
 '*********************************************************************************************************
 '          >>>> Run subroutine to import the forecast data from the cost center worksheets <<<<
 '
 
 'If SLT value is ALL then use values from tblHierarchy for value, if not ALL then use value provided in the argument sSQL_Value
  If sSLT_Value = "ALL" Then
    bUpdate_SLT_TRUE_FALSE = True
  End If
   
 '** COST CENTER LOOP - List of cost center workbooks to be imported **
  Do While Not rsHierarchy.EOF

     'Use the file path from the CostCenter_FilePath field from tbl_Hierarchy
      If bDefault_FilePath Then
                                        
       'Get FilePath and remove the SLT categories
        sFilePath = rsHierarchy.Fields("CostCenter_FilePath").Value
       
      Else
      
       'Use file path passed by the argument sWorkbookCopy_FilePath and add it the the STL folder
       'NOTE: STL value name needs
        sFilePath = sWorkbookCopy_FilePath & Replace(rsHierarchy.Fields(sSLT_ColumnName).Value, "&", "_") & "\"
      
      End If

     'If bUpdate_SLT_TRUE_FALSE is TRUE then update SLT Column with SLT value
      If bUpdate_SLT_TRUE_FALSE Then
        sSLT_Value = rsHierarchy.Fields(sSLT_ColumnName).Value
      End If
      
    'Get Cost Center
     lCostCenter = rsHierarchy.Fields(sCostCenter_ColumnName).Value
      
    'Get Workbook names from tblHierachy's active Cost Centers
    ' >> Set the Excel Forecast FilePath Name with Worksheet Name <<
    'sImport_CostCenter_WorkbookName = rsHierarchy.Fields(sCostCenter_ColumnName).Value & sCostCenter_WorkbookName_woPreFix
     sImport_CostCenter_WorkbookName = lCostCenter & "_" & "FY" & Right(lFiscalYear, 2) & sCostCenter_WorkbookName_woPreFix

      
     'Show the workbook name
      Debug.Print "Import Workbook Name " & sImport_CostCenter_WorkbookName
      Debug.Print "Start Import " & Now()
  
     'Reset count for loop
      rsNameRange_List.MoveFirst
      
      
     '** NAME RANGE LOOP - Go through each worksheet name range that begins with DBPull **
      Do While Not rsNameRange_List.EOF
      
          'Get NameRange and Import Table the CostCenter forecast data is being inserted into
           sDBPull_NameRange = rsNameRange_List(0).Value
           sAppendTo_Table = rsNameRange_List(1).Value
           
           Debug.Print sDBPull_NameRange
           Debug.Print sAppendTo_Table & vbCrLf
           
          'Run Subroutine
           If Not ac_WriteSQL_Import( _
                                         Conn_CurrentDb, _
                                         sSLT_ColumnName, _
                                         sSLT_Value, _
                                         lFiscalYear, _
                                         sFilePath, _
                                         sImport_CostCenter_WorkbookName, _
                                         sDBPull_NameRange, _
                                         sAppendTo_Table, _
                                         lForecastMonth_ID) Then
           
              ab_Import_CostCenterWorksheets_per_tblHierarchy = False
              
             'Exit sub if there was an error in the ac_WriteSQL_Import subroutine
              GoTo ProcExit
              
            Else
        
              'Set strings to null
               sDBPull_NameRange = vbNullString
               sAppendTo_Table = vbNullString
               
              'If a file is not found then exit this Loop rsNameRange_List
               If PUBLIC_FileNotFound Then
               
                 Exit Do
               
               Else
               
                'Go to the next record in the recordset
                 rsNameRange_List.MoveNext
                 
               End If
         
            End If
            
       Loop
       
     'Set string to null
      sImport_CostCenter_WorkbookName = vbNullString
      
     'Go to the next record in the recordset
      rsHierarchy.MoveNext
    
  Loop


ProcExit:

'NOTE: When connecting to the same database you are in DO NOT CLOSE THE CONNECTION

 'Close Recordset
  rsHierarchy.Close
  Set rsHierarchy = Nothing
  
 'Close Recordset
  rsNameRange_List.Close
  Set rsNameRange_List = Nothing

  Exit Function

ProcErr:

  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit

  Case 91
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 94 'Object not found
    ab_Import_CostCenterWorksheets_per_tblHierarchy = False
    PUBLIC_Table_NotFound = True
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Import Table name, Name Range or other parameter was not provided", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case 424   'Object not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset is already closed
    Resume Next

  Case 3709 'Connection object isnt open
    ab_Import_CostCenterWorksheets_per_tblHierarchy = False
    PUBLIC_Exit_wMessage = False
    
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case -2147217900, -2147217865 'Description Error in SQL
    ab_Import_CostCenterWorksheets_per_tblHierarchy = False
    PUBLIC_Exit_wMessage = False
    
    MsgBox "Description " & Err.Description, vbExclamation + vbOKOnly, "Error handled by FPA Database"
    Resume ProcExit

  Case -2147467259 'Invalid Path
    ab_Import_CostCenterWorksheets_per_tblHierarchy = False
    PUBLIC_DBLocked = True
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation, "Error handled by FPA Database"
    Resume ProcExit

  Case Else
    ab_Import_CostCenterWorksheets_per_tblHierarchy = False
    PUBLIC_Exit_wMessage = False
    
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function
Private Function ac_WriteSQL_Import( _
                                  Conn_CurrentDb As ADODB.Connection, _
                                  sSLT_ColumnName As String, _
                                  sSLT_Value As String, _
                                  lFiscalYear As Long, _
                                  sFilePath As String, _
                                  sImport_CostCenter_WorkbookName As String, _
                                  sDBPull_NameRange As String, _
                                  sAppendto_TableName As String, _
                                  Optional lForecastMonth_ID As Long) As Boolean
                                  

 'ADO Objects
  Dim Conn_Excel_CostCenter As ADODB.Connection
  
 'DAO Objects
  Dim qryDef As DAO.QueryDef


 'Local variables
  Dim sSQL As String
  Dim sSQLWhere As String
  Dim sSQL_Pull_WorksheetData As String
  Dim sConnectionString_ExcelWorksheet As String
  Dim sExcelPathFile As String


 'Set default return value for subroutine
  ac_WriteSQL_Import = True

  
 On Error GoTo ProcErr

'-------------------------------------------------------------------------------
'    ****************  CREATE SQL STATMENT **************************
'
 'Select SQL statment by Name Range passed to this function
  Select Case sDBPull_NameRange
  
  
  Case ""
  
  Case ""
  
  Case ""
  
  Case "DBPull_Forecast"
  
    sSQL_Pull_WorksheetData = ac1_SQL_Forecast(sAppendto_TableName, 4, 15)
    sSQLWhere = " WHERE ((([F2])>0) AND (([F3]) Is Not Null))"
    
  Case "DBPull_ForecastDelta"
  
    sSQL_Pull_WorksheetData = ac1_SQL_WorksheetDelta_Comments("Forecast Delta", False)
    sSQLWhere = " WHERE ((([F2])>0) AND (([F3]) Is Not Null))"
  
  Case "DBPull_TargetDelta"
  
   'NOTE: Sum of the Target Delta worksheets have DBPullTarget Delta as a name range
    sSQL_Pull_WorksheetData = ac1_SQL_WorksheetDelta_Comments("Target Delta", True)
    sSQLWhere = " WHERE ((([F2])>0) AND (([F3]) Is Not Null))"
    
  Case "DBPull_FTE_Roster"
  
    sSQL_Pull_WorksheetData = ac1_SQL_Roster(sAppendto_TableName, 18, 144)
    sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".[F2])<2))" & vbCrLf
    sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F3]) is not null)) OR (((" & sDBPull_NameRange & ".[F4]) is not null))"
    
  Case "DBPull_ETW_Roster"

    sSQL_Pull_WorksheetData = ac1_SQL_Roster(sAppendto_TableName, 18, 144)
    sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".[F2])<2))" & vbCrLf
    sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F4]) is not null)) OR (((" & sDBPull_NameRange & ".[F5]) is not null))"
    
  Case "DBPull_Other"

    sSQL_Pull_WorksheetData = ac1_SQL_Roster(sAppendto_TableName, 6, 19)
    sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".F3) Is Not Null And (" & sDBPull_NameRange & ".F3)<>0))" & vbCrLf
    sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F5]) Is Not Null And (" & sDBPull_NameRange & ".[F5])<>""0""))" & vbCrLf
    sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F6]) Is Not Null))"
    
  Case "DBPull_Depreciation"

    sSQL_Pull_WorksheetData = ac1_SQL_Roster(sAppendto_TableName, 13, 24)
    sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".[F2])<2)) OR (((" & sDBPull_NameRange & ".[F4]) is not null))"
    
  Case Else
   
    'Exit function Name Range does not Match
     ac_WriteSQL_Import = False
     PUBLIC_Table_NotFound = True
     
     MsgBox "Name Range not found. Name Range needs to be changed" & vbCrLf & _
            "See subroutine ac_WriteSQL_Import", vbInformation + vbOKOnly, "Error handled by FPA Database"
            
     GoTo ProcExit
  
  End Select
 

 '--------------------------------------------------------------------------------------------------------------
 '                         ****Excel xlsm file import***
 '
 '  NOTE: Use connection string build from connectionstrings.com under XLSM file
 '        https://www.connectionstrings.com/ace-oledb-12-0/xlsm-files/
 '
 '  ConnectionString: [Excel 12.0 Xml;HDR=NO;IMEX=2;ACCDB=YES;
 '                    DATABASE=C:\Users\PSeie2\Desktop\wip - Financial Forecasting Process Build\Test\].[vbaUpload_Forecast]

 'with named range
  sConnectionString_ExcelWorksheet = "[Excel 12.0 Xml;HDR=NO;IMEX=2;ACCDB=YES;" & vbCrLf & _
          "DATABASE=" & sFilePath & sImport_CostCenter_WorkbookName & "].[" & sDBPull_NameRange & "]"

 'SQL Header Data
  sSQL = "INSERT INTO " & sAppendto_TableName & " ("
  
  
 ' *** Add SQL statement return from function **
 'NOTE: The function returns the body of SQL statement: Then append and select columns
  sSQL = sSQL & Trim(sSQL_Pull_WorksheetData)
   
 'SQL Select Footer Data
  sSQL = sSQL & ", Clng(" & lForecastMonth_ID & ") AS Forecast_MonthID, Clng(" & lFiscalYear & ") AS FiscalYear"
  sSQL = sSQL & ", """ & sSLT_Value & """ AS STL_Value, """ & sSLT_ColumnName & """ AS STL_Column" & vbCrLf
  sSQL = sSQL & ", """ & sFilePath & """ as FilePath, """ & sImport_CostCenter_WorkbookName & """ as FileName" & vbCrLf
  sSQL = sSQL & " FROM " & sConnectionString_ExcelWorksheet & vbCrLf
  sSQL = sSQL & sSQLWhere & vbCrLf
  'sSQL = sSQL & " WHERE ((([F2])>0) AND (([F3]) Is Not Null))" & vbCrLf
  sSQL = sSQL & " ORDER BY CLng(Nz([F2],0))"

 'Trim
  sSQL = Trim(sSQL)


 '*****************        Import and Append Worksheet into a table    *******************
 '
 '  The code to import is nested in the if then statement   adImportExcel(sSQL)
 '  NOTE: if the file isnt found then have a warning msg
 '
 
      If Not ad_Execute_ExcelImport(Conn_CurrentDb, sSQL, sFilePath, sImport_CostCenter_WorkbookName, sDBPull_NameRange) Then
    
        ac_WriteSQL_Import = False
        GoTo ProcExit
    
      End If
  
  
ProcExit:

 'Set strings variables to null
  sSQL = vbNullString
  sSQL_Pull_WorksheetData = vbNullString
  sAppendto_TableName = vbNullString
  sDBPull_NameRange = vbNullString

 ' >>> DO NOT CLOSE THE CONNECTION OBJECT AT THIS POINT  <<<<
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3265 'Table Not Found
    ac_WriteSQL_Import = False
    PUBLIC_Exit_wMessage = False
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    MsgBox "Table " & sAppendto_TableName & " NOT found. Need to make the table!", vbExclamation + vbOKOnly, "Error handled by FPA Database ac_WriteSQL_Import"
    Resume ProcExit

  Case Else
    ac_WriteSQL_Import = False
    PUBLIC_Exit_wMessage = False
    
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit


End Function

Private Function ad_Execute_ExcelImport( _
                                          Conn_CurrentDb As ADODB.Connection, _
                                          sSQL_Import As String, _
                                          sFilePath As String, _
                                          sImport_CostCenter_WorkbookName As String, _
                                          sNameRange As String) As Boolean

 
'Variables
 Dim sSQL As String, sWorksheet As String
 
'Default ad_Execute_ExcelImport to true
 ad_Execute_ExcelImport = True
 PUBLIC_FileNotFound = False

'Check procedural erros with VBA
 On Error GoTo ProcErr
 
 DoCmd.Hourglass True

' Debug.Print "Connection string"
' Debug.Print Conn_CurrentDb.ConnectionString & vbCrLf

 Debug.Print "***  Executed SQL   ***"
 Debug.Print sSQL_Import & vbCrLf & vbCrLf

'   **************** Import Excel Spreadsheet via ADO connection object into a table *************
'
  Conn_CurrentDb.Execute sSQL_Import

ProcExit:
  
  DoCmd.Hourglass False
  
  ' >>> DO NOT CLOSE THE CONNECTION OBJECT AT THIS POINT  <<<<
  
Exit Function


ProcErr:

  'DoCmd.Hourglass False
  
  Select Case Err.Number
  
  Case -2147217900, -2147217904  'Missing SQL Statement and or parameters
    ad_Execute_ExcelImport = False
    PUBLIC_Exit_wMessage = False
    
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case -2147217911 'File Not found
   'Set the public variable to true
    PUBLIC_FileNotFound = True
    PUBLIC_FileNotFound_Name = PUBLIC_FileNotFound_Name & sImport_CostCenter_WorkbookName & vbCrLf
    
    Debug.Print "The file NOT found " & sImport_CostCenter_WorkbookName
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical

    Resume Next
    
  Case -2147217913, -2147467259 'Data Type mismatch / Invalid Path
    ad_Execute_ExcelImport = False
    PUBLIC_Exit_wMessage = False
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    MsgBox "Description " & Err.Description, vbExclamation, "Error handled by FPA Database"
    Resume ProcExit
    
  Case -2147217865 'Could Not Find given name range
    PUBLIC_NameRangeNotFound = True
    PUBLIC_NameRangeNotFound_Name = PUBLIC_namerange_Name & sNameRange & vbCrLf
    
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    
    Resume Next
    
  Case 3265, 3625 'The array of field being input dont match the recordset field names
    PUBLIC_Exit_wMessage = False
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbExclamation, "Error handled by FPA Database"
    Resume Next
    
  Case 3704 'Recordset empty End program to stop more errors
    Resume Next
    
  Case Else
    ad_Execute_ExcelImport = False
    PUBLIC_Exit_wMessage = False
    
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
 Resume ProcExit

End Function
Private Function ac1_SQL_WorksheetDelta_Comments(sDataset_Name As String, _
                                                Optional bExcluded_ExcelColumns_withFormulas As Boolean = True) As String

  Dim sSQL As String
  
  sSQL = sSQL & "CostCenter"
  sSQL = sSQL & ", WorksheetForecast_Row"
  sSQL = sSQL & ", Account"
  sSQL = sSQL & ", Acct_Description" & vbCrLf
  

 'Month
  sSQL = sSQL & ", Month_Actual" & vbCrLf
 
  If Not bExcluded_ExcelColumns_withFormulas Then
    sSQL = sSQL & ", Month_CompareNumber, Month_Delta, Month_Prct"
  End If
  
  sSQL = sSQL & ", Month_Delta_Comments" & vbCrLf

  
 'Qtr
  sSQL = sSQL & ", Qtr_Actual" & vbCrLf
  
  If Not bExcluded_ExcelColumns_withFormulas Then
      sSQL = sSQL & ", Qtr_CompareNumber, Qtr_Delta, Qtr_Prct"
  End If

  sSQL = sSQL & ", Qtr_Delta_Comments" & vbCrLf
  
  
 'Year
  sSQL = sSQL & ", Year_Actual" & vbCrLf
  
  If Not bExcluded_ExcelColumns_withFormulas Then
    sSQL = sSQL & ", Year_CompareNumber, Year_Delta, Year_Prct"
  End If
  
  sSQL = sSQL & ", Year_Delta_Comments" & vbCrLf
  sSQL = sSQL & ", Comparison_DataSet_Name" & vbCrLf
  
 'SQL Append Footer Data
  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
  sSQL = sSQL & ", FiscalYear" & vbCrLf
  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
  sSQL = sSQL & ", FilePath, FileName"
  sSQL = sSQL & ") " & vbCrLf
  
 'SQL Select Header Data
  sSQL = sSQL & " Select"
  sSQL = sSQL & " CLng(IIf(IsError([F1]),1,[F1])) AS CostCenter" & vbCrLf
  sSQL = sSQL & ", Clng(NZ([F2],0)) as WorksheetForecast_Row, CDbl(NZ([F3],1)) as Account, Trim([F4]) as Acct_Description" & vbCrLf
  
 'Month
  sSQL = sSQL & ", CDbl(NZ([F6],0)*1000) as Month_Actual" & vbCrLf
  
 'Include if Forecast Delta
  If Not bExcluded_ExcelColumns_withFormulas Then
    sSQL = sSQL & ", CDbl(NZ([F7],0)*1000) as Month_CompareNumber, CDbl(NZ([F10],0)*1000) as Month_Delta, CDbl(NZ([F11],0)) as Month_Prct" & vbCrLf
  End If
    
  sSQL = sSQL & ", Trim([F13]) as Month_Delta_Comments" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F15],0)*1000) as Qtr_Actual" & vbCrLf
  
  
 'Qtr
  If Not bExcluded_ExcelColumns_withFormulas Then
      sSQL = sSQL & ", CDbl(NZ([F16],0)*1000) as Qtr_CompareNumber, CDbl(NZ([F19],0)*1000) as Qtr_Delta, CDbl(NZ([F20],0)) as Qtr_Prct" & vbCrLf
  End If
  
  sSQL = sSQL & ", Trim(F22) as Qtr_Delta_Comments" & vbCrLf
  
  
 'Year
  sSQL = sSQL & ", CDbl(NZ([F24],0)*1000) as Year_Actual" & vbCrLf
 
 'Include if Forecast Delta
  If Not bExcluded_ExcelColumns_withFormulas Then
    sSQL = sSQL & ", CDbl(NZ([F25],0)*1000) as Year_CompareNumber, CDbl(NZ([F28],0)*1000) as Year_Delta, CDbl(NZ([F29],0)) as Year_Prct" & vbCrLf
  End If


  sSQL = sSQL & ", Trim(F31) as Year_Delta_Comments" & vbCrLf
  sSQL = sSQL & ", """ & sDataset_Name & """ as Comparison_Dataset_Name"
  
  
  ac1_SQL_WorksheetDelta_Comments = sSQL
  
  sSQL = vbNullString

End Function
Private Function ac1_SQL_Forecast(Optional sAppendto_TableName As String = "tblImport_WorksheetForecast", _
                                 Optional lBeginColumns_wCarraigeReturn As Long = 4, _
                                Optional lEndColumns_wCarraigeReturn As Long = 15) As String

  Dim sSQL As String
  
  
 'NOTE: the function needs to be set to "SQL_AppendTo_TableColumns" to columns
  sSQL = fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, lBeginColumns_wCarraigeReturn, lEndColumns_wCarraigeReturn, _
                                          "SQL_AppendTo_TableColumns", 7)

 'SQL Append Footer Data
  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
  sSQL = sSQL & ", FiscalYear" & vbCrLf
  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
  sSQL = sSQL & ", FilePath, FileName"
  sSQL = sSQL & ") " & vbCrLf
  
 'NOTE: the function needs to be set to "SQL_PullFrom_WorksheetColumns" to columns
  ac1_SQL_Forecast = sSQL & fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, lBeginColumns_wCarraigeReturn, lEndColumns_wCarraigeReturn, _
                                                          "SQL_PullFrom_WorksheetColumns", 7)
  
  sSQL = vbNullString


End Function
Private Function ac1_SQL_Roster(sAppendto_TableName As String, _
                                lBeginColumns_wCarraigeReturn As Long, _
                                lEndColumns_wCarraigeReturn As Long) As String

  Dim sSQL As String

 'NOTE: the function needs to be set to "SQL_AppendTo_TableColumns" to columns
  sSQL = sSQL & fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, lBeginColumns_wCarraigeReturn, lEndColumns_wCarraigeReturn, _
                                                "SQL_AppendTo_TableColumns", 7)

 'SQL Append Footer Data
  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
  sSQL = sSQL & ", FiscalYear" & vbCrLf
  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
  sSQL = sSQL & ", FilePath, FileName"
  sSQL = sSQL & ") " & vbCrLf
  
  'NOTE: the function needs to be set to "SQL_PullFrom_WorksheetColumns" to columns
  ac1_SQL_Roster = sSQL & fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, lBeginColumns_wCarraigeReturn, lEndColumns_wCarraigeReturn, _
                                                          "SQL_PullFrom_WorksheetColumns", 7)
  
  sSQL = vbNullString

End Function

Private Function fn_SQL_Append_Worksheet_toTable(sAppendto_TableName As String, _
                                               lStartColumn_AddCarriageReturns As Long, _
                                               lEndColumn_AddCarriageReturns As Long, _
                                               sSQL_Statment_TableOrWorksheet, Optional iRemove_Ending_ImportStatus_Fields = 7)

 '*************************************************************************************************************
 '*
 '* Dynamical create SQL append statment from target worksheet to target table based on the table structure
 '*
 '*
 '*
 '*
 
 'Create an append SQL statment with the table name defintion
  Dim sSQL As String
  Dim sStr As String
  Dim sQueryDef As String
  Dim iStartColumn_Number As Integer
  Dim lTable_FldNumber As Long
  Dim lWorksheet_Column As Long
  Dim bAccount_Columns As Boolean
   
 'DAO objects
  Dim db As DAO.Database
  Dim tblDef As DAO.TableDef
  Dim Table_Field As DAO.Field
  Dim qryDef As DAO.QueryDef

 'Get the append tp tables structure
  Set db = Application.CurrentDb
  Set tblDef = db.TableDefs(sAppendto_TableName)
  
'  sQueryDef = "app_001_Worksheet_SQLQuery"
'  Set qryDef = db.QueryDefs(sQueryDef)
  
 'Check if the SQL statement for the Table Columns or Worksheet Columns
  If sSQL_Statment_TableOrWorksheet = "SQL_AppendTo_TableColumns" Then
  
'     'Begging SQL Statement
'      sSQL = "Select ("
  
     'Get field names
      For Each Table_Field In tblDef.Fields
      
         'Get the field number. The first Ordinal Position of the field is 0
          lTable_FldNumber = Table_Field.OrdinalPosition
        
         'Skip first key Table_Field and the last date field
          If lTable_FldNumber > 0 And lTable_FldNumber < tblDef.Fields.Count - iRemove_Ending_ImportStatus_Fields Then

             'Debug.Print "# " & Table_Field.OrdinalPosition & " Name " & Table_Field.Name & " Type " & Table_Field.Type
              
             'Skip Empty column ie Skip columns with Integer Data Type ie 3
              If Table_Field.Type <> 3 Then
          
                 'Add carriage return after the following columns
                  If Right(tblDef.Fields(lTable_FldNumber).Name, 4) = "June" _
                      Or Right(tblDef.Fields(lTable_FldNumber).Name, 7) = "October" _
                      Or Right(tblDef.Fields(lTable_FldNumber).Name, 7) = "January" _
                      Or Right(tblDef.Fields(lTable_FldNumber).Name, 9) = "FTE_UseMe" _
                      Or lTable_FldNumber < lStartColumn_AddCarriageReturns Or lTable_FldNumber > lEndColumn_AddCarriageReturns Then
                  
                      sSQL = sSQL & vbCrLf & Table_Field.Name & ", "
                    
                  Else
                  
                      sSQL = sSQL & Table_Field.Name & ", "
                    
                  End If
              
              End If
              
            'If lTable_FldNumber + 1 >= tblDef.Fields.Count Then
            '   Exit For
            'End If
            
          End If
          
      Next
      
      sSQL = Left(Trim(sSQL), Len(Trim(sSQL)) - 1) & vbCrLf
      
      
  Else
  
  '*** Worksheet Columns ***
  
     'Set default value if the CostCenter Worksheet and IDAccount have been completed
      bAccount_Columns = True
      
     'Name Range column number
      lWorksheet_Column = 0
  
     'Begging SQL Statement
      sSQL = "Select "
  
     'Get field names
      For Each Table_Field In tblDef.Fields
      
         'Get the field number. The first Ordinal Position of the field is 0
          lTable_FldNumber = Table_Field.OrdinalPosition

          
         'Skip first key Table_Field in the TABLE (NOTE: lTable_FldNumber represent Table Field Numbers)
          If lTable_FldNumber > 0 And lTable_FldNumber < tblDef.Fields.Count - iRemove_Ending_ImportStatus_Fields Then

                        
              'Increment the Worksheet Column number by 1
               lWorksheet_Column = lWorksheet_Column + 1
                          
                          
              'Skip Empty column ie Skip columns with Integer Data Type ie 3
               If Table_Field.Type <> 3 Then
               
                 'Debug.Print "# " & lWorksheet_Column & " Name " & Table_Field.Name & " Type " & Table_Field.Type
                
                 '---------------------------------------------------------------------------------------------------------------------
                 '
                 '  Create Custome logical function if required my Worksheet row number
                 '  NOTE:
                 '
                 '
                 
                  Select Case lWorksheet_Column
                          
                  Case 1 'CostCenter
                  
                    sSQL = sSQL & "CLng(IIf(IsError([F" & lWorksheet_Column & "]),1,[F" & lWorksheet_Column & "])) AS " & Table_Field.Name & ", "
                  
                  Case Else
                  
                    bAccount_Columns = False
                     
                  End Select
                 
                 'Create all the other columns that are not the Account Columns ie CostCenter, WorksheetID and Account
                  If Not bAccount_Columns Then
                              
                     'Add carriage return after the following columns
                      If Right(Table_Field.Name, 4) = "June" _
                           Or Right(Table_Field.Name, 7) = "October" _
                           Or Right(Table_Field.Name, 7) = "January" _
                           Or Right(Table_Field.Name, 3) = "May" _
                           Or Right(Table_Field.Name, 9) = "FTE_UseMe" _
                           Or lTable_FldNumber < lStartColumn_AddCarriageReturns Or lTable_FldNumber > lEndColumn_AddCarriageReturns Then
                
                          'Don't Add Ending Carraige return the SQL Statement
                           If lTable_FldNumber < tblDef.Fields.Count - iRemove_Ending_ImportStatus_Fields - 1 Then
                           
                              sSQL = sSQL & vbCrLf
                              
                           
                           End If

                       End If
                      
                     '--------------------------------------------------------------------------------
                     '   >>>>>> SET FIELD TYPE BASE ON THE TABLE THAT IS BEING IMPORTED INTO !!!!
                     
                      Select Case Table_Field.Type
                      
                      Case 3 'Integer
                      
                        ' *** Skip columns with Interger Data Type are empty columns in the worksheet.Dont need to add data! ***
                        Debug.Print "This field should be skipped"
                      
                      Case 4 'long
                      
                        sSQL = sSQL & "CLng(IIf(IsError([F" & lWorksheet_Column & "]),0,[F" & lWorksheet_Column & "])) AS " & Table_Field.Name & ", "
                        
                       ' sSQL = sSQL & "CLng(IsEmpty([F" & lWorksheet_Column & "])) AS " & Table_Field.Name & ", "
                       ' sSQL = sSQL & "CLng(NZ(Trim([F" & lWorksheet_Column & "]),0)) AS " & Table_Field.Name & ", "
                       ' sSQL = sSQL & " IIf(Len([F" & lWorksheet_Column & "])=0,CLng(0),CLng(Nz([F" & lWorksheet_Column & "],0))) AS " & Table_Field.Name & ", "
                      
                      Case 7 'double
                      
                        sSQL = sSQL & "CDbl(NZ(Trim([F" & lWorksheet_Column & "]),0)) AS " & Table_Field.Name & ", "
                          
                      Case 8 'Date
                      
                        sSQL = sSQL & "CDate(DateValue(NZ(Trim([F" & lWorksheet_Column & "]),""1/1/1999""))) AS " & Table_Field.Name & ", "
                      
                      Case 10 'String
                      
                        sSQL = sSQL & "Trim([F" & lWorksheet_Column & "]) AS " & Table_Field.Name & ", "
                      
                      Case Else
    
                      End Select
                      
                  End If
                  
                 'Completed SQL statement
                  'Debug.Print "Completed SQL Statment " & vbCrLf
                 
              End If

          End If
        
      Next
      
     'Remove ending comma from SQL statement
      sSQL = Left(Trim(sSQL), Len(Trim(sSQL)) - 1)
  
  End If
  

 ' Debug.Print sSQL
  
 'UPDATE QUERY DEFINITION
  'qryDef.Sql = sSQL
  
 'Return SQL string
  fn_SQL_Append_Worksheet_toTable = sSQL
  
End Function
Private Sub Run_SQL_Append_Worksheet_toTable()
 
 'DAO objects
  Dim db As DAO.Database
  Dim qryDef As DAO.QueryDef

 'Local variable
  Dim sSQL As String
  Dim sFilePath As String
  Dim sAppendto_TableName As String
  Dim sDBPull_NameRange As String
  Dim sSQLWhere As String
  
'  Set db = Application.CurrentDb
'  Set qryDef = db.QueryDefs("app_001_TEST_Worksheet_SQLQuery")
  
 'Set Append table and Name Range
  sFilePath = "\\NKE-WIN-NAS-P22\GOT_Budget\Fin Planning\"
  sImport_CostCenter_WorkbookName = "Test Forecast Workbook v16.xlsm"
  sAppendto_TableName = "tblImport_WorksheetFTERoster"
  sDBPull_NameRange = "DBPull_FTE_Roster"
  
 'Set Where statement FTE
  sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".[F2])<2))" & vbCrLf
  sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F3]) is not null)) OR (((" & sDBPull_NameRange & ".[F4]) is not null))" & vbCrLf
 
 'ETW
'  sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".[F2])<2))" & vbCrLf
'  sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F4]) is not null)) OR (((" & sDBPull_NameRange & ".[F5]) is not null))" & vbCrLf

  'Other
'  sSQLWhere = " WHERE (((" & sDBPull_NameRange & ".F3) Is Not Null And (" & sDBPull_NameRange & ".F3)<>0))" & vbCrLf
'  sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F4]) Is Not Null And (" & sDBPull_NameRange & ".[F4])<>""0""))" & vbCrLf
'  sSQLWhere = sSQLWhere & " OR (((" & sDBPull_NameRange & ".[F5]) Is Not Null))"
  
 'Begging SQL Statement
  sSQL = "INSERT INTO " & sAppendto_TableName & "("
  
  sSQL = sSQL & fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, 18, 141, "SQL_AppendTo_TableColumns", 7)

 'SQL Append Footer Data
  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
  sSQL = sSQL & ", FilePath, FileName"
  sSQL = sSQL & ") " & vbCrLf
  
  sSQL = sSQL & fn_SQL_Append_Worksheet_toTable(sAppendto_TableName, 18, 141, "SQL_PullFrom_WorksheetColumns", 7)

  sSQL = sSQL & "," & vbCrLf ' Add ending comma that is removed in SQL_Append function
  sSQL = sSQL & " CLng(308) as ForecastMonth_ID, " & vbCrLf
  sSQL = sSQL & """FIN"" AS STL_Value, ""SLT5"" AS STL_Column, """ & sFilePath & """ AS FilePath, """ & sImport_CostCenter_WorkbookName & """ AS FileName" & vbCrLf
  sSQL = sSQL & " FROM [Excel 12.0 Xml;HDR=NO;IMEX=2;ACCDB=YES;" & vbCrLf
  sSQL = sSQL & " DATABASE=" & sFilePath & sImport_CostCenter_WorkbookName & "].[" & sDBPull_NameRange & "]" & vbCrLf
  'sSQL = sSQL & " DATABASE=\\NKE-WIN-NAS-P22\GOT_Budget\Fin Planning\FIN\100613_Forecast.xlsm].DBPull_FTE_Roster" & vbCrLf
  
  sSQL = sSQL & sSQLWhere
  sSQL = sSQL & "ORDER BY CLng(Nz([F2],0));"
  
  Debug.Print sSQL
  
'  qryDef.Sql = sSQL
    
End Sub

Private Function zzObsolete_ac1_SQL_FTE_Roster(Optional sAppendto_TableName As String = "tblImport_WorksheetFTERoster") As String

  Dim sSQL As String
  
 'Insert Colums
  sSQL = sSQL & "FTE_CostCenter"
  sSQL = sSQL & ", FTE_Row"
  sSQL = sSQL & ", FTE_PreviousPosition_FilledBy, FTE_Name, FTE_Business_Title, FTE_JobCode, FTE_Employee_ID, FTE_PositionID"
  sSQL = sSQL & ", FTE_Band, FTE_Budgeted_WageBase, FTE_Current_WageBase, FTE_Status, FTE_StartMonth, FTE_EndMonth, FTE_OpenFilled, FTE_Comments" & vbCrLf
  
  
  sSQL = sSQL & ", FTE_Q1_P2P, FTE_Q1_Cap, FTE_Q2_P2P, FTE_Q2_Cap, FTE_Q3_P2P, FTE_Q3_Cap, FTE_Q4_P2P, FTE_Q4_Cap" & vbCrLf
  
  
  sSQL = sSQL & ", FTE_TotalCost_June, FTE_TotalCost_July, FTE_TotalCost_August, FTE_TotalCost_September, FTE_TotalCost_October, FTE_TotalCost_November, FTE_TotalCost_December" & vbCrLf
  sSQL = sSQL & ", FTE_TotalCost_January, FTE_TotalCost_February, FTE_TotalCost_March, FTE_TotalCost_April, FTE_TotalCost_May, FTE_TotalCost_FiscalYear" & vbCrLf
 
  sSQL = sSQL & ", FTE_UseMe, FTE_StartMonth_Number, FTE_EndMonth_Number"
 
  sSQL = sSQL & ", FTE_HrsWorked_June, FTE_HrsWorked_July, FTE_HrsWorked_August, FTE_HrsWorked_September, FTE_HrsWorked_October, FTE_HrsWorked_November, FTE_HrsWorked_December" & vbCrLf
  sSQL = sSQL & ", FTE_HrsWorked_January, FTE_HrsWorked_February, FTE_HrsWorked_March, FTE_HrsWorked_April, FTE_HrsWorked_May, FTE_HrsWorked_FiscalYear" & vbCrLf

  sSQL = sSQL & ", FTE_Salary_June, FTE_Salary_July, FTE_Salary_August, FTE_Salary_September, FTE_Salary_October, FTE_Salary_November, FTE_Salary_December" & vbCrLf
  sSQL = sSQL & ", FTE_Salary_January, FTE_Salary_February, FTE_Salary_March, FTE_Salary_April, FTE_Salary_May, FTE_Salary_Merit_Increase" & vbCrLf

  sSQL = sSQL & ", FTE_PSP_June, FTE_PSP_July, FTE_PSP_August, FTE_PSP_September, FTE_PSP_October, FTE_PSP_November, FTE_PSP_December" & vbCrLf
  sSQL = sSQL & ", FTE_PSP_January, FTE_PSP_February, FTE_PSP_March, FTE_PSP_April, FTE_PSP_May, FTE_PSP_FiscalYear" & vbCrLf
 
  sSQL = sSQL & ", FTE_Fringe_June, FTE_Fringe_July, FTE_Fringe_August, FTE_Fringe_September, FTE_Fringe_October, FTE_Fringe_November, FTE_Fringe_December" & vbCrLf
  sSQL = sSQL & ", FTE_Fringe_January, FTE_Fringe_February, FTE_Fringe_March, FTE_Fringe_April, FTE_Fringe_May, FTE_Fringe_FiscalYear" & vbCrLf
 
  sSQL = sSQL & ", FTE_Allocation_June, FTE_Allocation_July, FTE_Allocation_August, FTE_Allocation_September, FTE_Allocation_October, FTE_Allocation_November, FTE_Allocation_December" & vbCrLf
  sSQL = sSQL & ", FTE_Allocation_January, FTE_Allocation_February, FTE_Allocation_March, FTE_Allocation_April, FTE_Allocation_May, FTE_Allocation_FiscalYear" & vbCrLf

'  sSQL = sSQL & ", PayToPlay_June, PayToPlay_July, PayToPlay_August, PayToPlay_September, PayToPlay_October, PayToPlay_November, PayToPlay_December" & vbCrLf
'  sSQL = sSQL & ", PayToPlay_January, PayToPlay_February, PayToPlay_March, PayToPlay_April, PayToPlay_May" & vbCrLf
  
'  sSQL = sSQL & ", Capitalization_June, Capitalization_July, Capitalization_August, Capitalization_September, Capitalization_October, Capitalization_November, Capitalization_December" & vbCrLf
'  sSQL = sSQL & ", Capitalization_January, Capitalization_February, Capitalization_March, Capitalization_April, Capitalization_May" & vbCrLf
  
  
 'SQL Append Footer Data
  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
  sSQL = sSQL & ", FilePath, FileName"
  sSQL = sSQL & ") " & vbCrLf


 'SQL Select Header Data
  sSQL = sSQL & " Select"
  sSQL = sSQL & " CLng(IIf(IsError([F1]),1,[F1])) AS CostCenter" & vbCrLf
  sSQL = sSQL & ", Clng(NZ([F2],0)) as WorksheetForecast_Row" & vbCrLf
  
 'SQL Select body Data
  sSQL = sSQL & ", Trim([F3]) as PreviousPosition, Trim([F4]) As FTE_Name, Trim([F5]) as Business_Title, Trim([F12]) as Job_Code, Trim([F7]) as Employee_ID, Trim([F8]) as Position_ID" & vbCrLf
  sSQL = sSQL & ", Trim([F9]) As FTE_Band, CDbl(NZ([F10],0)) As FTE_Budgeted, CDbl(NZ([F11],0)) as FTE_Current, Trim([F12]) as FTE_Status" & vbCrLf
  sSQL = sSQL & ", Trim([F13]) as StartMonth_Name, Trim([F14]) as EndMonth_Name, Trim([F15]) as OpenFilled, Trim([F16]) as FTE_Comments" & vbCrLf

 'Data Sets
  sSQL = sSQL & ", CDbl(NZ([F17],0)) as FTE_Q1_P2P, CDbl(NZ([F18],0)) as FTE_Q1_Cap, CDbl(NZ([F19],0)) as FTE_Q2_P2P, CDbl(NZ([F20],0)) as FTE_Q2_Cap" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F21],0)) as FTE_Q3_P2P, CDbl(NZ([F22],0)) as FTE_Q3_Cap, CDbl(NZ([F23],0)) as FTE_Q4_P2P, CDbl(NZ([F24],0)) as FTE_Q4_Cap" & vbCrLf
  
  sSQL = sSQL & ", CDbl(NZ([F27],0)) as FTE_TotalCost_June, CDbl(NZ([F28],0)) as FTE_TotalCost_July, CDbl(NZ([F29],0)) as FTE_TotalCost_August, CDbl(NZ([F30],0)) as FTE_TotalCost_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F31],0)) as FTE_TotalCost_October, CDbl(NZ([F32],0)) as FTE_TotalCost_November, CDbl(NZ([F33],0)) as FTE_TotalCost_December, CDbl(NZ([F34],0)) as FTE_TotalCost_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F35],0)) as FTE_TotalCost_February, CDbl(NZ([F36],0)) as FTE_TotalCost_March, CDbl(NZ([F37],0)) as FTE_TotalCost_April, CDbl(NZ([F38],0)) as FTE_TotalCost_May, CDbl(NZ([F39],0)) as FTE_TotalCost_FiscalYear" & vbCrLf

  sSQL = sSQL & ", CDbl(NZ([F40],0)) as FTE_UseMe, CLng(NZ([F41],0)) as FTE_StartMonth_Number, CLng(NZ([F42],0)) as FTE_EndMonth_Number"
  
  sSQL = sSQL & ", CDbl(NZ([F43],0)) as FTE_HrsWorked_June, CDbl(NZ([F44],0)) as FTE_HrsWorked_July, CDbl(NZ([F45],0)) as FTE_HrsWorked_August, CDbl(NZ([F46],0)) as FTE_HrsWorked_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F47],0)) as FTE_HrsWorked_October, CDbl(NZ([F48],0)) as FTE_HrsWorked_November, CDbl(NZ([F49],0)) as FTE_HrsWorked_December, CDbl(NZ([F50],0)) as FTE_HrsWorked_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F51],0)) as FTE_HrsWorked_February, CDbl(NZ([F52],0)) as FTE_HrsWorked_March, CDbl(NZ([F53],0)) as FTE_HrsWorked_April, CDbl(NZ([F54],0)) as FTE_HrsWorked_May, CDbl(NZ([F55],0)) as FTE_HrsWorked_FiscalYear" & vbCrLf

  sSQL = sSQL & ", CDbl(NZ([F57],0)) as FTE_Salary_June, CDbl(NZ([F58],0)) as FTE_Salary_July, CDbl(NZ([F59],0)) as FTE_Salary_August, CDbl(NZ([F60],0)) as FTE_Salary_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F61],0)) as FTE_Salary_October, CDbl(NZ([F62],0)) as FTE_Salary_November, CDbl(NZ([F63],0)) as FTE_Salary_December, CDbl(NZ([F64],0)) as FTE_Salary_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F65],0)) as FTE_Salary_February, CDbl(NZ([F66],0)) as FTE_Salary_March, CDbl(NZ([F67],0)) as FTE_Salary_April, CDbl(NZ([F68],0)) as FTE_Salary_May, CDbl(NZ([F69],0)) as FTE_Salary_Merit_Increase" & vbCrLf

  sSQL = sSQL & ", CDbl(NZ([F71],0)) as FTE_PSP_June, CDbl(NZ([F72],0)) as FTE_PSP_July, CDbl(NZ([F73],0)) as FTE_PSP_August, CDbl(NZ([F74],0)) as FTE_PSP_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F75],0)) as FTE_PSP_October, CDbl(NZ([F76],0)) as FTE_PSP_November, CDbl(NZ([F77],0)) as FTE_PSP_December, CDbl(NZ([F78],0)) as FTE_PSP_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F79],0)) as FTE_PSP_February, CDbl(NZ([F80],0)) as FTE_PSP_March, CDbl(NZ([F81],0)) as FTE_PSP_April, CDbl(NZ([F82],0)) as FTE_PSP_May, CDbl(NZ([F83],0)) as FTE_PSP_FiscalYear" & vbCrLf

  sSQL = sSQL & ", CDbl(NZ([F85],0)) as FTE_Fringe_June, CDbl(NZ([F86],0)) as FTE_Fringe_July, CDbl(NZ([F87],0)) as FTE_Fringe_August, CDbl(NZ([F88],0)) as FTE_Fringe_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F89],0)) as FTE_Fringe_October, CDbl(NZ([F90],0)) as FTE_Fringe_November, CDbl(NZ([F91],0)) as FTE_Fringe_December, CDbl(NZ([F92],0)) as FTE_Fringe_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F93],0)) as FTE_Fringe_February, CDbl(NZ([F94],0)) as FTE_Fringe_March, CDbl(NZ([F95],0)) as FTE_Fringe_April, CDbl(NZ([F96],0)) as FTE_Fringe_May, CDbl(NZ([F97],0)) as FTE_Fringe_FiscalYear" & vbCrLf

  sSQL = sSQL & ", CDbl(NZ([F99],0)) as FTE_Allocation_June, CDbl(NZ([F100],0)) as FTE_Allocation_July, CDbl(NZ([F101],0)) as FTE_Allocation_August, CDbl(NZ([F102],0)) as FTE_Allocation_September" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F103],0)) as FTE_Allocation_October, CDbl(NZ([F104],0)) as FTE_Allocation_November, CDbl(NZ([F105],0)) as FTE_Allocation_December, CDbl(NZ([F106],0)) as FTE_Allocation_January" & vbCrLf
  sSQL = sSQL & ", CDbl(NZ([F107],0)) as FTE_Allocation_February, CDbl(NZ([F108],0)) as FTE_Allocation_March, CDbl(NZ([F109],0)) as FTE_Allocation_April, CDbl(NZ([F110],0)) as FTE_Allocation_May, CDbl(NZ([F111],0)) as FTE_Allocation_FiscalYear" & vbCrLf

'  sSQL = sSQL & ", CDbl(NZ([F45],0)) as PayToPlay_June, CDbl(NZ([F46],0)) as PayToPlay_July, CDbl(NZ([F47],0)) as PayToPlay_August, CDbl(NZ([F48],0)) as PayToPlay_September" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F49],0)) as PayToPlay_October, CDbl(NZ([F50],0)) as PayToPlay_November, CDbl(NZ([F51],0)) as PayToPlay_December, CDbl(NZ([F52],0)) as PayToPlay_January" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F53],0)) as PayToPlay_February, CDbl(NZ([F54],0)) as PayToPlay_March, CDbl(NZ([F55],0)) as PayToPlay_April, CDbl(NZ([F56],0)) as PayToPlay_May" & vbCrLf
'
'  sSQL = sSQL & ", CDbl(NZ([F59],0)) as Capitalization_June, CDbl(NZ([F60],0)) as Capitalization_July, CDbl(NZ([F61],0)) as Capitalization_August, CDbl(NZ([F62],0)) as Capitalization_September" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F63],0)) as Capitalization_October, CDbl(NZ([F64],0)) as Capitalization_November, CDbl(NZ([F65],0)) as Capitalization_December, CDbl(NZ([F66],0)) as Capitalization_January" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F67],0)) as Capitalization_February, CDbl(NZ([F68],0)) as Capitalization_March, CDbl(NZ([F69],0)) as Capitalization_April, CDbl(NZ([F70],0)) as Capitalization_May" & vbCrLf
   
  zzObsolete_ac1_SQL_Roster = sSQL
  
  sSQL = vbNullString


End Function




Private Sub zzObsolete_SQL_app()

'  Dim sSQL As String
'  Dim sConnectionString_ExcelWorksheet As String
'
'  sConnectionString_ExcelWorksheet = "Provider=Microsoft.ACE.OLEDB.12.0;" & vbCrLf & _
'                                    "Data Source=" & sFilePath & sImport_CostCenter_WorkbookName & ";" & vbCrLf & _
'                                    "Extended Properties=""Excel 12.0 Macro;HDR=NO"""


  '  sSQL = sSQL & ", h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13, h14, h15, h16, h17, h18, h19, h20, h21, h22, h23, h24 )"
'  sSQL = sSQL & "SELECT tblAsset.Asset, tblAsset.Asset_Delivery_Location, " & vbCrLf
'  sSQL = sSQL & "IIf(Left(Trim([A.F5]),2)=""DA"",""DA"",IIf(Trim(Left([A.F5],2))=""HA"" Or Trim([A.F5])=""Estimate"",""HA"",""DA"")) AS Forecast_Generation, " & vbCrLf
'
'  ' sSQL = sSQL & "IIf(Trim([A.F5])=""Estimate"",""HA Estimate"",IIf(Trim([A.F5])=""3 Tier Forecast Generation"",""DA 3 Tier Forecast Generation"",[F5]))AS Forecast_Generation_Long_Desc, " & vbCrLf
'
'  sSQL = sSQL & "Trim(Left(IIf(Trim([A.F5])=""Estimate"",""HA Estimate"",IIf(Trim([A.F5])=""3 Tier Forecast Generation"",""DA 3 Tier"",[F5])),12)) AS Forecast_Generation_Long_Desc, " & vbCrLf
'  sSQL = sSQL & """ACES"" AS Data_Source, " & vbCrLf
'  sSQL = sSQL & "DateSerial(Left([A.F3],4),Mid([A.F3],5,2),Right([A.F3],2)) AS Delivery_Date, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F7],0)) AS h1, CLng(NZ(A.[F8],0)) AS h2, CLng(NZ(A.[F9],0)) AS h3, CLng(NZ(A.[F10],0)) AS h4, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F11],0)) AS h5, CLng(NZ(A.[F12],0)) AS h6, CLng(NZ(A.[F13],0)) AS h7, CLng(NZ(A.[F14],0)) AS h8, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F15],0)) AS h9, CLng(NZ(A.[F16],0)) AS h10, CLng(NZ(A.[F17],0)) AS h11, CLng(NZ(A.[F18],0)) AS h12, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F19],0)) AS h13, CLng(NZ(A.[F20],0)) AS h14, CLng(NZ(A.[F21],0)) AS h15, CLng(NZ(A.[F22],0)) AS h16, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F23],0)) AS h17, CLng(NZ(A.[F24],0)) AS h18, CLng(NZ(A.[F25],0)) AS h19, CLng(NZ(A.[F26],0)) AS h20, " & vbCrLf
'  sSQL = sSQL & "CLng(NZ(A.[F27],0)) AS h21, CLng(NZ(A.[F28],0)) AS h22, CLng(NZ(A.[F29],0)) AS h23, CLng(NZ(A.[F30],0)) AS h24 " & vbCrLf
'
'  sSQL = sSQL & "FROM " & sExcel_Forecast & " as A INNER JOIN tblAsset ON A.F4 = tblAsset.Asset_ACES_Name " & vbCrLf

  'sSQL = sSQL & "WHERE (((tblAsset.Asset)=""KLONDIKE3""))"
  
End Sub
Private Function zzObsolete_fnAppendTable(sQueryDefinition_Name As String, Optional sTableName)

  Dim qd As QueryDef
  Dim sSQL As String
  
  Select Case sQueryDefinition_Name
  
  Case "app_tblImport_Worksheet_into_tblForecastMonth"
  
   'Create the ForecastMonth table name by concatenating tbl_ with the number entered in the input box
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
    
    sSQL = Right(sSQL, Len(sSQL) - 27)
    sSQL = "INSERT INTO tblForecast_" & TempVars("ForecastMonth_ID").Value & " " & sSQL
    
  Case "app_tblImport_ExcelBudget_via_SQLquery"
  
   'Create the ForecastMonth table name by concatenating tbl_ with the number entered in the input box
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
  
    sSQL = Left(sSQL, InStr(1, sSQL, "q4"))
    'sSQL = "INSERT INTO tblForecast_" & TempVars("ForecastMonth_ID").Value & " " & sSQL
    
    
    sSQL = sSQL & Right(sSQL, Len(sSQL) - InStr(1, sSQL, "Z:\"))
  
'    sSQL = sSQL & ", ""Z:\Fin Planning\SAP and Excel Import Data"" AS FilePath, ""BudgetData_Placeholder.xlsx"" AS FileName"
'    sSQL = sSQL & ""
    
  Case Else
  
    Exit Function
  
  End Select
  
  Debug.Print sSQL
  
 ' >>>>> EXECUTE QUERY  <<<<<<<
  CurrentProject.Connection.Execute sSQL
  
End Function
Private Function zzObsolete_SQL_Forecast() As String

'  sSQL = sSQL & "CostCenter"
'  sSQL = sSQL & ", WorksheetForecast_Row"
'  sSQL = sSQL & ", Account"
'  sSQL = sSQL & ", SAP_Account"
'  sSQL = sSQL & ", Acct_Description" & vbCrLf
'
' 'SQL Append to Body
'  sSQL = sSQL & ", June, July, August" & vbCrLf
'  sSQL = sSQL & ", September, October" & vbCrLf
'  sSQL = sSQL & ", November" & vbCrLf
'
' 'NOTE: Dec is a reserved word. It can NOT be used as a abbreviation for Decemeber
' '      Stack Overflow https://stackoverflow.com/questions/44373372/excel-sql-insert-into-query-syntax-error
'  sSQL = sSQL & ", December" & vbCrLf
'  sSQL = sSQL & ", January, February, March, April" & vbCrLf
'  sSQL = sSQL & ", May, Total_Months" & vbCrLf
'
' 'SQL Append Footer Data
'  sSQL = sSQL & ", ForecastMonth_ID" & vbCrLf
'  sSQL = sSQL & ", SLT_Value, SLT_ColumnName"
'  sSQL = sSQL & ", FilePath, FileName"
'  sSQL = sSQL & ") " & vbCrLf
'
'
' 'SQL Select Header Data
'  sSQL = sSQL & " Select "
'  sSQL = sSQL & " CLng(IIf(IsError([F1]),1,[F1])) AS CostCenter, Clng(NZ([F2],0)) as WorksheetForecast_Row" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F3],1)) as Account, CLng(NZ([F3],1)) as SAP_Account, Trim([F4]) as Acct_Description" & vbCrLf
'
' 'SQL Select body Data
'  sSQL = sSQL & ", CDbl(NZ([F5],0)*1000) as June, CDbl(NZ([F6],0)*1000) as July, CDbl(NZ([F7],0)*1000) as August" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F8],0)*1000) as September, CDbl(NZ([F9],0)*1000) as October" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F10],0)*1000) as November" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F11],0)*1000) as December" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F12],0)*1000) as January, CDbl(NZ([F13],0)*1000) as February, CDbl(NZ([F14],0)*1000) as March, CDbl(NZ([F15],0)*1000) as April" & vbCrLf
'  sSQL = sSQL & ", CDbl(NZ([F16],0)*1000) as May, CDbl(NZ([F17],0)*1000) as Total_Months" & vbCrLf
'
'  zzObsolete_SQL_Forecast = sSQL


End Function
