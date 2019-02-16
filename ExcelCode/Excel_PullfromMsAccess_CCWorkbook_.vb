Option Explicit

Public PUBLIC_ConnectionDB As ADODB.Connection

Public Sub aaMain_Set_Workbook_Data_Sources(xlWorksheet_Control As Excel.Worksheet)

'***************************************************************************************************
'*
'*  NOTES
'*  Update by: Phil Seiersen 11/15/2018 Contact info: www.pseiersen.com
'*  Created by Jason McCracken (ETW) on 9/21/08 - jason_mccracken@hotmail.com
'*
 
 'Local Variables
  Dim intQuery_Count As Integer 'Total number of queries
  Dim strSelected_Worksheet_QuerySource As String
  Dim strDatabase_PathAndName As String
  Dim sConnectionString As String
  Dim intQuery As Integer 'The current query
  
 'Excel Objecects
  Dim xlWorkBook As Excel.Workbook
  
 'ADO Objects
  Dim ConnectionToDB As ADODB.Connection
 
 
 On Error GoTo ProcErr
 
 'Get Workbook object from worksheet object
  Set xlWorkBook = xlWorksheet_Control.Parent
  
 'Set Update Run date in Worksheet -- Update Run Data
  xlWorkBook.Worksheets("Update Run Data").Range("vbaPull_DBdata_StartRunTime").Value = Now()
  
  
 ' 'OLD CODE
 ' 'Refresh the query on the Control worksheet so current period and quarter are set before queries are run
 '  Sheets("Control").Select
 '  Range("L6").Select
 '  Selection.QueryTable.Refresh BackgroundQuery:=False

 'Instantiate COM objects
  Set PUBLIC_ConnectionDB = New ADODB.Connection
 
 '-----------------------------------------------------------------------------------------------------------------
 '              >>> Connection String  <<<
  strDatabase_PathAndName = xlWorksheet_Control.Range("vbaDatabase_PathAndName")
 
  sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & vbCrLf
  sConnectionString = sConnectionString & "Data Source=" & strDatabase_PathAndName & ";" & vbCrLf
  
  Debug.Print strDatabase_PathAndName
  
 'Connect Database to Recordset
  With PUBLIC_ConnectionDB
    .CursorLocation = adUseClient
    .Open sConnectionString
    .CommandTimeout = 0
  End With

 
 
 'Get the number of queries for this Workbook from the Control Worksheet
  intQuery_Count = xlWorksheet_Control.Range("vbaQueryCount").Value
 
 '-------------------------------------------------------------------------------------------------------------
 'Loop through the queries and return the values to the specificed worksheets
  For intQuery = 1 To intQuery_Count
      
      
     'Get name of query worksheet to run
      strSelected_Worksheet_QuerySource = "Query" & intQuery
       
     '-----------------------------------------------------------------------------------------------
     '  ***** CALL Function to Get Database Data *****
     '  NOTE: if function
     '
      If abGet_Database_Data( _
                              xlWorksheet_Control, _
                              strSelected_Worksheet_QuerySource) = False Then
  
       'If there is an error with running the function then exit this subroutine
        GoTo ProcExit
      
      End If
      
      
     'Hide/Unhide query worksheets based on value set on control tab
      If xlWorksheet_Control.Range("vbaHideShow_Query") = "No" Then
      
        xlWorkBook.Worksheets(strSelected_Worksheet_QuerySource).Visible = False
        'Sheets(strSelected_Worksheet_QuerySource).Visible = False
        
      Else
        
        xlWorkBook.Worksheets(strSelected_Worksheet_QuerySource).Visible = True
        'Sheets(strSelected_Worksheet_QuerySource).Visible = True
        
      End If
      
  Next intQuery

    
  'Hide/Unhide worksheets
  'Call Change_Worksheet_Visibility(Sheets("Control").Range("D11"))
  
  
  'Set Update Run date in Worksheet -- Update Run Data
   xlWorkBook.Worksheets("Update Run Data").Range("vbaPull_DBdata_FinishRunTime").Value = Now()
  
  'Worksheet is uploaded
   MsgBox "Download Complete from FPA database!", vbExclamation + vbOKOnly, "Download Complete"


ProcExit:

 'Close connection objects
  PUBLIC_ConnectionDB.Close
  Set PUBLIC_ConnectionDB = Nothing

 'Set the focus back to the Control worksheet - cell A1
  xlWorksheet_Control.Select
  xlWorksheet_Control.Range("J21").Select
    

  Exit Sub

ProcErr:

  Select Case Err.Number
  
  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 1004
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3704 'ADO Object is Closed
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbExclamation
    Resume Next
    
  Case 3705   'Object Open
    MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
    
  Case -2147467259
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbExclamation, "Error handled by Forecast Template"
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source
    Resume ProcExit
  
  
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit


End Sub

Public Function abGet_Database_Data( _
                                     xlWorksheet_Control As Excel.Worksheet, _
                                     strWorksheet_QueryParametersSource As String) As Boolean


 'Local variables
  Dim sSQL As String
  Dim lngCost_Center As Long

  Dim strDestination_Range As String
  Dim strDatabase_PathAndName As String
  Dim sConnectionString As String
  Dim strDestination_Worksheet As String
  Dim strQuery_Name As String
  Dim strSource_Name_Cell As String
  
 'Excel Objects
  Dim xlWorkBook As Excel.Workbook
  Dim xlWorksheet_QuerySource As Excel.Worksheet
  Dim xlWorksheet_Destination As Excel.Worksheet
  Dim xlQueryTable As Excel.QueryTable
  
 'Excel Range Objects (this object is used for selected range within a worksheet)
  Dim xlRange_Destination_Selection As Excel.Range
  Dim xlRange_StartCell As Excel.Range
  
 'ADO Objects
  Dim rsDataset_TableQuery As ADODB.Recordset
  Dim ConnectionToDB As ADODB.Connection
  
  
 'Set function to TRUE
  abGet_Database_Data = True
    
 On Error GoTo ProcErr
 'Instantiate COM objects
  Set rsDataset_TableQuery = New ADODB.Recordset
  '  Set ConnectionToDB = New ADODB.Connection
 
 'Get Workbook object from worksheet object
  Set xlWorkBook = xlWorksheet_Control.Parent
    
 'Set Query Parameter Source worksheet object
  Set xlWorksheet_QuerySource = xlWorkBook.Worksheets(strWorksheet_QueryParametersSource)
  
 'NOTE: Get the Destination Worksheet name from cell C5 in the Query Worksheet: = strWorksheet_QueryParametersSource
  strDestination_Worksheet = xlWorksheet_QuerySource.Range("C5").Value
  Set xlWorksheet_Destination = xlWorkBook.Worksheets(strDestination_Worksheet)
  
  
 'Set worksheet to Visible and then set the focus to that worksheet
  With xlWorksheet_Destination
    .Visible = True
    .Select
  End With

 'Select cell A1 in destination worksheet
  xlWorksheet_Destination.Range("A1").Select

 'Set the range to the entire destination worksheet
  Set xlRange_Destination_Selection = xlWorksheet_Destination.Cells

  
 'Clear the Cells values only NOT formatting!
  xlRange_Destination_Selection.ClearContents
  
 'Clear any existing query on the Destination worksheet - if none exists an exception is trapped below and the code resumes the next step
  xlRange_Destination_Selection.QueryTable.Delete
    
    
 '----------------------------------------------------------------------------
 '                  *** Get Parameters ***

 ' >> Not Used <<
 'lngCost_Center = xlWorksheet_Control.Range("vbaCostCenter").Value 'Get the Cost Center for this Workbook from the Control Worksheet
 
 'SQL String
  sSQL = xlWorksheet_QuerySource.Range("C2")
 'Database filepath and name
  strDatabase_PathAndName = xlWorksheet_QuerySource.Range("C3")
 'Template write data to worksheet cell
  strDestination_Range = xlWorksheet_QuerySource.Range("C4")
 'Get QueryTable's query name
  strQuery_Name = xlWorksheet_QuerySource.Range("C6")
 'Optional parameter
  strSource_Name_Cell = xlWorksheet_QuerySource.Range("C7")
  
  'Debug.Print "Path database " & strDatabase_PathAndName
  

 
' '-----------------------------------------------------------------------------------------------------------------
' '              >>> Connection String  <<<
'  sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & vbCrLf
'  sConnectionString = sConnectionString & "Data Source=" & strDatabase_PathAndName & ";" & vbCrLf
'

'
' 'Connect Database to Recordset
'  With ConnectionToDB
'    .CursorLocation = adUseClient
'    .Open sConnectionString
'    .CommandTimeout = 0
'  End With

'  Debug.Print PUBLIC_ConnectionDB.ConnectionString
  

 'Open Recordset
  rsDataset_TableQuery.Open sSQL, PUBLIC_ConnectionDB
  
 'Set start cell range
  Set xlRange_StartCell = xlWorksheet_Destination.Range(strDestination_Range)
  
 'Create Table query
 'NOTE: xlRange_StartCell is where
  Set xlQueryTable = xlWorksheet_Destination.QueryTables.Add(rsDataset_TableQuery, xlRange_StartCell)
  
 'Create query table
  With xlQueryTable
    .Name = strQuery_Name
    .Refresh
  End With
  
 'Debug.Print sSQL
 
  
' 'Execute the query, return the results, and save the query to the worksheet
' '----------------------------------------------------------------------------
 'sConnectionString = "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDatabase_PathAndName 'Build the connection string

'  With xlWorksheet_Destination.QueryTables.Add(Connection:=Array(sConnectionString), Destination:=Range(strDestination_Range))
'      .CommandType = xlCmdSql 'Must use xlCmdSql to set the SQL statement below - using the table option by passes the SQL statement
'      .CommandText = sSQL 'The SQL statement that is built above is set as the source query here
'      .Name = strQuery_Name 'The Excel Query name
'      .FieldNames = True
'      .RowNumbers = False
'      .FillAdjacentFormulas = False
'      .PreserveFormatting = True
'      .RefreshOnFileOpen = False
'      .BackgroundQuery = True
'      .RefreshStyle = xlInsertDeleteCells
'      .SavePassword = False
'      .SaveData = True
'      .AdjustColumnWidth = True
'      .RefreshPeriod = 0
'      .PreserveColumnInfo = True
'      .SourceDataFile = strDatabase_PathAndName 'Database File
'      .Refresh BackgroundQuery:=False
'  End With
    
    
 'Clean up and formatting
 '----------------------------------------------------------------------------
  If strSource_Name_Cell <> "" Then 'This Section Writes the source worksheet of the query if the parameter is set

      With xlWorksheet_Destination.Range(strSource_Name_Cell)
          .Value = "Source: " & strWorksheet_QueryParametersSource 'Set the text value
          With .Font
              .Bold = True 'Make the text bold
              .Italic = False
              .ColorIndex = 11
          End With
      End With
      
  End If

 'Select Cell first cell in the destination worksheet
  xlWorksheet_Destination.Range("A1").Select
  
 'Sets the worksheet window zoom to 80%
  ActiveWindow.Zoom = 80
  
 'Hide the worksheet if control parameter is 'No'
  If Sheets("Control").Range("D9") = "No" Then
  
    Sheets(strDestination_Worksheet).Visible = False
    
  End If
        
ProcExit:

' 'Close connection objects
'  ConnectionToDB.Close
'  Set ConnectionToDB = Nothing
  
 'Close Recordset
  rsDataset_TableQuery.Close
  Set rsDataset_TableQuery = Nothing
     
 Exit Function

ProcErr:

  Select Case Err.Number
  
  Case 91 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbExclamation
    Resume ProcExit
    
  Case 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbExclamation
    Resume Next
    
  Case 1004 'Query table does NOT exist proceed to create a query table
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbExclamation
    Resume Next
    
  Case 3704 'ADO Object is Closed
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbExclamation
    Resume Next
    
  Case 3705   'Object Open
    MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit

  Case -2147217842, -2147217843 'Database not found
    abGet_Database_Data = False 'Exit out of Main_Set_Workbook_DataSources when this error occurs

    MsgBox "Data did NOT update in table" & vbCrLf & vbCrLf & _
    "Database still open/locked DB in wrong folder" & vbCrLf & "Error #" & Err.Number, vbExclamation + vbOKOnly, "Error Database connection problem"
    Resume ProcExit
    
    
  Case -2147217865, -2147467259 'Missing SQL Statment, Description type mismatch in
    abGet_Database_Data = False 'Exit out of Main_Set_Workbook_DataSources when this error occurs

    MsgBox "Data did NOT update for query " & strQuery_Name & vbCrLf & vbCrLf & _
    "Incorrect SQL Statment - Can't find query or Table" & vbCrLf & "Error #" & Err.Number, vbInformation + vbOKOnly, "Error SQL query"
    Resume ProcExit
    

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit


End Function

Public Function Return_YTD_SQL(inCurrent_Month As String, inMonth1 As String, inMonth2 As String, inMonth3 As String, inMonth4 As String, inMonth5 As String, inMonth6 As String, inMonth7 As String, inMonth8 As String, inMonth9 As String, inMonth10 As String, inMonth11 As String, inMonth12 As String) As String
On Error GoTo Error_Return_YTD_SQL
'-------------------------------------------------------------------------------------------------------------------
' This function returns the year to date SQL string
' Created by Jason McCracken (ETW) on 8-13-09
'-------------------------------------------------------------------------------------------------------------------


    Select Case inCurrent_Month
        Case "Jun"
            Return_YTD_SQL = inMonth1
        Case "Jul"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2
        Case "Aug"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3
        Case "Sep"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4
        Case "Oct"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5
        Case "Nov"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6
        Case "Dec"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7
        Case "Jan"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7 & " + " & inMonth8
        Case "Feb"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7 & " + " & inMonth8 & " + " & inMonth9
        Case "Mar"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7 & " + " & inMonth8 & " + " & inMonth9 & " + " & inMonth10
        Case "Apr"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7 & " + " & inMonth8 & " + " & inMonth9 & " + " & inMonth10 & " + " & inMonth11
        Case "May"
            Return_YTD_SQL = inMonth1 & " + " & inMonth2 & " + " & inMonth3 & " + " & inMonth4 & " + " & inMonth5 & " + " & inMonth6 & " + " & inMonth7 & " + " & inMonth8 & " + " & inMonth9 & " + " & inMonth10 & " + " & inMonth11 & " + " & inMonth12
    End Select

Exit_Return_YTD_SQL:
Exit Function

Error_Return_YTD_SQL:
    MsgBox "(" & Err.Number & ") " & Err.Description, vbExclamation, "Error at Function Return_YTD_SQL"
    Resume Exit_Return_YTD_SQL
    Resume
End Function


Public Function Return_Current_Quarter(inCurrent_Quarter As String, inQ1 As String, inQ2 As String, inQ3 As String, inQ4 As String) As String
On Error GoTo Error_Return_Current_Quarter
'-------------------------------------------------------------------------------------------------------------------
' Created by Jason McCracken(ETW) on 8/25/09 - jason_mccracken@hotmail.com
'-------------------------------------------------------------------------------------------------------------------
  
    Select Case inCurrent_Quarter
        Case "Q1"
            Return_Current_Quarter = inQ1
        Case "Q2"
            Return_Current_Quarter = inQ2
        Case "Q3"
            Return_Current_Quarter = inQ3
        Case "Q4"
            Return_Current_Quarter = inQ4
    End Select

Exit_Return_Current_Quarter:
Exit Function

Error_Return_Current_Quarter:
    MsgBox "(" & Err.Number & ") " & Err.Description, vbExclamation, "Error at Function Return_Current_Quarter"
    Resume Exit_Return_Current_Quarter
    Resume
End Function


Public Sub Change_Worksheet_Visibility(inHide As Variant)
On Error GoTo Error_Change_Worksheet_Visibility
'-------------------------------------------------------------------------------------------------------------------
' This Sub checks to see if a user wishes to hide or unhide the worksheets listed in the named range.
' Created by Jason McCracken(ETW) on 12/9/09 - jason_mccracken@hotmail.com
'-------------------------------------------------------------------------------------------------------------------
    
    Dim arrWorksheets As Variant 'An Array
    Dim intRecord As Integer
    Dim strWorksheet As String
    Dim booHide As String
    
    'Don't run this section if inHide has an incorrect or missing value
    If inHide = "Yes" Or inHide = "No" Then
    
        If inHide = "Yes" Then
        
            'Put the values of the named range into the array
            arrWorksheets = WorksheetFunction.Transpose(Worksheets("Control").Range("Hide_Worksheets").Value)
        
            'Loop through the array
            For intRecord = 1 To 30
        
                strWorksheet = arrWorksheets(intRecord)
            
                If strWorksheet <> "" Then
                    Sheets(strWorksheet).Visible = True 'Hide/Unhide the worksheet
                End If
            
            Next intRecord
         
        ElseIf inHide = "No" Then
               
            'Put the values of the named range into the array
            arrWorksheets = WorksheetFunction.Transpose(Worksheets("Control").Range("Hide_Worksheets").Value)
        
            'Loop through the array
            For intRecord = 1 To 30
        
                strWorksheet = arrWorksheets(intRecord)
            
                If strWorksheet <> "" Then
                    Sheets(strWorksheet).Visible = xlSheetVeryHidden 'Hide/Unhide the worksheet
                End If
            
            Next intRecord
        End If
    End If


Exit_Change_Worksheet_Visibility:
Exit Sub

Error_Change_Worksheet_Visibility:
    MsgBox "(" & Err.Number & ") " & Err.Description, vbExclamation, "Error at Sub Change_Worksheet_Visibility"
    Resume Exit_Change_Worksheet_Visibility
    Resume
End Sub
Public Sub Change_Query_Data_Visibility(inHide As Variant)
On Error GoTo Error_Change_Query_Data_Visibility
'-------------------------------------------------------------------------------------------------------------------
' This Sub checks to see if a user wishes to hide or unhide the worksheets listed in the named range.
' Created by Jason McCracken(ETW) on 12/9/09 - jason_mccracken@hotmail.com
'-------------------------------------------------------------------------------------------------------------------
    
    Dim arrWorksheets As Variant 'An Array
    Dim intRecord As Integer
    Dim strWorksheet As String
    Dim booHide As Boolean
    
    'Don't run this section if inHide has an incorrect or missing value
    If inHide = "Yes" Or inHide = "No" Then
    
        If inHide = "Yes" Then
        
            'Put the values of the named range into the array
            arrWorksheets = WorksheetFunction.Transpose(Worksheets("Control").Range("Hide_Query_Data").Value)
        
            'Loop through the array
            For intRecord = 1 To 30
        
                strWorksheet = arrWorksheets(intRecord)
            
                If strWorksheet <> "" Then
                    Sheets(strWorksheet).Visible = True 'Hide/Unhide the worksheet
                End If
            
            Next intRecord
         
        ElseIf inHide = "No" Then
               
            'Put the values of the named range into the array
            arrWorksheets = WorksheetFunction.Transpose(Worksheets("Control").Range("Hide_Query_Data").Value)
        
            'Loop through the array
            For intRecord = 1 To 30
        
                strWorksheet = arrWorksheets(intRecord)
            
                If strWorksheet <> "" Then
                    Sheets(strWorksheet).Visible = xlSheetVeryHidden 'Hide/Unhide the worksheet
                End If
            
            Next intRecord
        End If
    End If


Exit_Change_Query_Data_Visibility:
Exit Sub

Error_Change_Query_Data_Visibility:
    MsgBox "(" & Err.Number & ") " & Err.Description, vbExclamation, "Error at Sub Change_Query_Data_Visibility"
    Resume Exit_Change_Query_Data_Visibility
    Resume
End Sub

Public Sub Final_Format(Optional x As Boolean)
Application.ScreenUpdating = False

'Hide/Unhide worksheets
    Sheets("Control").Select
    Range("d9") = "No"
    Range("d11") = "No"
    Call Change_Worksheet_Visibility(Sheets("Control").Range("D11"))
    Call Change_Query_Data_Visibility(Sheets("Control").Range("D9"))
    
'Hide Columns
    Call Column_Hiding("True")
    
'Protect Sheets
    Call Protect_Unprotect("True")

Application.ScreenUpdating = True

End Sub
Public Sub Work_Format(Optional x As Boolean)
Application.ScreenUpdating = False

 'Hide/Unhide worksheets
    Sheets("Control").Visible = True
    Sheets("Control").Select
    Range("d9") = "Yes"
    Range("d11") = "Yes"
    Call Change_Worksheet_Visibility(Sheets("Control").Range("D11"))
    Call Change_Query_Data_Visibility(Sheets("Control").Range("D9"))
    
'Unprotect Sheets
    Call Protect_Unprotect("False")
    
'Unhide Columns
    Call Column_Hiding("False")
    
Sheets("Control").Select

Exit_Work_Format:
Exit Sub

Application.ScreenUpdating = True
End Sub


Public Sub Column_Hiding(intProtect As String)
Dim Oppo As String
    If intProtect = "True" Then
        Oppo = "False"
    ElseIf intProtect = "False" Then
        Oppo = "True"
    End If
    Sheets("FTE Roster").Range("A:A").EntireColumn.Hidden = True
    Sheets("FTE Roster").Range("B:C").EntireColumn.Hidden = intProtect
    Sheets("FTE Roster").Range("K:L").EntireColumn.Hidden = intProtect
    Sheets("FTE Roster").Range("R:AE").EntireColumn.Hidden = intProtect
    Sheets("FTE Roster").Range("CI:DI").EntireColumn.Hidden = intProtect
    Sheets("FTE Roster").Range("DJ:FI").EntireColumn.Hidden = intProtect
    Sheets("FTE Roster").Range("FJ:GC").EntireColumn.Hidden = intProtect
    Application.GoTo Sheets("FTE Roster").Range("D1"), True
        Range("F7").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("ETW Roster").Range("B:C").EntireColumn.Hidden = intProtect
    Sheets("ETW Roster").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("ETW Roster").Range("D1"), True
        Range("E10").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("Chargebacks").Range("B:D").EntireColumn.Hidden = intProtect
    Sheets("Chargebacks").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("Chargebacks").Range("E1"), True
        Range("J6").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("Cap Labor").Range("B:C").EntireColumn.Hidden = intProtect
    Sheets("Cap Labor").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("Cap Labor").Range("D1"), True
        Range("H9").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("SW-HW_OTHER EXP").Range("B:D").EntireColumn.Hidden = intProtect
    Sheets("SW-HW_OTHER EXP").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("SW-HW_OTHER EXP").Range("E1"), True
        Range("G7").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("CAPITAL & DEPRECIATION").Range("B:D").EntireColumn.Hidden = intProtect
    Sheets("CAPITAL & DEPRECIATION").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("CAPITAL & DEPRECIATION").Range("E1"), True
        Range("I9").Select
        ActiveWindow.FreezePanes = intProtect
        
    Application.GoTo Sheets("Notes").Range("A1"), True
        Range("A2").Select
        ActiveWindow.FreezePanes = intProtect
        
    Sheets("Forecast").Range("B:D,U:W,AN:AN").EntireColumn.Hidden = intProtect
    Sheets("Forecast").Range("A:A").EntireColumn.Hidden = True
    Application.GoTo Sheets("Forecast").Range("E1"), True
        Range("F7").Select
        ActiveWindow.FreezePanes = intProtect
End Sub


Public Sub Protect_Unprotect(intProtect As String)
    Dim strPwd As String
    strPwd = "focus"

    Sheets("Notes").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("FTE Roster").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
        ActiveSheet.EnableSelection = xlUnlockedCells
    Sheets("ETW Roster").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("Chargebacks").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("Cap Labor").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("SW-HW_OTHER EXP").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("CAPITAL & DEPRECIATION").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    Sheets("Forecast").Select
        ActiveSheet.Protect Password:=strPwd, DrawingObjects:=intProtect, Contents:=intProtect, Scenarios:=intProtect
    End Sub


Public Sub zzObsolete_PasteValues(Optional inMessage As Boolean)
    Sheets("FTE Roster").Select
    Range("D7:K7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M7:O7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("ETW Roster").Select
    Range("D10:V10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Chargebacks").Select
    Range("E6:AC6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Cap Labor").Select
    Range("D9:U58").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("SW-HW_OTHER EXP").Select
    Range("E7:Z56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("CAPITAL & DEPRECIATION").Select
    Range("D9:H9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J9:M9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N5:Z5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
