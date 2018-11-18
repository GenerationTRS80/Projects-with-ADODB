Public objRS_ZaiNetEOD As ADODB.Recordset
Public objCONN_Oracle As ADODB.Connection
Option Compare Database
Public Function Run_ZaiNetEOD_Reports(dStartDate As Date, dEndDate As Date, sReportName As String, Optional dReportMonth As Date)
'Run_PnL_Variance_Dataset_Report
 'Variables
    Dim sSQL As String

   On Error GoTo ProcErr
           
           
 'Set the file path name for the spreadsheet you want to update
    Const sFILEPATH = "\\porfiler02\shared\PPM\Asset Management - Wind\PhilS_AssetManagementWind\Dev_SandBox\SandBox_Reports\Downloaded_Oracle_Data.xls"
  
  
 '*****************************************************************************
 '*
 '*  CREATE SQL sting from V_Deal_Valuation
 '*
 '*
 
        Select Case sReportName
        
        
            Case "Assets"
 
                sSQL = SQL_V_Deal_Valuation_Asset(dStartDate, dEndDate)

 
            '    sSQL = SQL_V_Deal_Valuation_Asset(CDate(DateSerial(Year(Now()), Month(Now()), 1)), CDate(DateSerial(Year(Now()), Month(Now()), 2)), dReportMonth)
                 
            Case "PnL_Variance"
            
                sSQL = SQL_V_Deal_Valuation_PL(dStartDate, dEndDate)
        
            Case "PnL_VarianceTradePriceDetail"
            
                sSQL = SQL_V_Deal_Valuation_PL_TradePrice(dStartDate, dEndDate)
             
            Case Else
    
                MsgBox "Selection Not found", vbCritical
  
  
        End Select
        
            '    Debug.Print sSQL


 '*****************************************************************************
 '*
 '*  CREATE RECORDSET / CONNECT to ZaiNetEOD
 '*

     'If the function is false set to false
        If Connect_ZaiNetEOD_Create_RecordSet(sSQL) = False Then

            GoTo ProcExit

        End If


 '****************************************************************************
 '*
 '*  COPY RECORDSET to SPREADSHEET AND OPEN THE SPREADSHEET
 '*
 '*

      'Note if false then exit function

    Select Case sReportName

    Case "PnL_Variance", "PnL_VarianceTradePriceDetail"

        If CopyRecordset_to_Spreadsheet_Format_PnL_Variance(sFILEPATH, objRS_ZaiNetEOD) = False Then

            GoTo ProcExit

        End If

    Case Else

        If CopyRecordset_to_Spreadsheet(sFILEPATH, objRS_ZaiNetEOD) = False Then

            GoTo ProcExit

        End If


    End Select
    
    
ProcExit:

  'Close recordset and the connection
    objRS_ZaiNetEOD.Close
    Set objRS_ZaiNetEOD = Nothing
    
    objCONN_Oracle.Close
    Set objCONN_Oracle = Nothing

           
    Exit Function

ProcErr:
  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit
    
    
  Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
    'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
          
  Case 3704 'Recordset is already closed
    Resume Next
    
  Case 3705   'Objecte Open
   MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
     
  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit
  
    'MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function
Private Function SQL_V_Deal_Valuation_Asset(dStart_Date As Date, dEnd_Date As Date, Optional dReportMonth As Date) As String
 Dim sSQL As String


  sSQL = "Select A.""Asset"", A.""Report Date T"", A.""Instrument"",  A.""Strategy"", Cast(Sum(A.""Qty (native)"") as decimal(10,2)) ""Sum of Qty(native)"" " & vbCrLf

  sSQL = sSQL & "FROM ZAINETEOD.V_DEAL_VALUATIONS A " & vbCrLf
  
  sSQL = sSQL & "Where Instr(A.""Strategy"",'-',4,1)=0 and A.""Instrument""='PHYSICAL' and " & vbCrLf
  
 'Where Start Date
 
  ' sSQL = sSQL & "A.""Report Date T"" = to_date('" & Format(dReportMonth, "mm-dd-yyyy") & "','mm/dd/yyyy') " & vbCrLf

   sSQL = sSQL & "A.""Report Period"" between to_date('" & Format(dStart_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') AND to_date('" & Format(dEnd_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') " & vbCrLf
    
    
  'GROUP BY
    sSQL = sSQL & "Group by A.""Asset"", A.""Report Date T"", A.""Instrument"", A.""Strategy"" " & vbCrLf
    
    
 'Order by
   sSQL = sSQL & "Order by A.""Asset"",  A.""Strategy"";"
   
   
   SQL_V_Deal_Valuation_Asset = sSQL


End Function
Private Function SQL_V_Deal_Valuation_PL_TradePrice(dStart_Date As Date, dEnd_Date As Date) As String
    Dim sSQL As String

    
'Note: with Zkey 3800 rows returned for Hiwinds and Shiloh

   sSQL = "SELECT B.* From ("
 
   sSQL = sSQL & "SELECT " & vbCrLf
   
 '  sSQL = sSQL & "A.""Counterparty"", " & vbCrLf
   
   sSQL = sSQL & "A.""Counterparty"", A.""Book Name"", A.""Book ID"", A.""POD"", " & vbCrLf
   
   sSQL = sSQL & "A.""Asset"", A.""Report Date T"",  " & vbCrLf
   
   sSQL = sSQL & "A.""Start Date"" ""Report Period"", A.""Deliv Schedule""" & vbCrLf
   
   sSQL = sSQL & ", A.""Commodity"", A.""Instrument"",  A.""Strategy""" & vbCrLf
   
  'Add Remove Zkey
   sSQL = sSQL & ", A.""ZKey"" " & vbCrLf
   
  'Add Remove Side
   sSQL = sSQL & ", A.""Side"" " & vbCrLf
   

'   sSQL = sSQL & ", A.""Undisc Value2 USD"" "
'
'   sSQL = sSQL & ", A.""Undisc Value USD"" "
'
   sSQL = sSQL & ", A.""Strike"" "
     
'   sSQL = sSQL & ", A.""Underlying Price (native)"" "
   
   
'Add Remove Side
   sSQL = sSQL & ", A.""Trade Price"" " & vbCrLf
   

 'Aggregated Columns
   sSQL = sSQL & ", Cast(Sum(A.""Qty (native)"") as decimal(10,2)) ""Sum of Qty(native)"", Cast(Sum(A.""Exposed Value"") as decimal(10,2)) ""Sum of Exposed Value"" " & vbCrLf

   sSQL = sSQL & "FROM ZAINETEOD.V_DEAL_VALUATIONS A " & vbCrLf
  
  
'WHERE STATEMENT---------------------
  
  'Asset
    sSQL = sSQL & "Where A.""Asset"" ='FLATROCK' OR A.""Asset"" ='FLATROCK2' OR A.""Asset"" ='PLSNTVAL' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='BHORN' OR A.""Asset"" ='KLONDIKE2' OR A.""Asset"" ='KLONDIKE3' OR A.""Asset"" ='KLONDIKE3A'  OR A.""Asset"" ='KLWIND' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='TWINBUTTES' OR A.""Asset"" ='MTVIEW3' OR A.""Asset"" ='MOWIND' OR A.""Asset"" ='COLORADO' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='FLCL' OR A.""Asset"" ='TRIMONT' OR A.""Asset"" ='ELKRVR' OR A.""Asset"" ='PHNXWIND' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='DILLON' OR A.""Asset"" ='MINNDAKOTA' OR A.""Asset"" ='TOI' OR A.""Asset"" ='CASSELMAN' " & vbCrLf

    sSQL = sSQL & "OR A.""Asset"" ='LOCUST' OR A.""Asset"" ='PROVIDENCE' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='WINNEBAGO' OR A.""Asset"" ='BARTONCHAP' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='STWIND' OR A.""Asset"" ='SHILOH' OR A.""Asset"" ='HIWIND' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='WINNEBAGO' " & vbCrLf
        
'    sSQL = sSQL & "OR A.""Asset"" ='BARTONCHAP' " & vbCrLf
    
    
  'GROUP BY
  
    sSQL = sSQL & "Group by "
    
  '  sSQL = sSQL & "A.""Counterparty"", "
  
    sSQL = sSQL & "A.""Counterparty"", A.""Book Name"", A.""Book ID"", A.""POD"", "
  
    sSQL = sSQL & "A.""Asset"", A.""Report Date T"",  A.""Start Date"", A.""Deliv Schedule"", "
    
    sSQL = sSQL & "A.""Commodity"", A.""Instrument"",  A.""Strategy"""
    
    sSQL = sSQL & ", A.""ZKey"" "
        
    sSQL = sSQL & ", A.""Side"" "
    
'    sSQL = sSQL & ", A.""Undisc Value2 USD"" "
'
'    sSQL = sSQL & ", A.""Undisc Value USD"" "
'
    sSQL = sSQL & ", A.""Strike"" "
    
   sSQL = sSQL & ", A.""Underlying Price (native)"" "
     
    sSQL = sSQL & ", A.""Trade Price"""

  
  'End of Subquery
    sSQL = sSQL & ") B " & vbCrLf
    
    
     sSQL = sSQL & "Where "
     
  '   sSQL = sSQL & "B.""Commodity"" <>'GAS' And "
    
   'Where Start Date
    sSQL = sSQL & "B.""Report Period"" between to_date('" & Format(dStart_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') AND to_date('" & Format(dEnd_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') "
    
    
   'Order by
    sSQL = sSQL & "Order by ""Asset"", ""Report Period"";"
  
  
    SQL_V_Deal_Valuation_PL_TradePrice = sSQL
  
  
End Function
Private Function SQL_V_Deal_Valuation_PL(dStart_Date As Date, dEnd_Date As Date) As String
    Dim sSQL As String

    
'Note: with Zkey 3800 rows returned for Hiwinds and Shiloh

   sSQL = "SELECT B.* From ("
 
   sSQL = sSQL & "SELECT " & vbCrLf
   
   sSQL = sSQL & "A.""Counterparty"", " & vbCrLf
   
 '  sSQL = sSQL & "A.""Counterparty"", A.""Book Name"", A.""Book ID"", A.""POD"", " & vbCrLf
   
   sSQL = sSQL & "A.""Asset"", A.""Report Date T"",  " & vbCrLf
   
   sSQL = sSQL & "A.""Start Date"" ""Report Period"", A.""Deliv Schedule""" & vbCrLf
   
   sSQL = sSQL & ", A.""Commodity"", A.""Instrument"",  A.""Strategy""" & vbCrLf
   
  'Add Remove Zkey
   sSQL = sSQL & ", A.""ZKey"" " & vbCrLf
   
  'Add Remove Side
   sSQL = sSQL & ", A.""Side"" " & vbCrLf

 'Aggregated Columns
   sSQL = sSQL & ", Max(A.""Trade Price"") ""Max Trade Price"", Cast(Sum(A.""Qty (native)"") as decimal(10,2)) ""Sum of Qty(native)"", Cast(Sum(A.""Exposed Value"") as decimal(10,2)) ""Sum of Exposed Value"" " & vbCrLf

   sSQL = sSQL & "FROM ZAINETEOD.V_DEAL_VALUATIONS A " & vbCrLf
  
  
'WHERE STATEMENT---------------------
  
  'Asset
    sSQL = sSQL & "Where A.""Asset"" ='FLATROCK' OR A.""Asset"" ='FLATROCK2' OR A.""Asset"" ='PLSNTVAL' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='BHORN' OR A.""Asset"" ='KLONDIKE2' OR A.""Asset"" ='KLONDIKE3' OR A.""Asset"" ='KLONDIKE3A'  OR A.""Asset"" ='KLWIND' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='TWINBUTTES' OR A.""Asset"" ='MTVIEW3' OR A.""Asset"" ='MOWIND' OR A.""Asset"" ='COLORADO' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='FLCL' OR A.""Asset"" ='TRIMONT' OR A.""Asset"" ='ELKRVR' OR A.""Asset"" ='PHNXWIND' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='DILLON' OR A.""Asset"" ='MINNDAKOTA' OR A.""Asset"" ='TOI' OR A.""Asset"" ='CASSELMAN' " & vbCrLf

    sSQL = sSQL & "OR A.""Asset"" ='LOCUST' OR A.""Asset"" ='PROVIDENCE' " & vbCrLf
    
    sSQL = sSQL & "OR A.""Asset"" ='WINNEBAGO' OR A.""Asset"" ='BARTONCHAP' " & vbCrLf

    sSQL = sSQL & "OR A.""Asset"" ='STWIND' OR A.""Asset"" ='SHILOH' OR A.""Asset"" ='HIWIND' " & vbCrLf
    
    
  'GROUP BY
  
    sSQL = sSQL & "Group by "
    
    sSQL = sSQL & "A.""Counterparty"", "
  
  '  sSQL = sSQL & "A.""Counterparty"", A.""Book Name"", A.""Book ID"", A.""POD"", "
  
    sSQL = sSQL & "A.""Asset"", A.""Report Date T"",  A.""Start Date"", A.""Deliv Schedule"", "
    
    sSQL = sSQL & "A.""Commodity"", A.""Instrument"",  A.""Strategy"""
    
    sSQL = sSQL & ", A.""ZKey"" "
    
    sSQL = sSQL & ", A.""Side"""

  
  'End of Subquery
    sSQL = sSQL & ") B " & vbCrLf
    
    
     sSQL = sSQL & "Where "
     
  '   sSQL = sSQL & "B.""Commodity"" <>'GAS' And "
    
   'Where Start Date
    sSQL = sSQL & "B.""Report Period"" between to_date('" & Format(dStart_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') AND to_date('" & Format(dEnd_Date, "mm-dd-yyyy") & "','mm/dd/yyyy') "
    
    
   'Order by
    sSQL = sSQL & "Order by ""Asset"", ""Report Period"";"
  
  
    SQL_V_Deal_Valuation_PL = sSQL
  
  
End Function
Private Function Connect_ZaiNetEOD_Create_RecordSet(sSQL As String) As Boolean
    
 'Variable
  Dim Time_Out As Long
  
'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim Conn As ADODB.Connection

   
  On Error GoTo ProcErr
  
  'If the function is false set to false
     Connect_ZaiNetEOD_Create_RecordSet = True
    
'-----------Instantiate objects------------
  Set objCONN_Oracle = New ADODB.Connection
  Set objRS_ZaiNetEOD = New ADODB.Recordset
  


 'OLEDB connection string to Oracle ZaiNetEOD
  objCONN_Oracle.Open sConn_Oracle_ZaiNetEOD
  
'Set the time out to 6 minutes
  Time_Out = 360

  objCONN_Oracle.CommandTimeout = Time_Out
  
 'Open Recordset
  objRS_ZaiNetEOD.Open sSQL, objCONN_Oracle, adOpenForwardOnly, adLockReadOnly


ProcExit:

    Exit Function

ProcErr:
  'If the function is false set to false
     Connect_ZaiNetEOD_Create_RecordSet = False
     
  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit

       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    Resume Next
    
  Case 3705   'Object Open
   MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
    
  Case 3706
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
     
  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit
  
    'MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
  
  
End Function
Private Function CopyRecordset_to_Spreadsheet(sTargetSpreadsheet As String, rsSpeadsheet As ADODB.Recordset, Optional bSave As Boolean = True) As Boolean

'****************************************************************************
'*
'* NOTE: DO NOT USE MS EXCEL's  "Selection" for a substitute for a range of cells
'*          Excel will explicitly instantiate the "Selection" if you do use it
'*            You can 't close the instantiate selection
'*


'------- Variables ----------
 Dim lFields As Long
 Dim lRows As Long
 Dim iCol As Integer
 
 
'------- Excel Objects -------
 Dim appXl As Excel.Application
 Dim WB As Excel.Workbook
 Dim Wks As Excel.Worksheet
 
 Dim oCell_1 As Excel.Range
 
 Dim oCell_2 As Excel.Range
 
 Dim oRange As Excel.Range
 
 Const lROWS_CLEAR = 40000
 
 Const lCOLUMNS_CLEAR = 25
 
'Set CopyRecordset to Spreadsheet to TRUE
 CopyRecordset_to_Spreadsheet = True
    
    
  On Error GoTo ProcErr

'Instantiates Excel
  Set appXl = New Excel.Application  'CreateObject("Excel.Application")

  
   With appXl
  
        .AskToUpdateLinks = False
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
    
   End With

 'Update Links False, Read Only True ASLO if there is a password an errow will be raised
   Set WB = appXl.Workbooks.Open("" & sTargetSpreadsheet & "", True, False)
   
   Set Wks = WB.Worksheets("Sheet1")
   
 'Select worksheet Sheet1
   Wks.Activate

    
 '*** CLEAR ALL THE DATA IN THE WORKSHEET (EXCEPT THE HEADER) ***
    Set oCell_1 = Wks.Cells(1, 1)
    Set oCell_2 = Wks.Cells(lROWS_CLEAR, lCOLUMNS_CLEAR + 1)
    
    Set oRange = Wks.Range(oCell_1, oCell_2)
    
    oRange.Clear
    
    Set oCell_1 = Wks.Range("A1")
    
    oCell_1.Select
    
  'Count number of fields
    lFields = rsSpeadsheet.Fields.Count
    
  'Copy field names to the first row of the worksheet
    For iCol = 1 To lFields
        Wks.Cells(1, iCol).Value = rsSpeadsheet.Fields(iCol - 1).Name
    Next

  
   'Copy Recordset to spreadsheet
    Wks.Cells(2, 1).CopyFromRecordset rsSpeadsheet

   'Zoom 75 %
    appXl.ActiveWindow.Zoom = 75

   'Auto-fit teh column widths and row heights
    appXl.Selection.CurrentRegion.Columns.AutoFit
    appXl.Selection.CurrentRegion.Rows.AutoFit

   'Select the last column and format it
    Set oCell_1 = Wks.Cells(2, lFields)
    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
    
    oRange.NumberFormat = "#,##0_);[Red](#,##0)"

  'Show and save
   If bSave Then
   
     WB.Save

   End If

   'Show the spreadsheet
    appXl.Visible = True
    
    Set oCell_1 = Wks.Range("A1")
    
    oCell_1.Select
   

ProcExit:

    With appXl
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .AskToUpdateLinks = True
    End With
    
'Close recordset
    rsSpeadsheet.Close
    Set rsSpeadsheet = Nothing

    
    'Debug.Print "Spreadsheet is updated"

Exit Function

ProcErr:

 'If error then set CopyRecordset_to_Spreadsheet = False
  CopyRecordset_to_Spreadsheet = False

  Select Case Err.Number
  Case 91
    Resume Next
    
  Case 424 'Hourglass Comand
    MsgBox "None Excel object in use"
    Stop
    Resume Next
    
  Case 462
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical
    Stop
    Resume Next
  
  Case 1004
    MsgBox " Too Many instances of Excel open. Close one or more instances", vbCritical
    Resume ProcExit
    
  Case 3704 'Recordset is already closed
    Resume Next

    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
Resume ProcExit

End Function
Private Function CopyRecordset_to_Spreadsheet_Format_PnL_Variance(sTargetSpreadsheet As String, rsSpeadsheet As ADODB.Recordset, Optional bSave As Boolean = True) As Boolean

'****************************************************************************
'*
'* NOTE: DO NOT USE MS EXCEL's  "Selection" for a substitute for a range of cells
'*          Excel will explicitly instantiate the "Selection" if you do use it
'*            You can 't close the instantiate selection
'*


'------- Variables ----------
 Dim lFields As Long
 Dim lRows As Long
 Dim iCol As Integer
 
 
'------- Excel Objects -------
 Dim appXl As Excel.Application
 Dim WB As Excel.Workbook
 Dim Wks As Excel.Worksheet
 
 Dim oCell_1 As Excel.Range
 
 Dim oCell_2 As Excel.Range
 
 Dim oRange As Excel.Range
 
 Const lROWS_CLEAR = 40000
 
 Const lCOLUMNS_CLEAR = 25
 
'Set CopyRecordset to Spreadsheet to TRUE
 CopyRecordset_to_Spreadsheet_Format_PnL_Variance = True
    
  On Error GoTo ProcErr

'Instantiates Excel
  Set appXl = New Excel.Application  'CreateObject("Excel.Application")

  
   With appXl
  
        .AskToUpdateLinks = False
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
    
   End With

 'Update Links False, Read Only True ASLO if there is a password an errow will be raised
   Set WB = appXl.Workbooks.Open("" & sTargetSpreadsheet & "", True, False)
   
   Set Wks = WB.Worksheets("Sheet1")
   
 'Select worksheet Sheet1
   Wks.Activate

    
 '*** CLEAR ALL THE DATA IN THE WORKSHEET (EXCEPT THE HEADER) ***
    Set oCell_1 = Wks.Cells(1, 1)
    Set oCell_2 = Wks.Cells(lROWS_CLEAR, lCOLUMNS_CLEAR + 1)
    
    Set oRange = Wks.Range(oCell_1, oCell_2)
    
    oRange.Clear
    
'    Set oCell_1 = Wks.Cells(lROWS_CLEAR + 1, 1)
'    Set oCell_2 = Wks.Cells(lROWS_CLEAR + lROWS_CLEAR, lCOLUMNS_CLEAR + 1)
'
'    Set oRange = Wks.Range(oCell_1, oCell_2)
'
'    oRange.Clear
    
    Set oCell_1 = Wks.Range("A1")
    
    oCell_1.Select
    
  'Count number of fields
    lFields = rsSpeadsheet.Fields.Count
    
  'Copy field names to the first row of the worksheet
    For iCol = 1 To lFields
        Wks.Cells(1, iCol).Value = rsSpeadsheet.Fields(iCol - 1).Name
    Next
    
   'Copy Recordset to spreadsheet
    Wks.Cells(2, 1).CopyFromRecordset rsSpeadsheet

   'Zoom 75 %
    appXl.ActiveWindow.Zoom = 75

   'Auto-fit teh column widths and row heights
    appXl.Selection.CurrentRegion.Columns.AutoFit
    appXl.Selection.CurrentRegion.Rows.AutoFit
    
    
'Select Max Trad Price and format the column to currency
    Set oCell_1 = Wks.Cells(2, lFields - 2)
    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
    
        oRange.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

    Set oCell_1 = Wks.Cells(2, lFields - 1)
    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
    
        oRange.NumberFormat = "#,##0_);[Red](#,##0.00)"

    Set oCell_1 = Wks.Cells(2, lFields)
    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
    
        oRange.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

    
'Add Data Source column
 
  'Add header
    Set oCell_1 = Wks.Cells(1, lFields + 1)
    
        oCell_1.FormulaR1C1 = "Data Source"
    

  'Insert ZAINETEOD V_DEAL_VALUATIONS in the last column on the right
    Set oCell_1 = Wks.Cells(2, lFields)
    
    lHeaderRowNumber = oCell_1.Row
    
    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
    
    'Debug.Print "oRange " & oRange.Address
    
    Set oCell_1 = Wks.Cells(lHeaderRowNumber, lFields + 1)
    
    'Debug.Print "oCell_1  " & oCell_1.Address
    
    Set oCell_2 = Wks.Cells(oRange.Rows.Count + lHeaderRowNumber - 1, lFields + 1)
    
    'Debug.Print "oCell_2  " & oCell_2.Address
    
    Set oRange = Wks.Range(oCell_1, oCell_2)
    
    'Debug.Print "oRange FormulaR1C1" & oRange.Address
    

    oRange.FormulaR1C1 = "ZAINETEOD V_DEAL_VALUATIONS"
    

  'Show and save
   If bSave Then
   
     WB.Save

   End If
   
   
   'Show the spreadsheet
    appXl.Visible = True
    
    Set oCell_1 = Wks.Range("A1")
    
    oCell_1.Select
   

ProcExit:

    With appXl
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .AskToUpdateLinks = True
    End With
    
'Close recordset
    rsSpeadsheet.Close
    Set rsSpeadsheet = Nothing
    

    
    Debug.Print "Spreadsheet is updated"

Exit Function

ProcErr:

 'If error then set CopyRecordset_to_Spreadsheet = False
  CopyRecordset_to_Spreadsheet_Format_PnL_Variance = False

  Select Case Err.Number
  Case 91
    Resume Next
    
  Case 424 'Hourglass Comand
    MsgBox "None Excel object in use"
    Stop
    Resume Next
    
  Case 462
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical
    Stop
    Resume Next
  
  Case 1004
    MsgBox " Too Many instances of Excel open. Close one or more instances", vbCritical
    Resume ProcExit
    
  Case 3704 'Recordset is already closed
    Resume Next

    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  End Select
  
Resume ProcExit

End Function
