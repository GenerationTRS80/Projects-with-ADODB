Option Compare Database

Public Function Run_Import_ACES()
 Dim dLastReportDate As Date
 Dim dInputBox_Date As Date
 Dim lBeforeImport_RowNumber As Long
 Dim lAfterImport_RowNumber As Long

 Dim sFilePathName As String
 Dim sSQL As String
 Dim sWorksheet_Name As String

'ADO Recordset
  Dim RS As ADODB.Recordset
  Dim Conn As ADODB.Connection

  Const INSERT_TABLE = "tblGeneration"

  Const Data_Source = "ACES"

  Const FILE_PATH = "\\porfiler02\data1Sh\TRANSFER\Caminus ACES\Reports\User Reports\"

  On Error GoTo ProcErr

'------Instantiate objects----------
  Set Conn = New ADODB.Connection
  Set RS = New ADODB.Recordset

'OLEDB connection string to Access's Jet DB
  Conn.Open CurrentProject.Connection

'Use for Input Box default value
  sSQL = "SELECT Max(A.Delivery_Date) AS Delivery_Date, Max(A.Generation_Key) as RowNumber"
  sSQL = sSQL & " FROM " & INSERT_TABLE & "  as A"
  sSQL = sSQL & " WHERE (((A.Data_Source)=""" & Data_Source & """))"

 'Restrict by asset
 ' sSQL = sSQL & " WHERE (((A.Data_Source)=""" & Data_Source & """)) AND (((A.Asset)=""KLONDIKE3""))"

'  Debug.Print sSQL

'Open Recordset
  RS.Open sSQL, Conn, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

'Find the last row number in the table BEFORE being inserted into for the given date
  lBeforeImport_RowNumber = CDate(Nz(RS.Fields(1).Value, "1/1/1999"))

'Get Delivery Date from tblGeneration
  dLastReportDate = CDate(Nz(RS.Fields(0).Value, "1/1/1999"))

'Close recordset and the connection
  RS.Close
  Conn.Close

'Enter the date of the files you want to download
'*Note dLastReportDate is the most recent update of ACES data in tblGeneration
  dInputBox_Date = CDate(DateValue(InputBox("The date in the dialog box is the day after the latest ACES delivery date in the database ", "Import ACES file", DateAdd("d", 1, dLastReportDate))))

'Check to see if the input date is the same as today's date if so warn user and exit the subroutine
'Note:  You can NOT use today's ACES to download. You can only download up to the previous day and prior
    If dInputBox_Date = Date Then

        MsgBox "You can NOT download the ACES file with a delivery date that is the same as today's date" & vbCrLf & vbCrLf & _
                "You need to download the delivery date for " & Date & " on a date of tomorrow or later ", vbExclamation

        GoTo ProcExit

    End If

'Create the path and file name to the Text file download from ACES
  sWorksheet_Name = Format(dInputBox_Date, "YYYYMMDD") & "_AM_FORECAST"

'Get the file path for the worksheet
  sFilePathName = FILE_PATH & sWorksheet_Name & ".xls"

    '************** Import the ACES File **************
    'The code to import the ACES file is nested in the If then statement below
      If Not ImportAppend_ACES(sFilePathName, sWorksheet_Name, FILE_PATH, INSERT_TABLE) Then

        'If import failed exit subroutine
         GoTo ProcExit

      End If

'Reconnect to database
  Conn.Open CurrentProject.Connection

'Get the most recently downloaded Delivery_Date
'SQL
  sSQL = "SELECT tblGeneration.Delivery_Date AS Delivery_Date, tblGeneration.Generation_Key AS RowNumber"
  sSQL = sSQL & " FROM " & INSERT_TABLE
  sSQL = sSQL & " WHERE (((tblGeneration.Generation_Key)=(Select Max(Generation_Key) from tblGeneration)))"

'Reopen Recordset
  RS.Open sSQL, Conn, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

'Find the last row number in the table AFTER being inserted into for the given date
 lAfterImport_RowNumber = CDate(Nz(RS.Fields(1).Value, "1/1/1999"))

'Compare the before import row number to the the after import row number, if they are the same then 0 records were imported
 If lBeforeImport_RowNumber = lAfterImport_RowNumber Then

    MsgBox " 0 records were imported. The ACES file you just tried to download for " & dInputBox_Date & " doesn't contain records." _
            & vbCrLf & vbCrLf & "Download ACES file for  " & dInputBox_Date & "  again", vbExclamation

 Else

    'Get the date for the most recent delivery date
     dLastReportDate = CDate(Nz(RS.Fields(0).Value, "1/1/1999"))

     MsgBox "ACES Imported in the database for delivery date " & dLastReportDate, vbInformation

 End If


ProcExit:

'Close Recordset
    RS.Close
    Set RS = Nothing

    Exit Function

ProcErr:
  Select Case Err.Number
  Case 13 'Cancel button hit on input box
    Resume ProcExit

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3061 'File has not been downloaded
    MsgBox "The Date " & dInputBx_Date & " has NOT been downloaded " & vbCrLf & vbCrLf & "Try down loading the file again with your date criteria", vbExclamation
    Resume ProcExit

   Case 3070 'File doesnt have the correct columns
    MsgBox "The format of the PowerDex Index file is not correct. Download the PowerDex file from ZEMarketAnalyszer again.", vbExclamation
    Resume ProcExit

  Case 3704 'Recordset is already closed
    Resume Next

  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit

  Case -2147467259 'Invalid Path
    MsgBox "You need to close the database and reopen it. It is locked!", vbExclamation
    Resume ProcExit

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit


End Function
Private Function ImportAppend_ACES(File_Path_Name As String, sWorksheet_Name As String, Optional sFile_Path As String, Optional sInsert_TableName As String) As Boolean
'NOTE sFile_Path is for csv (text file) import
'NOTE you can change the table you insert to but the sturcture of the table must be the same
'NOTE ImportAppend_ACES if it returns false means the import failed. It couldnt find the file or file path

 Dim sSQL As String
 Dim sSQL_False As String
 Dim sWorksheet As String
 Dim sImport_Spreadsheet As String
 Dim sExcelPathFile As String
 Dim QD As QueryDef

'Set default value of ImportAppend ACES to TRUE
 ImportAppend_ACES = True


''Set the querydef for ACES
'    Set qd = CurrentDb.QueryDefs("q_App_ACES")

'Connect to Spreadsheet

'****TEXT file import***
  sImport_Spreadsheet = "[Text;DSN=ACES_CSV;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=437;DATABASE=" & sFile_Path & "].[" & sWorksheet_Name & ".csv]"

'SQL to insert 3TIER Asset Hour MW and date
'Note sInsert_TableName is the table inserted into

  sSQL = "INSERT INTO " & sInsert_TableName & " ( Asset, Delivery_Location, Forecast_Generation, Forecast_Generation_Long_Desc, Data_Source, Delivery_Date" & vbCrLf

  sSQL = sSQL & ", h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13, h14, h15, h16, h17, h18, h19, h20, h21, h22, h23, h24 )"
  sSQL = sSQL & "SELECT tblAsset.Asset, tblAsset.Asset_Delivery_Location, " & vbCrLf
  sSQL = sSQL & "IIf(Left(Trim([A.F5]),2)=""DA"",""DA"",IIf(Trim(Left([A.F5],2))=""HA"" Or Trim([A.F5])=""Estimate"",""HA"",""DA"")) AS Forecast_Generation, " & vbCrLf

 ' sSQL = sSQL & "IIf(Trim([A.F5])=""Estimate"",""HA Estimate"",IIf(Trim([A.F5])=""3 Tier Forecast Generation"",""DA 3 Tier Forecast Generation"",[F5]))AS Forecast_Generation_Long_Desc, " & vbCrLf

  sSQL = sSQL & "Trim(Left(IIf(Trim([A.F5])=""Estimate"",""HA Estimate"",IIf(Trim([A.F5])=""3 Tier Forecast Generation"",""DA 3 Tier"",[F5])),12)) AS Forecast_Generation_Long_Desc, " & vbCrLf
  sSQL = sSQL & """ACES"" AS Data_Source, " & vbCrLf
  sSQL = sSQL & "DateSerial(Left([A.F3],4),Mid([A.F3],5,2),Right([A.F3],2)) AS Delivery_Date, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F7],0)) AS h1, CLng(NZ(A.[F8],0)) AS h2, CLng(NZ(A.[F9],0)) AS h3, CLng(NZ(A.[F10],0)) AS h4, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F11],0)) AS h5, CLng(NZ(A.[F12],0)) AS h6, CLng(NZ(A.[F13],0)) AS h7, CLng(NZ(A.[F14],0)) AS h8, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F15],0)) AS h9, CLng(NZ(A.[F16],0)) AS h10, CLng(NZ(A.[F17],0)) AS h11, CLng(NZ(A.[F18],0)) AS h12, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F19],0)) AS h13, CLng(NZ(A.[F20],0)) AS h14, CLng(NZ(A.[F21],0)) AS h15, CLng(NZ(A.[F22],0)) AS h16, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F23],0)) AS h17, CLng(NZ(A.[F24],0)) AS h18, CLng(NZ(A.[F25],0)) AS h19, CLng(NZ(A.[F26],0)) AS h20, " & vbCrLf
  sSQL = sSQL & "CLng(NZ(A.[F27],0)) AS h21, CLng(NZ(A.[F28],0)) AS h22, CLng(NZ(A.[F29],0)) AS h23, CLng(NZ(A.[F30],0)) AS h24 " & vbCrLf

  sSQL = sSQL & "FROM " & sImport_Spreadsheet & " as A INNER JOIN tblAsset ON A.F4 = tblAsset.Asset_ACES_Name " & vbCrLf

' sSQL = sSQL & "WHERE (((tblAsset.Asset)=""KLONDIKE3""))"

  Debug.Print sSQL


'****************Import the file data*******************
'The code to import is nested in the if then statement   ADO_ImportExcel(sSQL)

'NOTE: if the file isnt found then have a warning msg
  If Not ADO_ImportExcel(sSQL) Then

    MsgBox "Could not find the file at the location " & File_Path_Name & vbCrLf & vbCrLf & " With the file name " & sWorksheet_Name, vbExclamation

    ImportAppend_ACES = False

  End If


End Function
Private Function ADO_ImportExcel(sSQL_Import) As Boolean
'NOTE NEED UPDATE BY IN PRICECURVE FILE

'Append Excel Sheet data to Access Table

'ADO objects
 Dim Conn As ADODB.Connection
 
'Variables
 Dim sSQL As String, sWorksheet As String
 
'Default ADO_ImportExcel to true
 ADO_ImportExcel = True

'Instantiate objects
 Set Conn = New ADODB.Connection

'Open a connection using the connection string
 Conn.Open StrConnectDB

'Check procedural erros with VBA
 On Error GoTo ProcErr
 
 DoCmd.Hourglass True
 
'Import Excel Spreadsheet via ADO connection object into a table
 Conn.Execute sSQL_Import

ProcExit:
  
  DoCmd.Hourglass False
  
'Close Connection object
  Conn.Close
  Set Conn = Nothing
  
Exit Function


ProcErr:

  DoCmd.Hourglass False
  Select Case Err.Number
  Case -2147217900 'Missing SQL Statement
   MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    'MsgBox " The Columns in the Access Gas Curves.xls" & vbCrLf & "don't match the columns in the tbl_PriceCurves table", vbCritical
    Resume ProcExit
    
  Case -2147467259 'Invalid Path
    MsgBox "Description " & Err.Description, vbExclamation
    Resume ProcExit
    
  Case -2147217865 'Wrong Excel Sheet Name
    
    '******** NOTE the file was not created ****************
    'Set default to false and exit function
    
     'Debug.Print "Description " & Err.Description
    
     ADO_ImportExcel = False
     Resume ProcExit
    
  Case 3625 'The array of field being input dont match the recordset field names
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Resume Next
  Case 3704 'Recordset empty End program to stop more errors
    Resume Next
  Case Else
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    Resume Next
  End Select
Resume ProcExit

End Function
