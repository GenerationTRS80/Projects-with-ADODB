Option Compare Database
Public pub_objRS_TagList As ADODB.Recordset
Public pub_objRS_Daily_TagList_Reliability As ADODB.Recordset


Private Sub Test_Connection()

   'Variables
    Dim sConn As String
    Dim sSQL As String
    Dim sSQL_TagList As String
    Dim i As Integer
    Dim dReport_Start As Date
    Dim dReport_End As Date
    Dim iBeginRow As Long
    Dim iEndRow As Long
    
    
   'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim objRS_Temp As ADODB.Recordset
    Dim objRS_TagList As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection
    Dim objFld As ADODB.Field


    On Error GoTo ProcErr
           
   '-----------Instantiate objects------------
    Set objConnOLEDB = New ADODB.Connection
    Set objRS = New ADODB.Recordset
    Set objRS_Temp = New ADODB.Recordset
    Set objRS_TagList = New ADODB.Recordset
  
  
   '--- Set Constants ----
    Const sFILEPATH = "\\porfiler02\shared\PPM\Asset Management - Wind\PhilS_AssetManagementWind\Dev_SandBox\SandBox_Reports\PI_Data_Report.xls"
  
    Const ROW_COUNT_TAG_LIST_PARSE_LIMIT = 901
  
  
    dReport_Start = Format(#10/1/2008#, "yyyy-mm-dd hh:mm:ss")

    dReport_End = Format(#10/31/2008#, "yyyy-mm-dd hh:mm:ss")
  

'   'OLEDB connection string to PI
'    objConnOLEDB.ConnectionTimeout = 360
'
'
'    objConnOLEDB.Open sPI_OLEDB


'   'Open Recordset
'    objRS.Open sSQL, objConnOLEDB, adOpenForwardOnly, adLockReadOnly


    DoCmd.Hourglass True


   '***Create and populate the recordset pub_objRS_TagList with tags and their associated assets
   'Note 4 fields in recordset Row_Number, Tag_List, Scan, Asset_Name
    PI_Create_pub_objRS_TagList
    
    
   '************  COPY RECORDSET TO SPREADSHEET  *********
    PI_CopyRecordset_to_Spreadsheet sFILEPATH, pub_objRS_TagList
    
    
'    PI_Create_pub_objRS_Daily_TagList_Reliability dReport_Start, dReport_End
'
'
'   '************  COPY RECORDSET TO SPREADSHEET  *********
'    PI_CopyRecordset_to_Spreadsheet sFILEPATH, pub_objRS_Daily_TagList_Reliability

  
    
'   '************  COPY RECORDSET TO SPREADSHEET  *********
'    PI_CopyRecordset_to_Spreadsheet sFILEPATH, objRS


    Debug.Print "Subroutine Completed "


ProcExit:

    pub_objRS_TagList.Close
    Set pub_objRS_TagList = Nothing
    
    pub_objRS_Daily_TagList_Reliability.Close
    Set pub_objRS_TagList = Nothing
    
    objRS_Temp.Close
    Set objRS_Temp = Nothing
    
    objRS.Close
    Set objRS = Nothing
    
    objConnOLEDB.Close
    Set objConnOLEDB = Nothing
    
    DoCmd.Hourglass False

    Exit Sub

ProcErr:
  Select Case Err.Number

  Case 91 'Object does not exist
    'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & err.Number & vbCrLf & "Description " & err.Description & vbCrLf & vbCrLf & " The source " & err.Source, vbCritical
    'Stop
    Resume Next
    
  Case 3705   'Object Open
   MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
     
  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit
    
  Case -2147217900 'SQL syntax NOT correct
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
  

End Sub
Private Function PI_sSQL_Daily_Reliability(RS_TagList As ADODB.Recordset, dReport_Start As Date, dReport_End As Date) As String
  
    Dim sSQL As String
    Dim i As Integer
    
    
    On Error GoTo ProcErr
 
    i = 0
 

    sSQL = "Select B.Tag_Asset Asset_Name, B.ReliETime , B.UnReliETime, B.ReliETime / (B.ReliETime + B.UnReliETime) Daily_Reliability From (" & vbCrLf

 'Aggregate by Asset
    sSQL = sSQL & "Select A.Tag_Asset, Sum(A.ReliElapsedTime) ReliETime, Sum(A.UnReliElapsedTime) UnReliETime From (" & vbCrLf

Debug.Print sSQL

' 'Aggregate by Tag Name
'    sSQL = "Select A.Row_Number, A.Tag_Asset,  A.TagName, Sum(A.ReliElapsedTime) ReliETime, Sum(A.UnReliElapsedTime) UnReliETime From ("


'Cycle through all the tags and total

''Set seek to index
'RS_TagList.Seek 952, adSeekAfter

   Do While Not RS_TagList.EOF

        
        If i <> 1 Then

            sSQL = sSQL & " Union All " & vbCrLf

        End If
        
    
 
       sSQL = sSQL & "SELECT " & RS_TagList.Fields(0).Value & " Row_Number, '" & RS_TagList.Fields(3).Value & "' Tag_Asset, '" & RS_TagList.Fields(1).Value & "' TagName, cast(ReliableT as float64) ReliElapsedTime, cast(UnreliableT as float64) UnReliElapsedTime " & _
                        "FROM (SELECT TIMEEQ('" & RS_TagList.Fields(1).Value & "', '" & dReport_Start & "', '" & dReport_End & "', 'Good') ReliableT, TIMEEQ('" & RS_TagList.Fields(1).Value & "', '" & dReport_Start & "', '" & dReport_End & "', 'Bad') UnreliableT) Time"


     
        RS_TagList.MoveNext
        
   Loop


        sSQL = sSQL & " ) A"
    
    
    'Group By
     '   sSQL = sSQL & " Group by A.Row_Number, A.Tag_Asset, A.TagName"
        
        
        sSQL = sSQL & " Group by A.Tag_Asset"


        sSQL = sSQL & " ) B"


    'Return SQL statement
        PI_sSQL_Daily_Reliability = sSQL
        
    
        Debug.Print "Items " & i
        

ProcExit:
   
    RS_TagList.Close
    Set RS_TagList = Nothing

    Exit Function

ProcErr:
  Select Case Err.Number
  
   Case 91 'Object does not exist
   MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
   Stop
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & err.Number & vbCrLf & "Description " & err.Description & vbCrLf & vbCrLf & " The source " & err.Source, vbCritical
    'Stop
    Resume Next
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
Private Function PI_sSQL_Aggregate_Daily_TagList_Reliability(RS_TagList As ADODB.Recordset, dReport_Start As Date, dReport_End As Date) As String
  
    Dim sSQL As String
    Dim iRowNumber_TagList As Integer
    
    
    On Error GoTo ProcErr
 
    iRowNumber_TagList = 0
 
 '*** Create SQL to aggregate Taglists on a daily basis
 'Aggregate by Tag Name
    sSQL = "Select A.Row_Number, A.Start_Date, A.End_Date, A.Tag_Asset,  A.TagName, Sum(A.ReliElapsedTime) ReliETime, Sum(A.UnReliElapsedTime) UnReliETime From (" & vbCrLf


 'Cycle through all the tags and total
   Do While Not RS_TagList.EOF

        iRowNumber_TagList = iRowNumber_TagList + 1

        
        If iRowNumber_TagList <> 1 Then

            sSQL = sSQL & " Union All " & vbCrLf

        End If
            
        sSQL = sSQL & "SELECT " & RS_TagList.Fields(0).Value & " Row_Number, '" & RS_TagList.Fields(3).Value & "' Tag_Asset, '" & RS_TagList.Fields(1).Value & "' TagName, cast(ReliableT as float64) ReliElapsedTime, cast(UnreliableT as float64) UnReliElapsedTime " & _
                        ", '" & dReport_Start & "' Start_Date, '" & dReport_End & "' End_Date " & _
                        "FROM (SELECT TIMEEQ('" & RS_TagList.Fields(1).Value & "', '" & dReport_Start & "', '" & dReport_End & "', 'Good') ReliableT, TIMEEQ('" & RS_TagList.Fields(1).Value & "', '" & dReport_Start & "', '" & dReport_End & "', 'Bad') UnreliableT) Time"

     
        RS_TagList.MoveNext
        
   Loop

        sSQL = sSQL & " ) A"

        
      'Group By
        sSQL = sSQL & " Group by A.Row_Number, A.Tag_Asset, A.TagName, A.Start_Date, A.End_Date "


    'Return SQL statement
        PI_sSQL_Aggregate_Daily_TagList_Reliability = sSQL
    
      '  Debug.Print "Items " & i
        
ProcExit:
   
    RS_TagList.Close
    Set RS_TagList = Nothing

    Exit Function

ProcErr:
  Select Case Err.Number
  
   Case 91 'Object does not exist
   MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
   Stop
   Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & err.Number & vbCrLf & "Description " & err.Description & vbCrLf & vbCrLf & " The source " & err.Source, vbCritical
    'Stop
    Resume Next
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
Private Function PI_CopyRecordset_to_Spreadsheet(sTargetSpreadsheet As String, rsSpeadsheet As ADODB.Recordset, Optional bSave As Boolean = True) As Boolean

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

''Select the last column and format it
'    Set oCell_1 = Wks.Cells(2, lFields)
'    Set oRange = Wks.Range(oCell_1, oCell_1.End(xlDown))
'
'        oRange.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

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

Private Sub Run_PI_Create_pub_objRS_Daily_TagList_Reliabili()

    On Error GoTo ProcErr
    

   PI_Create_pub_objRS_Daily_TagList_Reliability #10/2/2008#, #10/3/2008#
      

   
   Do While Not pub_objRS_Daily_TagList_Reliability.EOF

     Debug.Print pub_objRS_Daily_TagList_Reliability(0).Value & " " & pub_objRS_Daily_TagList_Reliability(1).Value & _
                " " & pub_objRS_Daily_TagList_Reliability(2).Value & " " & pub_objRS_Daily_TagList_Reliability(3).Value & _
                " " & pub_objRS_Daily_TagList_Reliability(4).Value & " " & pub_objRS_Daily_TagList_Reliability(5).Value & _
                " " & pub_objRS_Daily_TagList_Reliability(6).Value


     pub_objRS_Daily_TagList_Reliability.MoveNext

   Loop

   
ProcExit:
   

    Exit Sub

ProcErr:
  Select Case Err.Number
  
   Case 91, 424 'Object does not exist
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
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

 
End Sub
Private Sub PI_Create_pub_objRS_Daily_TagList_Reliability(dReport_Start As Date, Optional dReport_End As Date)

   'Variables
    Dim sSQL As String
    Dim dReport_Date As Date
    Dim lRowCount_DailyReliability As Long
    Dim lRowCount_TagList As Long
    Dim lTotalRows_TagList As Long
    Dim lParseLimit As Long
    Dim iBeginRow As Integer
    Dim iEndRow As Integer
    Dim intDateBegin As Integer
    Dim intDateAdd As Integer
    
   'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim objRS_Temp1 As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection
    
    
    Const ROW_COUNT_TAG_LIST_PARSE_LIMIT = 901
    
    
   '-----------Instantiate objects------------
    Set objConnOLEDB = New ADODB.Connection
    Set objRS = New ADODB.Recordset
    Set objRS_Temp1 = New ADODB.Recordset

    On Error GoTo ProcErr
    
    
   'Open a connection to the PI Database
   'OLEDB connection string to PI
    objConnOLEDB.ConnectionTimeout = 360
 
    objConnOLEDB.Open sPI_OLEDB
    
    
  'This overwrites the instantance of Pubrs_1 and creates disconnect Recordset
    Set pub_objRS_Daily_TagList_Reliability = New ADODB.Recordset
    
    
  'To disconnect recordset us cursor location client
    pub_objRS_Daily_TagList_Reliability.CursorLocation = adUseClient
  
  
  'Fabricate Recordset Append Fields to recordset
    pub_objRS_Daily_TagList_Reliability.Fields.Append "Row_Number", adBigInt
    pub_objRS_Daily_TagList_Reliability.Fields.Append "Start_Date", adDate
    pub_objRS_Daily_TagList_Reliability.Fields.Append "End_Date", adDate
    pub_objRS_Daily_TagList_Reliability.Fields.Append "Tag_Asset", adBSTR
    pub_objRS_Daily_TagList_Reliability.Fields.Append "TagName", adBSTR
    pub_objRS_Daily_TagList_Reliability.Fields.Append "ReliETime", adBigInt
    pub_objRS_Daily_TagList_Reliability.Fields.Append "UnReliETime", adBigInt


   'Open pub_objRS to add records from objRS
    pub_objRS_Daily_TagList_Reliability.Open

    
   'Total TagList Row Count
    lTotalRows_TagList = pub_objRS_TagList.RecordCount

   
   '**** Create Daily Reliability recordset ***
    'You can only have 900 or less tags for the SQL statement
    'The OLEDB parser will only accept a SQL string up to 900 tags
   
    lRowCount_DailyReliability = 0
    
    intDateAdd = 0


  'Run for each day between dReport_Start to dReport_End date
    For intDateBegin = Day(dReport_Start) To Day(dReport_End)
    
         Debug.Print "This is the date that is running " & DateAdd("d", dReport_Start, intDateAdd)
   
        'The ParseLimit is the number of Tags the OLEDB Parser can handle
        'from the SQL string generated by PI_sSQL_Daily_Reliability
         lParseLimit = 0
    
         lRowCount_TagList = 0
    
    
        'This loop queries 901 taglist with their reliability and unreliablity numbers.
        'Note Only 901 tags are appended at a time due to the OLEDB Parser can only handle a SQL string for that many tags and their reliability
        'the tags that are returned are then appended to the pub_objRS_Daily_TagList_Reliability recordset
    
         Do While lRowCount_TagList <= lTotalRows_TagList
    
    
            'Filter pub_objRS_TagList recordset by  row
             If lRowCount_TagList = lParseLimit Then
        
         
                'If recordset is open close it
                 If objRS.State = adStateOpen Then
            
                    objRS.Close

                 End If
   
                            
                 lParseLimit = lParseLimit + ROW_COUNT_TAG_LIST_PARSE_LIMIT
         
                 iBeginRow = lRowCount_TagList
         
                 iEndRow = lParseLimit
            
                 Debug.Print " Row Count " & lRowCount_TagList & " Parse Limit " & lParseLimit
            

                 'Create Temp recordset for filtering
                 'NOTE: Once filtering is done to a fabricated  recordset such as pub_objRS_TagList it is premanent
          
          
                 '**** pub_objRS_TagList is a PUBLIC RECORDSET created by PI_Create_pub_objRS_TagList ***
                  Set objRS_Temp1 = pub_objRS_TagList.Clone
            
            
                 'Set the row number you want the tag list to begin at
                  objRS_Temp1.Filter = "Row_Number>" & iBeginRow & " and Row_Number<= " & iEndRow & ""
            
            
                 '***Create SQL statement for Daily Relaibility
                  sSQL = PI_sSQL_Aggregate_Daily_TagList_Reliability(objRS_Temp1, DateAdd("d", dReport_Start, intDateAdd), DateAdd("d", dReport_Start, intDateAdd + 1))
            

                 'Open Recordset
                  objRS.Open sSQL, objConnOLEDB, adOpenKeyset, adLockBatchOptimistic
           

                'Add recoreds from objRS_TagList to PR_TagList_RS
                 objRS.MoveFirst
   
   
             End If
         
            'Create row numbers
             lRowCount_TagList = lRowCount_TagList + 1
         

            '*********** NEED TO CHECK LOOP ********************
             If Not objRS.EOF Then
        
                 With pub_objRS_Daily_TagList_Reliability

                     .AddNew fieldlist:=Array("Row_Number", "Start_Date", "End_Date", "Tag_Asset", "TagName", "ReliETime", "UnReliETime"), _
                        values:=Array(lRowCount_TagList, CDate(objRS("Start_Date").Value), CDate(objRS("End_Date").Value), objRS("Tag_Asset").Value, objRS("TagName").Value, objRS("ReliETime").Value, objRS("UnReliETime").Value)
                    '.UpdateBatch  This is only need, if you are resyncing a recordset
                  
                 End With

                 objRS.MoveNext
            
             Else
        
                 Debug.Print " objRS at EOF with row number " & lRowCount_TagList
            
    
             End If
        
         Loop

        'Add one date to start date
         intDateAdd = intDateAdd + 1
        
    Next intDateBegin
    
    
   'Move the record to the first row
    pub_objRS_Daily_TagList_Reliability.MoveFirst
        
    
ProcExit:

    objRS_Temp1.Close
    Set bjRS_Temp1 = Nothing
    
    objRS.Close
    Set objRS = Nothing
    
    objConnOLEDB.Close
    Set objConnOLEDB = Nothing

    Exit Sub

ProcErr:
  Select Case Err.Number

  Case 91 'Object does not exist
    'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    'Stop
    Resume Next
    
 Case 3021 'No Records returned
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    'Stop
    Resume Next
    
  Case 3705   'Object Open
   ' MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   ' Stop
    Resume Next
     
  Case 3709 'Connection object isnt open
   ' MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   ' Stop
    Resume Next
    
  Case -2147217900 'SQL syntax NOT correct
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
      

End Sub
Public Sub PI_Create_pub_objRS_TagList()

   'Variable
    Dim sSQL As String
    Dim sAssetName As String
    Dim lRecordCount As Long
    Dim lRowNumber As Long

   'DAO Objects
    Dim QD As DAO.QueryDef
    Dim objRS_DAO As DAO.Recordset
    
   'ADO Objects
    Dim objRS_TagList As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection


    On Error GoTo ProcErr
    
   '-----------Instantiate objects------------
    Set objConnOLEDB = New ADODB.Connection
    Set objRS_TagList = New ADODB.Recordset

  
   'OLEDB connection string to PI
    objConnOLEDB.ConnectionTimeout = 360
 
    objConnOLEDB.Open sPI_OLEDB
    
   
   'This overwrites the instantance of Pubrs_1 and creates disconnect Recordset
    Set pub_objRS_TagList = New ADODB.Recordset
    
  'To disconnect recordset us cursor location client
    pub_objRS_TagList.CursorLocation = adUseClient
  
  
   'Fabricate Recordset Append Fields to recordset
    pub_objRS_TagList.Fields.Append "Row_Number", adBigInt
    pub_objRS_TagList.Fields.Append "Tag_List", adBSTR
    pub_objRS_TagList.Fields.Append "Scan", adBSTR
    pub_objRS_TagList.Fields.Append "Asset_Name", adBSTR

    
   'Open pub_objRS_TagList to add records from objRS_TagList
    pub_objRS_TagList.Open
    

   'Get a list Park Tag Prefixs from local table tblAsset
    sSQL = "SELECT A.Asset_Park_Prefix,A.Asset_Park_Name, A.Asset "
    sSQL = sSQL & "FROM tblAsset A "
    sSQL = sSQL & "WHERE A.Asset_Park_Name Is Not Null"


   'NOTE I use QueryDef when referencing tables in a MS Access Database
    Set objRS_DAO = CurrentDb.CreateQueryDef("", sSQL).OpenRecordset


   'Create SQL string to get Tag List from PI Database
    sSQL = "Select tag, scan from pipoint where tag like " & vbCrLf


   'Where statement
    sSQL = sSQL & "'" & objRS_DAO.Fields(0).Value & ".%.Reli' " & vbCrLf
   
   
    Do While Not objRS_DAO.EOF

        sSQL = sSQL & " or tag like '" & objRS_DAO.Fields(0).Value & ".%.Reli' " & vbCrLf

        objRS_DAO.MoveNext

    Loop


    sSQL = sSQL & "and scan=1"
    
    
  ' Debug.Print sSQL_TagList

 
   '*** Create Tag List recordset
    objRS_TagList.Open sSQL, objConnOLEDB, adOpenKeyset, adLockBatchOptimistic
    
    
   'Count records returned
    lRecordCount = objRS_TagList.RecordCount
    
    
    'Debug.Print lRecordCount

   'Add recoreds from objRS_TagList to PR_TagList_RS
    objRS_TagList.MoveFirst
    
    
    lRowNumber = 0


    Do While Not objRS_TagList.EOF
  
        'Create row numbers
        lRowNumber = lRowNumber + 1
    
    
        With pub_objRS_TagList
        
        
            sAssetName = DLookup("[Asset]", "tblAsset", "[Asset_Park_Prefix] ='" & Left(objRS_TagList("tag").Value, 3) & "'")
   
            .AddNew fieldlist:=Array("Row_Number", "Tag_List", "Scan", "Asset_Name"), _
                values:=Array(lRowNumber, objRS_TagList("tag").Value, objRS_TagList("scan").Value, sAssetName)
                '.UpdateBatch  This is only need, if you are resyncing a recordset

                
        End With
    
    
        objRS_TagList.MoveNext
    
    
     Loop

  'Move the record to the first row
    pub_objRS_TagList.MoveFirst
    
    
'Set the recordset to the beginning
'Note setting the record set to the beginning is required
'  PubRS_1.MoveFirst
  

ProcExit:
   
    objRS_DAO.Close
    Set objRS_DAO = Nothing
   
    RS_TagList.Close
    Set RS_TagList = Nothing
    
    'Set objConnOLEDB = Nothing
    objConnOLEDB.Close

    Exit Sub

ProcErr:
  Select Case Err.Number
  
   Case 91, 424 'Object does not exist
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
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



End Sub

Private Sub PI_ADOX_Catalog()

    'Variables
     Dim sSQL As String

   'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection
    Dim objFld As ADODB.Field
    
    Dim objTbl As ADOX.Table
    Dim objView As ADOX.View
    Dim objProcedure As ADOX.Procedure
    
   'ADOX Catalog
    Dim objCatalog As ADOX.Catalog

    
'-----------Instantiate objects------------
  Set objConnOLEDB = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  Set objRS_TagList = New ADODB.Recordset
  Set objCatalog = New ADOX.Catalog
  
'--- Set Constants ----
  Const sFILEPATH = "\\porfiler02\shared\PPM\Asset Management - Wind\PhilS_AssetManagementWind\Dev_SandBox\SandBox_Reports\PI_Data_Report.xls"

  
 On Error GoTo ProcErr
  
 'OLEDB connection string to PI
  objConnOLEDB.ConnectionTimeout = 360
 
  objConnOLEDB.Open sPI_OLEDB
  
  Set objCatalog.ActiveConnection = objConnOLEDB
  
'
''List all tables in the connected database
'  For Each objTbl In objCatalog.Tables
'
'        Debug.Print "Table Name " & objTbl.Name
'
'  Next
  
    sSQL = "Select A.Tag_Name from ("
    sSQL = sSQL & " Select Left(tag,3) Tag_Name from pipoint"
    sSQL = sSQL & " where tag like '***.%.Reli' and scan=1"
    sSQL = sSQL & " Group by tag"
    sSQL = sSQL & " ) A"
    sSQL = sSQL & " Group by A.Tag_Name"


    'Open Recordset
     objRS.Open sSQL, objConnOLEDB, adOpenKeyset, adLockBatchOptimistic
     
     
     Debug.Print objRS.RecordCount
      
  
    '************  COPY RECORDSET TO SPREADSHEET  *********
     PI_CopyRecordset_to_Spreadsheet sFILEPATH, objRS


''*****************************************************************************
''Field in the Recordset object
'  i = 0
'
' For Each objFld In objRS.Fields
'
'    i = i + 1
'
'    Debug.Print i & " #" & objFld.Name
'
' Next



''Row values in Recordset Object
'   i = 0
'
'  Do While Not objRS.EOF
'
'    i = i + 1
'
'
'    Debug.Print i & " " & objRS.Fields(0).Value & " " & objRS.Fields(1).Value
'
'    objRS.MoveNext
'
'  Loop
  
  
ProcExit:

    
    objRS.Close
    Set objRS = Nothing
    
    objConnOLEDB.Close
    Set objConnOLEDB = Nothing
    
    DoCmd.Hourglass False

    Exit Sub

ProcErr:
  Select Case Err.Number

  Case 91 'Object does not exist
    'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & err.Number & vbCrLf & "Description " & err.Description & vbCrLf & vbCrLf & " The source " & err.Source, vbCritical
    'Stop
    Resume Next
    
  Case 3705   'Object Open
   MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
     
  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit
    
  Case -2147217900 'SQL syntax NOT correct
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
    
End Sub
Private Sub PI_All_TagName_and_TagList()

    'Variables
     Dim sSQL_TagName As String
     Dim sSQL As String

   'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection
    Dim objFld As ADODB.Field
    
    Dim objTbl As ADOX.Table
    Dim objView As ADOX.View
    Dim objProcedure As ADOX.Procedure
    
   'ADOX Catalog
    Dim objCatalog As ADOX.Catalog

    
'-----------Instantiate objects------------
  Set objConnOLEDB = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  Set objRS_TagList = New ADODB.Recordset
  Set objCatalog = New ADOX.Catalog
  
'--- Set Constants ----
  Const sFILEPATH = "\\porfiler02\shared\PPM\Asset Management - Wind\PhilS_AssetManagementWind\Dev_SandBox\SandBox_Reports\PI_Data_Report.xls"

  
 On Error GoTo ProcErr
  
 'OLEDB connection string to PI
  objConnOLEDB.ConnectionTimeout = 360
 
  objConnOLEDB.Open sPI_OLEDB
  
  Set objCatalog.ActiveConnection = objConnOLEDB
  

    sSQL_TagName = "Select A.Tag_Name from ("
    sSQL_TagName = sSQL_TagName & " Select Left(tag,3) Tag_Name from pipoint"
    sSQL_TagName = sSQL_TagName & " where tag like '***.%.Reli' and scan=1"
    sSQL_TagName = sSQL_TagName & " Group by tag"
    sSQL_TagName = sSQL_TagName & " ) A"
    sSQL_TagName = sSQL_TagName & " Group by A.Tag_Name"

 '   sSQL = sSQL_TagName
    

    sSQL_TagList = "Select tag Tag_List from pipoint"
    sSQL_TagList = sSQL_TagList & " where tag like '***.%.Reli' and scan=1"
    sSQL_TagList = sSQL_TagList & " Group by tag"
 
    
    sSQL = sSQL_TagList
    
    'Open Recordset
     objRS.Open sSQL, objConnOLEDB, adOpenKeyset, adLockBatchOptimistic
     
     
     Debug.Print objRS.RecordCount

  
    '************  COPY RECORDSET TO SPREADSHEET  *********
     PI_CopyRecordset_to_Spreadsheet sFILEPATH, objRS

  
ProcExit:

    
    objRS.Close
    Set objRS = Nothing
    
    objConnOLEDB.Close
    Set objConnOLEDB = Nothing
    
    DoCmd.Hourglass False

    Exit Sub

ProcErr:
  Select Case Err.Number

  Case 91 'Object does not exist
    'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
       
 Case 3021 'No Records returned
    MsgBox "Error Number " & Err.Number & vbCrLf & vbCrLf & " No records were returned!", vbExclamation
    Resume ProcExit
     
  Case 3704 'Recordset is already closed
    'MsgBox " The error # is " & err.Number & vbCrLf & "Description " & err.Description & vbCrLf & vbCrLf & " The source " & err.Source, vbCritical
    'Stop
    Resume Next
    
  Case 3705   'Object Open
   MsgBox "The Database had an error close and reopen the database" & vbCrLf & vbCrLf & "Error Number " & Err.Number, vbExclamation
    Resume ProcExit
     
  Case 3709 'Connection object isnt open
    MsgBox "Import failed! The data was NOT imported. Close and reopen the database.", vbInformation
    Resume ProcExit
    
  Case -2147217900 'SQL syntax NOT correct
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
    
End Sub
Public Function CreateRunningTotRS()
'The function fabricates a recordset, then poplulates the recordset from the "q_NetDP_Criteria" query definition
  Dim QD As QueryDef
  Dim i&, CurrentQty&, RunningTotal&, RecCount&
  Dim RS As New ADODB.Recordset  'NOTE: ALWAYS USE NEW RECORDSET
  Dim Conn As ADODB.Recordset

 On Error GoTo ProcErr

 DoCmd.Hourglass True

'This overwrites the instantance of Pubrs_1 and creates disconnect Recordset
  Set PubRS_1 = New ADODB.Recordset
  PubRS_1.CursorLocation = adUseClient

'Set RS to Nothing
  If Not RS Is Nothing Then
    Set RS = Nothing
  End If

'Fabricate Recordset Append Fields to recordset
  PubRS_1.Fields.Append "RunTot_ID", adBigInt
  PubRS_1.Fields.Append "QtyID", adBSTR
  PubRS_1.Fields.Append "Qty", adBigInt
  PubRS_1.Fields.Append "RunningTotal", adBigInt
 
'Create Recordset to append to Fabricated Recordset PubRS_1 NOTE the below query def q_NetDP_Season_Submit_a is based on NetDP table
  RS.Open "q_NetDP_Criteria", CurrentProject.Connection, adOpenKeyset, adLockBatchOptimistic
   
'Count records in recordset
  RecCount = RS.RecordCount
  
'Count to see if there are more than 3 records if NOT then Exit
  If RecCount < 4 Then
    MsgBox "There are only " & RecCount & " Styles"
    DoCmd.Hourglass False
    RS.Close
    Set RS = Nothing
    End
  End If
   
'  NOTE:If you use PubRS_1 to connect to the query def or table ie PubRS_1.Open ,
'       then you have to end the active connection as below
'       Set PubRS_1.ActiveConnection = Nothing

'Add value from "q_NetDP_Criteria" to the fabricated recordset
  PubRS_1.Open
  Do While Not RS.EOF
    RunningTotal = RunningTotal + RS("qty").Value
    With PubRS_1
        .AddNew fieldlist:=Array("RunTot_ID", "QtyID", "Qty", "RunningTotal"), _
            values:=Array(RS.AbsolutePosition, RS("QtyID").Value, RS("qty").Value, RunningTotal)
            '.UpdateBatch  This is only need, if you are resyncing a recordset
    End With
    RS.MoveNext
  Loop
  
'Set the recordset to the beginning
'Note setting the record set to the beginning is required
  PubRS_1.MoveFirst
  
ProcExit:
  DoCmd.Hourglass False
  RS.Close
  Set RS = Nothing
Exit Function

ProcErr:
  DoCmd.Hourglass False
  Select Case Err.Number
  Case -2147217900 'Missing SQL Statement
    Resume ProcExit
  Case 3021 'BOF or EOF not found
    DoCmd.Hourglass False
    RS.Close
    Set RS = Nothing
    End
  Case 3704 'Recordset empty End program to stop more errors
    End
  Case 3151
    If MsgBox("Wrong PassWord !" & vbCrLf & "Click on Yes to re-enter Password or No to Exit", vbYesNo + vbCritical) = vbYes Then
        Resume
    Else
        Resume ProcExit
    End If
  Case Else
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    Resume Next
  End Select
Resume ProcExit

End Function

