Private Sub Test_Connection()
    Dim sConn As String
    Dim sSQL As String
    Dim i As Integer
    
   'ADO Objects
    Dim objRS As ADODB.Recordset
    Dim objConnOLEDB As ADODB.Connection
    
    Dim QD As QueryDef
   
  On Error GoTo ProcErr
    
'-----------Instantiate objects------------
  Set objConnOLEDB = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  
  Debug.Print sWebTrader_OLEDB & vbCrLf & vbCrLf
  

 'OLEDB connection string to Oracle ZaiNetEOD
 
  objConnOLEDB.ConnectionTimeout = 60
 
  objConnOLEDB.Open sWebTrader_OLEDB
  
  sSQL = "Select * from dbo.Company"
  
 ' qd.OpenRecordset.Connection
  
 'Open Recordset
  objRS.Open SQL_WebTrader_DayAhead_HourAhead, objConnOLEDB, adOpenForwardOnly, adLockReadOnly

 i = 0
 
  Do While Not objRS.EOF
  
    i = i + 1
    
    Debug.Print i & " " & objRS.Fields(1).Value & " " & objRS.Fields(8).Value

    objRS.MoveNext
  
  Loop
  
  
ProcExit:

    Exit Sub

ProcErr:
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
Private Sub Run_Subroutines()

  'CreateSPT "q_DNSLESS_WebTrader", SQL_WebTrader_DayAhead_HourAhead, sWebTrader_ODBC

  ' Debug.Print SQL_WebTrader_DayAhead_HourAhead

End Sub
Private Function SQL_WebTrader_DayAhead_HourAhead()
 Dim sSQL As String
 
 
 sSQL = "select t1.AlternateName, t1.Name, " & vbCrLf
 sSQL = sSQL & "Case When Left(t1.TypeName,4)='Best' Then 'Best' else Substring(t1.Name,7+Len(t1.AlternateName),2) End as Forecast_DA_HA, " & vbCrLf
 sSQL = sSQL & "t1.TypeName, " & vbCrLf
 sSQL = sSQL & "Cast(Substring(CONVERT(varchar,t2.Date),5,2)+'/'+Right(CONVERT(varchar,t2.Date),2)+'/'+Left(CONVERT(varchar,t2.Date),4)as datetime) as Transaction_Date, t2.Date as Date_String, " & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW5,0) as bigint)  as HE1,Cast(ISNULL(t2.MW6,0) as bigint)  as HE2 ,Cast(ISNULL(t2.MW7,0) as bigint)  as HE3, " & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW8,0) as bigint) as HE4, Cast(ISNULL(t2.MW9,0) as bigint) as HE5 ,Cast(ISNULL(t2.MW10,0) as bigint) as HE6, " & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW11,0) as bigint) as HE7,Cast(ISNULL(t2.MW12,0) as bigint) as HE8 ,Cast(ISNULL(t2.MW13,0) as bigint)  as HE9 ," & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW14,0) as bigint) as HE10,Cast(ISNULL(t2.MW15,0) as bigint) as HE11,Cast(ISNULL(t2.MW16,0) as bigint) as HE12, " & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW20,0) as bigint) as HE16,Cast(ISNULL(t2.MW21,0) as bigint) as HE17,Cast(ISNULL(t2.MW22,0) as bigint) as HE18, " & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW23,0) as bigint) as HE19,Cast(ISNULL(t2.MW24,0) as bigint) as HE20,Cast(ISNULL(t2.MW25,0) as bigint) as HE21," & vbCrLf
 sSQL = sSQL & "Cast(ISNULL(t2.MW26,0) as bigint) as HE22,Cast(ISNULL(t2.MW27,0) as bigint) as HE23,Cast(ISNULL(t2.MW28,0) as bigint) as HE24 " & vbCrLf
 
 
' sSQL = sSQL & "ISNULL(t2.MW23,0) as HE19,ISNULL(t2.MW24,0) as HE20,ISNULL(t2.MW25,0) as HE21,ISNULL(t2.MW26,0) as HE22,ISNULL(t2.MW27,0) as HE23,ISNULL(t2.MW28,0) as HE24 " & vbCrLf
 sSQL = sSQL & "from ppm_operational.dbo.WT_Resource t1, " & vbCrLf
 sSQL = sSQL & "ppm_operational.dbo.Unit_Daily_Hourly t2 " & vbCrLf
 sSQL = sSQL & "where t1.Id = t2.Id " & vbCrLf
 sSQL = sSQL & "and t2.Date  >= Cast(Convert(varchar, GetDate()-10,112) as bigint) " & vbCrLf
 sSQL = sSQL & "and t2.Date  <= Cast(Convert(varchar, GetDate()-1,112) as bigint) " & vbCrLf
 sSQL = sSQL & "and t2.DSTAction = 0" & vbCrLf
 sSQL = sSQL & "and t1.AlternateName in ('HIWIND', 'BHORN', 'KLND2', 'KLND3', 'SHILO') " & vbCrLf
 sSQL = sSQL & "Order by t1.AlternateName,t1.Name"
 
'--and t1.TypeName = 'Forecast'


SQL_WebTrader_DayAhead_HourAhead = sSQL

End Function
Function LinkToPubsAuthorsDSNLess()
    'This example comes from DataFast consulting webpage
    'http://www.amazecreations.com/datafast/ShowArticle.aspx?File=Articles/odbctutor01.htm


    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    
    
    Dim strServer As String
    Dim strDatabase As String
    Dim strUID As String
    Dim strPWD As String
    
    'strServer = "208.254.145.33"
    
    strServer = "PPMClient"
    
    strDatabase = "ppm_operational"
    
    strUID = "PPMClient"
    
    strPWD = "P5EPheyu"
    
    On Error GoTo ProcErr

        strConnect = "ODBC;DRIVER={SQL Server}" _
                    & ";SERVER=" & strServer _
                    & ";DATABASE=" & strDatabase _
                    & ";UID=" & strUID _
                    & ";PWD=" & strPWD & ";"
                    
        Debug.Print strConnect

        Set db = CurrentDb()
        Set tdf = db.CreateTableDef("Link_WebTrader_ODBC")
        tdf.SourceTableName = strConnect

        tdf.Connect = Trim(strConnect)

        db.TableDefs.Append tdf
        db.TableDefs.Refresh

        Set tdf = Nothing
        Set db = Nothing
        

ProcExit:

    Exit Function

ProcErr:
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


Public Function CreateSPT(SPTQueryName As String, strSQL As String, sOLEDB_Connection_String As String)
'Create OLEDB Pass Through Query connection

  Dim cat As ADOX.Catalog
  Dim cmd As ADODB.Command

  Set cat = New ADOX.Catalog
  Set cmd = New ADODB.Command

  On Error GoTo ProcErr
  
  cat.ActiveConnection = CurrentProject.Connection

  Set cmd.ActiveConnection = cat.ActiveConnection

  cmd.CommandText = strSQL
  cmd.Properties("Jet OLEDB:ODBC Pass-Through Statement") = True

 'Modify the following connection string to reference an existing DSN for
 'the sample SQL Server PUBS database.

' cmd.Properties("Jet OLEDB:Pass Through Query Connect String") = "ODBC;DSN=myDSN;database=pubs;UID=sa;PWD=;"
  cmd.Properties("Jet OLEDB:Pass Through Query Connect String") = sOLEDB_Connection_String

  cat.Procedures.Append SPTQueryName, cmd

  Set cat = Nothing
  Set cmd = Nothing


ProcExit:

    Exit Function

ProcErr:
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

