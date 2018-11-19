Option Explicit
Public Sub Run_PullFrom_Database()

  ConnectACCDB_viaADODB

End Sub
Public Function ConnectACCDB_viaADODB()


'*--------------------------------------------------------------------------------------------------
'*
'*  NOTE: (ConnectionString.com) https://www.connectionstrings.com/ace-oledb-12-0/
'*        https://github.com/GenerationTRS80/Projects-with-ADODB/blob/master/AccessCode/PI_DATA.vb
'*



 'ADO objects
  Dim Conn As ADODB.Connection
  Dim ConnMdb As ADODB.Connection
  Dim rsTable As ADODB.Recordset
  
 'Local variables
  Dim sConnectionString As String
  Dim sConnectionString_Mdb As String
  Dim sFilePath As String
  Dim sFilePath_Mdb As String
  Dim sSQL As String
  Dim sSQL_Mdb As String
  Dim i As Integer
  
  
 On Error GoTo ProcErr
  
 '-----------Instantiate objects------------
  Set Conn = New ADODB.Connection
  Set ConnMdb = New ADODB.Connection
  Set rsTable = New ADODB.Recordset
 
 'Set file path and database name
  sFilePath = "C:\Users\micro\Desktop\Code\Access365_Database\Template_Production.accdb"
  sFilePath_Mdb = "C:\Users\micro\Desktop\As_Traded_Database.mdb"
  
  
 'Create connection string
  sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & vbCrLf
  sConnectionString = sConnectionString & "Data Source=" & sFilePath & ";" & vbCrLf
  
 'Create connection string (to an Access mdb database)
  sConnectionString_Mdb = "Provider=Microsoft.ACE.OLEDB.12.0;" & vbCrLf
  sConnectionString_Mdb = sConnectionString_Mdb & "Data Source=" & sFilePath_Mdb & ";" & vbCrLf
  
 'Connect Database to Recordset
  With Conn
    .CursorLocation = adUseClient
    .CommandTimeout = 10
    .Open sConnectionString
  End With
   
  With ConnMdb
    .CursorLocation = adUseClient
    .CommandTimeout = 10
    .Open sConnectionString_Mdb
  End With
   
  Debug.Print "Connection String " & sConnectionString_Mdb & vbCrLf & vbCrLf
  
 'Write SQL statment
  sSQL = "SELECT Products.ID, Products.ProductCode, Products.ProductName" & vbCrLf
  sSQL = sSQL & ", Products.StandardCost, Products.ListPrice, Products.Discontinued" & vbCrLf
  sSQL = sSQL & " FROM Products"
  
 'Write SQL for AS_Traded_Database.mdb
  sSQL_Mdb = "SELECT tblGeneration.Generation_Key, tblGeneration.Asset, tblGeneration.Delivery_Date, "
  sSQL_Mdb = sSQL_Mdb & "tblGeneration.Delivery_Location, tblGeneration.Forecast_Generation" & vbCrLf
  sSQL_Mdb = sSQL_Mdb & " FROM tblGeneration"
  
  Debug.Print sSQL_Mdb

 'Open Recordset
 ' rsTable.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly

 'Access Mdb database
  rsTable.Open sSQL_Mdb, ConnMdb, adOpenForwardOnly, adLockReadOnly
  
  

 'Set interger value for Loop to 0
  i = 0
 
 'Write the values in the Recordset in the Immediate window
  Do While Not rsTable.EOF
  
    i = i + 1
    
    Debug.Print i & " " & rsTable.Fields(0).Value & " " & rsTable.Fields(1).Value & " " & rsTable.Fields(2).Value

    rsTable.MoveNext
  
  Loop
  
  rsTable.MoveFirst

ProcExit:
  
 'Close connection objects
  Conn.Close
  Set Conn = Nothing
  
  ConnMdb.Close
  Set ConnMdb = Nothing


 'Close Recordset
  rsTable.Close
  Set rsTable = Nothing
    

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
  
    
  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit
  

End Function

