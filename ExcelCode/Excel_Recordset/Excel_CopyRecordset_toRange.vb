Public Function Sub_CopyRangeXML_ToRecordset(xlWrkBk_CORP As Excel.Workbook, sPull_WorksheetName As String) As Boolean


'COM Objects
  Dim xDocModelTab As Object
  Dim rsModelTab As ADODB.Recordset
  Dim rsFilter As ADODB.Recordset
  Dim Fld As ADODB.Field
 
'Excel objects
  Dim xlWrkSht_ModelTab As Excel.Worksheet
  Dim rngModelTab As Excel.Range

'Local Variables
  Dim sAddress_ModelTab As String
  Dim sFilterString As String
   
'Constants
  Const MODEL_TAB_START_ROWNUM = 6
  Const MODEL_TAB_START_COLUMNNUM = 2
  Const MODEL_TAB_ROWCOUNT = 400
  Const MODEL_TAB_COLUMNCOUNT = 66
   
  
 'Set Report FTE_Count as TRUE
  Sub_CopyRangeXML_ToRecordset = True

  
 On Error GoTo ProcErr

 'Instantiate Model an Dom objects
  Set rsModelTab = New ADODB.Recordset
  Set xDocModelTab = CreateObject("MSXML2.DOMDocument")
  
 'Set worksheet object
  Set xlWrkSht_ModelTab = xlWrkBk_CORP.Worksheets(sPull_WorksheetName)
  
 'Get cells C1 through BH397 from CMI tab
 '*** NOTE: using these references instead of a name range allow to select historical cost models without a name range created in them ***
  Set rngModelTab = xlWrkSht_ModelTab.Range(xlWrkSht_ModelTab.Cells(MODEL_TAB_START_ROWNUM, MODEL_TAB_START_COLUMNNUM), _
                                                                    xlWrkSht_ModelTab.Cells(MODEL_TAB_ROWCOUNT, MODEL_TAB_COLUMNCOUNT))

 'Load range into XML object
  xDocModelTab.LoadXML rngModelTab.Value(xlRangeValueMSPersistXML)
  
 'Instantiate Public Recordsets used in the subroutine
  Set rsPUBLIC_Report_FTE_LineItems = New ADODB.Recordset
  Set rsFilter = New ADODB.Recordset

 '*** Disconnect the Public Recordsets ***
  rsModelTab.CursorLocation = adUseClient
  rsFilter.CursorLocation = adUseClient
  

 'Open recordset from XML
  rsModelTab.Open xDocModelTab, , adOpenStatic, adLockBatchOptimistic

 'Populate PUBLIC Recordsets with clone method
  Set rsPUBLIC_Report_FTE_LineItems = rsModelTab.Clone

 '-----------------------Find FTE Line items-----------
 'NOTE: Use the last field (Column BM) to filter
  sFilterString = rsPUBLIC_Report_FTE_LineItems.Fields(rsPUBLIC_Report_FTE_LineItems.Fields.Count - 1).Name & "='1'"
  rsPUBLIC_Report_FTE_LineItems.Filter = sFilterString
  
 'Get Records for the number of rows returned
  Pub_RowCount_Report_LineItems = rsPUBLIC_Report_FTE_LineItems.RecordCount


    

ProcExit:

'Close Recordset
 rsModelTab.Close
 Set rsModelTab = Nothing
 

 rsFilter.Close
 Set rsFilter = Nothing

 Exit Function

ProcErr:

  Select Case Err.Number
  
  Case 9 'Description Subscript out of range
    Sub_CopyRangeXML_ToRecordset = False
    bClose_SP_Workbook = False
    MsgBox " This Cost Model does not have a Corp Model Input tab." & vbCrLf & vbCrLf & "Copy the data directly from the model into the appropriate section of the Export Corp tab!", vbInformation + vbOKOnly, "Corp Model Input NOT in this Cost Model"
    xlWrkBk_SP.Activate
    
    Resume ProcExit

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset is already closed
    Resume Next
    
  Case -2147467259 'Steam Object can't be read because it is empty
    Sub_CopyRangeXML_ToRecordset = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
    Resume ProcExit

  Case Else
    Sub_CopyRangeXML_ToRecordset = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function