Public Function adPOPULATE_rsPUBLIC_CMIworksheet_LineItems(rsCMIworksheet As ADODB.Recordset, lStartRow As Long) As Boolean

 'ADO objects
  Dim Fld As ADODB.Field
  Dim cmd As ADODB.Command
  Dim Prm As ADODB.Parameter


 'Local Variables
  Dim lRowNumber As Long
  Dim lFieldCount As Long
  Dim i As Integer
  
 'Constants
  Const KEY_FIELD_NAME = "CMI_ID"
  
 'Set function to TRUE
  adPOPULATE_rsPUBLIC_CMIworksheet_LineItems = True
  
  
  On Error GoTo ProcErr
  
 'Instantiate New Recordset
  Set rsPUBLIC_CMIworksheet_LineItems = New ADODB.Recordset
  Set cmd = New ADODB.Command
  Set Prm = New ADODB.Parameter
  
 'To disconnect recordset us cursor location client
  rsPUBLIC_CMIworksheet_LineItems.CursorLocation = adUseClient
  
  
 '-------->>> Append ID field CMI_ID to recordset <<<-----------
  rsPUBLIC_CMIworksheet_LineItems.Fields.Append KEY_FIELD_NAME, adBigInt

 'Get field names
  For Each Fld In rsCMIworksheet.Fields
  
    'Append fields from recordset
     rsPUBLIC_CMIworksheet_LineItems.Fields.Append Fld.Name, Fld.Type, Fld.DefinedSize, Fld.Attributes
        
  Next

 'Get count of fields
  lFieldCount = rsPUBLIC_CMIworksheet_LineItems.Fields.Count
  
  
 'OPEN Recordset CMIworksheet_LineItems
  rsPUBLIC_CMIworksheet_LineItems.Open


 'Set beginning row
  lRowNumber = lStartRow - 1
  
  
 '------------ >>> Add records to rsPUBLIC_CMIworksheet_LineItems <<<-----------
  Do While Not rsCMIworksheet.EOF
  
   'Add New Record to fabricated recordset
    rsPUBLIC_CMIworksheet_LineItems.AddNew
  
   'Increment RowNumber
    lRowNumber = lRowNumber + 1
  
  
   'Get values from each field of the recordset passed to update the current record of the fabricated recordset
    For Each Fld In rsPUBLIC_CMIworksheet_LineItems.Fields
    
       'Append KEY_FIELD_NAME value all other fields are appended after the else
        If Fld.Name = KEY_FIELD_NAME Then
        
            Set Prm = cmd.CreateParameter(, Fld.Type, adParamInput, Fld.DefinedSize, lRowNumber)
            
        Else
        
           '--->> Create parameter and append to cmd object<<---
           'NOTE This ensure that the data coming from the Form is the right data type and right data size
           'For Example: If a value is copied into a cell that is read by the recordset does not require that cell to maintain the type of the copied value
            Set Prm = cmd.CreateParameter(, Fld.Type, adParamInput, Fld.DefinedSize, rsCMIworksheet(Fld.Name).Value)
        
        End If
       

        '--->> Set Parameter Value to fields value <<---
        'rsPUBLIC_CMIworksheet_LineItems(Fld.Name).Value = Prm.Value
         Fld.Value = Prm.Value
      
     Next
     
     
    'Update all fields in the current records
     rsPUBLIC_CMIworksheet_LineItems.UpdateBatch
       
    'Move to next record for both recordsets
     rsCMIworksheet.MoveNext
    
  Loop

 
 'Move the first record
  rsPUBLIC_CMIworksheet_LineItems.MoveFirst

ProcExit:

''Close Recordset
 rsCMIworksheet.Close
 Set rsCMIworksheet = Nothing
 

 Exit Function

ProcErr:

  Select Case Err.Number

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset is already closed
    Resume Next
    
  Case Else
    adPOPULATE_rsPUBLIC_CMIworksheet_LineItems = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function