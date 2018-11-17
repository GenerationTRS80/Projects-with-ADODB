Private Sub Test_FabricatedRecordset_TypePrecision()

 'ADO objects
  Dim Fld As ADODB.Field
  Dim cmd As ADODB.Command
  Dim Prm As ADODB.Parameter
  Dim rsTest_PublicFunction_SetDataType As ADODB.Recordset
  
  'Local Variable
  Dim sCMIWorksheet_LineItems_Value As String
  Dim sLineItem_Value As String
  Dim lFieldType As Long
  Dim lFieldPrecision As Long
  Dim iFieldNumber As Integer
  Dim sFieldName As String
  
  
 'Instantiate recordsets
  Set rsTest_PublicFunction_SetDataType = New ADODB.Recordset

 'Disconnect the Public Recordsets
  rsTest_PublicFunction_SetDataType.CursorLocation = adUseClient

 'Set fabricated FTE RECORDSET
   Set rsTest_PublicFunction_SetDataType = PubFN_rsSetDataType_FTEs.Clone

    'Get values from each field of the recordset pass it to update the recordset rsTest_PublicFunction_SetDataType
      For Each Fld In rsTest_PublicFunction_SetDataType.Fields
      
          sFieldName = Fld.Name
          lFieldType = Fld.Type
          lFieldPrecision = Fld.Precision
          
          Debug.Print sFieldName & " Field Type " & lFieldType & " Field Precision  " & lFieldPrecision
      
      Next

End Sub

'******    This is the function that fabricates the recordset createding field names and field types ******

Public Function PubFN_rsSetDataType_FTEs() As ADODB.Recordset


'******************************************************************************
'*
'* This function creates a public fabricated recordset for FTE items
'*
'*


'Local variable
  Dim rsSetDataType As ADODB.Recordset
  
'Instantiate recordsets
  Set rsSetDataType = New ADODB.Recordset

 'Disconnect the Public Recordsets
  rsSetDataType.CursorLocation = adUseClient

 'Fabricate Recordset Append Fields to recordset
    rsSetDataType.Fields.Append "CMI_LineItem", adInteger
    rsSetDataType.Fields.Append "No Header FTEs 01", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header FTEs 02", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header FTEs 03", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Location", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Job Title Description", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Annual Salary", adBigInt, , adFldIsNullable
    
    rsSetDataType.Fields.Append "Wage Esc 1_3", adDouble, , adFldIsNullable
    'rsSetDataType.Fields("Wage Esc 1_3").Precision = 4
    
    rsSetDataType.Fields.Append "Wage Esc 4_10", adDouble, , adFldIsNullable
    'rsSetDataType.Fields("Wage Esc 4_10").Precision = 4
      
    rsSetDataType.Fields.Append "Standard Benefits Load", adDouble, , adFldIsNullable
    'rsSetDataType.Fields("Standard Benefits Load").Precision = 4
    
    rsSetDataType.Fields.Append "Bonus", adDouble, , adFldIsNullable
   ' rsSetDataType.Fields("Bonus").Precision = 4
    
    rsSetDataType.Fields.Append "Monthly Employee Related Expense", adBigInt, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 01", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 02", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 03", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 04", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 05", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 06", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 07", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "No Header 08", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 1", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 2", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 3", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 4", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 5", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 6", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 7", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 8", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 9", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 10", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 11", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 12", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 1", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 13", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 14", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 15", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 16", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 17", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 18", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 19", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 20", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 21", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 22", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 23", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Month 24", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 2", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 3", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 4", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 5", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 6", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 7", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 8", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 9", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 10", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Year 11", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Total", adDouble, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BF", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BG", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BH", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BI", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BJ", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BK", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BL", adBSTR, , adFldIsNullable
    rsSetDataType.Fields.Append "Column BM", adInteger, , adFldIsNullable

    rsSetDataType.Open
    
    Set PubFN_rsSetDataType_FTEs = rsSetDataType.Clone
    
    rsSetDataType.Close
    
    Set rsSetDataType = Nothing


End Function