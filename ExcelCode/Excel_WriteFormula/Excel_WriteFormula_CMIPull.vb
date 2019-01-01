Public Function abReport_SQLquery_ModelWorksheet(xlWrkBk_CORP As Excel.Workbook, _
                                                        sPull_WorksheetName As String, _
                                                        Optional lSubTowerNumber As Integer = 1, _
                                                        Optional sModelFileName As String, _
                                                        Optional sTowerName As String) As Boolean


   'Local Variables
    Dim sSQL As String
    Dim sPull_Worksheet_CellAddress As String
    Dim sFilePathName As String
    Dim sTarget_WorksheetName As String
    
   'Excel Objects
    Dim xlWrkSht_ModelTab As Excel.Worksheet
    Dim rngPullWorksheet As Range
    
   'ADO Objects
    Dim Fld As ADODB.Field
    Dim Rec As ADODB.Record
    
   'Constants
    Const MODEL_TAB_START_ROWNUM = 16
    Const MODEL_TAB_START_COLUMNNUM = 2
    Const MODEL_TAB_ROWCOUNT = 600
    Const MODEL_TAB_COLUMNCOUNT = 66
    
   'Default Report_CreateRecordset_Tower_LineItems to true
    abReport_SQLquery_ModelWorksheet = True
    
    
   On Error GoTo ProcErr
    
    Set xlWrkSht_ModelTab = xlWrkBk_CORP.Worksheets(sPull_WorksheetName)
    sFilePathName = xlWrkBk_CORP.FullName
    
   'Get cells C1 through BH397 from CMI tab
   'NOTE: using these references instead of a name range allow to select historical cost models without a name range created in them
    Set rngPullWorksheet = xlWrkSht_ModelTab.Range(xlWrkSht_ModelTab.Cells(MODEL_TAB_START_ROWNUM, MODEL_TAB_START_COLUMNNUM), _
                                                                    xlWrkSht_ModelTab.Cells(MODEL_TAB_ROWCOUNT, MODEL_TAB_COLUMNCOUNT))
    
    sPull_Worksheet_CellAddress = "[" & sPull_WorksheetName & "$" & rngPullWorksheet.Address(False, False) & "]"
    'Debug.Print "Pull cell address " & sPull_Worksheet_CellAddress
    
    
  'Set Header
    Set rngHeader = Range(Cells(rngFirstCell.Row - 1, rngFirstCell.Column), Cells(rngFirstCell.Row - 1, rngLastCell.Column - 1))
   

 '-------------------------------------------------------------------------------
 '*
 '*  Create SQL query from worksheet FTE Line Item Reports
 '*

    sSQL = "Select Clng([F1]) As [Row Number]" & vbCrLf & vbCrLf
    
    
   'Model and File name
    sSQL = sSQL & ", IIf(Len(""" & sModelFileName & """) = 0, Null, """ & sModelFileName & """) as Model " & vbCrLf
    sSQL = sSQL & ", IIf(Len(""" & sTowerName & """) = 0, Null, """ & sTowerName & """) as [Tower Name] " & vbCrLf & vbCrLf
    
   'Location and Job Description
    sSQL = sSQL & ", IIf([F5]=""U.S. - ACS Site"",""United States"",[F5]) as Location, [F6] as [Job Description]" & vbCrLf
    sSQL = sSQL & ", CCur([F7]) As [Base], CDbl([F8]) as [Wage Esc 1-3], CDbl([F9]) as [Wage Esc 4-10], CDbl([F10]) as [Standard Benefit Load], CDbl(Format([F11],""0.##"")) as Bonus, CCur([F12]) as [Employee Related Expense]" & vbCrLf & vbCrLf
    
   'Labor Load rates
    sSQL = sSQL & ", Null as [Labor Load Rates]" & vbCrLf
    
    
   '-----------------------------------------------------------------------
   ' NOTES for the loaded rates are the actual Excel formulas

     sSQL = sSQL & ", CCur(" & vbCrLf      'Annual Salary Escalation
     sSQL = sSQL & "[F7]*((1+[F8])^0)*((1+[F9])^0)" & vbCrLf      'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "(CCur([F7])*((1+[F8])^0)*((1+[F9])^0))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                            'Employee Related Expenses (Note: Expense is in Months, Need to multiply by 12
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^0)*((1+[F9])^0))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 1] "
 
    'Wage Escalation Year 2
    
     sSQL = sSQL & ", CCur(" & vbCrLf      'Annual Salary Escalation
     sSQL = sSQL & "[F7]*((1+[F8])^1)*((1+[F9])^0)" & vbCrLf      'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "(CCur([F7])*((1+[F8])^1)*((1+[F9])^0))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                            'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^1)*((1+[F9])^0))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 2] "
     
    'FTE Salary Year 3
     
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^0)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^0))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^0))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 3] "
    
    'FTE Salary Year 4
     
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^1)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^1))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^1))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 4] "

    
    'FTE Salary Year 5
      
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^2)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^2))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12] * 12"                                         'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^2))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 5] "
      
    'FTE Salary Year 6
      
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^3)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^3))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^3))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 6] "
        
    'FTE Salary Year 7
      
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^4)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^4))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^4))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 7] "

    'FTE Salary Year 8
      
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^5)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^5))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^5))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 8] "
     
    'FTE Salary Year 9
      
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^6)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^6))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^6))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 9] "
     
    'FTE Salary Year 10
  
     sSQL = sSQL & ", CCur(" & vbCrLf
     sSQL = sSQL & "[F7]*((1+[F8])^2)*((1+[F9])^7)" & vbCrLf            'Annual Salary Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "([F7]*((1+[F8])^2)*((1+[F9])^7))*[F10]" & vbCrLf    'Benifit Escalation
     sSQL = sSQL & "+"
     sSQL = sSQL & "[F12]*12"                                           'Employee Related Expenses
     sSQL = sSQL & "+"
     sSQL = sSQL & "(([F7]*((1+[F8])^2)*((1+[F9])^7))*[F11])" & vbCrLf  'Bonus Escalation
     sSQL = sSQL & ") as [Wage Escalation Year 10] "
     
   'Aunnual Rate Per FTE
    sSQL = sSQL & ", Null as [Annual Rate Per FTE]" & vbCrLf & vbCrLf
    
  
   'Month 13 through 24
    sSQL = sSQL & ",[F21] as [Month 1], [F22] as [Month 2], [F23] as [Month 3], [F24] as [Month 4], [F25] as [Month 5], [F26] as [Month 6] " & vbCrLf
    sSQL = sSQL & ",[F27] as [Month 7], [F28] as [Month 8], [F29] as [Month 9], [F30] as [Month 10], [F31] as [Month 11], [F32] as [Month 12] " & vbCrLf & vbCrLf
    
   'Year 1
    sSQL = sSQL & ",[F33] as [Year 1] " & vbCrLf
    
   'Month 13 through 24
    sSQL = sSQL & ",[F34] as [Month 13], [F35] as [Month 14], [F36] as [Month 15], [F37] as [Month 16], [F38] as [Month 17], [F39] as [Month 18] " & vbCrLf
    sSQL = sSQL & ",[F40] as [Month 19], [F41] as [Month 20], [F42] as [Month 21], [F43] as [Month 22], [F44] as [Month 23], [F45] as [Month 24] " & vbCrLf & vbCrLf
    
   'Year 2
    sSQL = sSQL & ",[F46] as [Year 2]" & vbCrLf
    
   'Year 3
    sSQL = sSQL & ",[F47] as [Year 3]" & vbCrLf
    
   'Year 4
    sSQL = sSQL & ",[F48] as [Year 4]" & vbCrLf
    
   'Year 5
    sSQL = sSQL & ",[F49] as [Year 5]" & vbCrLf
     
   'Year 6
    sSQL = sSQL & ",[F50] as [Year 6]" & vbCrLf
    
   'Year 7 through 10 and Total
    sSQL = sSQL & ",[F51] as [Year 7], [F52] as [Year 8], [F53] as [Year 9], [F54] as [Year 10]" & vbCrLf
    sSQL = sSQL & ",CDbl([F55]) as [Year 11], CDbl([F56]) as [Total]" & vbCrLf
    
   'Column BM and BN Filter Code and Sub Tower Number
    sSQL = sSQL & ",Clng([F64]) as [Filter Code]" & vbCrLf
    sSQL = sSQL & ",Clng([F65]) as [Sub tower Number]" & vbCrLf
    

  'Below is the WHERE Statement for the SQL Clause **
  'NOTE: Column BM filter needs to equal 1 (FTE line item) or 18 (Total FTE Labor Expense )
  
    sSQL = sSQL & "from " & sPull_Worksheet_CellAddress & vbCrLf
    sSQL = sSQL & "Where ([F64]=1 or [F64]=18) and [F65]=" & CLng(lSubTowerNumber) & ""
    
    
   'sSQL = sSQL & "Where ([F64]=1 or [F64]=18) or ([F64]=99 and [F33] is null)"
   'sSQL = sSQL & "Where ([F64]=1 or [F64]=18) or ([F64]=18 and [F33] is not null)"


 '-------------------------------------------------------------------------------
 '*
 '*  Execute SQL and write it to a Recordet
 '*

    If Report_CreateRecordset_Tower_LineItems(sFilePathName, sSQL) = False Then

        abReport_SQLquery_ModelWorksheet = False
        Debug.Print "**** Error Exited Report_CreateRecordset_Tower_LineItems Subroutine ******"
        GoTo ProcExit
        
    Else
    
    
        'Debug.Print vbCrLf & "Succeded Import!"

    End If

  'Check to see if any records were returned, if there were 0 records return exit the program
 'Get Records for the number of rows returned
  Pub_RowCount_Report_LineItems = rsPUBLIC_Report_FTE_LineItems.RecordCount
  Pub_ColumnCount_Report_LineItems = rsPUBLIC_Report_FTE_LineItems.Fields.Count
  
  rsPUBLIC_Report_FTE_LineItems.MoveFirst


ProcExit:



Exit Function

ProcErr:

  Select Case Err.Number
    
   Case 5
    abReport_SQLquery_ModelWorksheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # in abReport is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
  
  Case 9
    abReport_SQLquery_ModelWorksheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # in abReport is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
   ' Debug.Print " The error # in abReport is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case 3021 'Zero recordset returned
   'Debug.Print " The error # in abReport is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume Next
  
  Case 3265 'Item cannot be found
    abReport_SQLquery_ModelWorksheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # in abReport is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  Case 3704 'Recordset is already closed
   'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Stop
    Resume Next

  End Select
  
  Resume ProcExit

End Function