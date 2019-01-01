Public Sub aaReport_Main(xlWrkBk_CORP As Excel.Workbook)


 '-------------------------------------------------------------------------------------------------------------
 '
 '       Take CMI Pull data from Model worksheet 1 through 4 and do the following tasks:
 '

   'Excel Variables
    Dim xlWrkSht_PDL As Excel.Worksheet
    Dim rngFileName As Excel.Range
    Dim rngTower As Excel.Range
    Dim rngWorksheetName As Excel.Range
    Dim rngTemp As Range
    
   'Local variables
    Dim sWorksheet_Tab As String
    Dim sModelFileName As String
    Dim sTowerName As String
    Dim lStartRow_FTE_Header As Long
    Dim lStartRow_FTE_LineItems As Long
    Dim lStartColumn_WriteFormula_PxQ As Long
    Dim lCountofRows_FTE_LineItems As Long
    Dim iRange_RowNumber As Integer
    Dim iSubtowerNumber As Integer
    
    Const lHEADER_ROWS = 17
    Const lFIRST_COLUMN_REFERENCE_NUMBER = 3
    
    Const lCOUNT_OF_SUBTOWERS = 6


   On Error GoTo ProcErr
   

   'Set object to Pull Down List worksheet
    Set xlWrkSht_PDL = xlWrkBk_CORP.Worksheets("Pull Down Lists")
    
    Set rngWorksheetName = xlWrkSht_PDL.Range("vba_NmRng_PDL_WorksheetName")
    Set rngFileName = xlWrkSht_PDL.Range("vba_NmRng_PDL_FileName")
    Set rngTower = xlWrkSht_PDL.Range("vba_NmRng_PDL_Tower")

'**************************************************************************************************************
'       Clear FTE Report Tab
'

 Sub_Report_Clear_DataFormats xlWrkBk_CORP, "FTE Report", True, FTE_REPORT_HEADER_START_ROWNUM - 2, , , 110
  
  

'-------------------------------------------------------------------------------------------------------------
'
'       Take CMI Pull data from Model worksheet 1 through 4 and do the following tasks:
'
'            1) Filter for FTE line items (where the data in Model tab on column BM=1)
'            2) Calculate the Monthly and Yearly Escalation by FTE Line item
'            3) Write formulas:
'                   a) PxQ calculations
'                   b) time adjustments for pre go-live Corp Model Input report
'


'Start row and column for FTE line items, PxQ formulas, Time adjustment formulas
 lStartRow_FTE_Header = FTE_REPORT_HEADER_START_ROWNUM
 lStartRow_FTE_LineItems = FTE_REPORT_LINEITEM_START_ROWNUM
 
 
'iRange_RowNumber needs to start at value = 1, So to increment from 2 through 5
 iRange_RowNumber = 1
    
    
'Run report on Model Tab 1 through Model tab 4
 For i = 1 To 4
        
        
        'Increment for each row starting on row 2 in the vba_NmRng_PDL_WorksheetName on the Pull Down List tab
         iRange_RowNumber = iRange_RowNumber + 1
         
        '----------------------------------------------------------------------------------------------------
        'Get header information on the current Model Tab's FTE line items being retrieved
        '   1) Model Tab Name
        '   2) Worksheet file name for each Model tab
        '   3) Tower Name for each model tab
        
         sWorksheet_Tab = rngWorksheetName.Cells(iRange_RowNumber, 1).Value
         sModelFileName = rngFileName.Cells(iRange_RowNumber, 1).Value
         sTowerName = rngTower.Cells(iRange_RowNumber, 1).Value
         
         
        '----------------------------------------------------------
        'Copy Subtowers individually with in the worksheet tab
        
         For iSubtowerNumber = 1 To lCOUNT_OF_SUBTOWERS
       
       
               '*************************************************************************************************
               '     Task 1 & 2: SQL query the CMI Pull report in the Model tabs filtering on column BM
               '
               '     NOTE: subroutine abReport_SQLquery_ModelWorksheet
               '
               '     1)Filter just the FTE line items (where column BM=1 or 18) on the Model sheet then
               '     2)Calculate the escalation of those lineitems
               '
               '     3)Create Public Recordset:  rsPUBLIC_Report_FTE_LineItems
                              
                              
                 If abReport_SQLquery_ModelWorksheet(xlWrkBk_CORP, _
                                                            sWorksheet_Tab, _
                                                            iSubtowerNumber, _
                                                            sModelFileName, _
                                                            sTowerName) = False Then
            
                       'Exit worksheet
                        Debug.Print "*** Error in subReport_CopyRangeXML_ToRecordset Sub ***"
                        GoTo ProcExit
                    
                 Else
                
                        'Debug.Print "Sub ran successfully for " & sWorksheet_Tab
             
                 End If
                 

                'If there are 0 FTE line items then skip the next 2 steps
                 If Pub_RowCount_Report_LineItems > 0 Then
                 
                 
                    '****************************************************************
                    '   Format the header for FTE LineItems
                     Report_Format_TowerHeaders xlWrkBk_CORP, _
                                                      lStartRow_FTE_LineItems - 1, _
                                                      FTE_REPORT_START_COLUMNNUM, _
                                                      lStartRow_FTE_LineItems + Pub_RowCount_Report_LineItems, _
                                                      Pub_RowCount_Report_LineItems, _
                                                      Pub_ColumnCount_Report_LineItems + lFIRST_COLUMN_REFERENCE_NUMBER, _
                                                      "Black"


                    '****************************************************************
                    '   Copy Recordset to Spreadsheet
                     If SUB_CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, "FTE Report", _
                                                    rsPUBLIC_Report_FTE_LineItems, _
                                                    lStartRow_FTE_LineItems - 1, _
                                                    FTE_REPORT_START_COLUMNNUM, _
                                                    True) = False Then
            
                        Debug.Print "**** Error Exited Copy FTE Report Recordset Sub ******"
                        GoTo ProcExit
            
                      End If
                      
            
                     'Debug.Print "Model 1 Pub_RowCount_Report_LineItems = " & Pub_RowCount_Report_LineItems
                     'Debug.Print "lStartRow_FTE_LineItems = " & lStartRow_FTE_LineItems
                     
                     
                '**************************************************************************************************************************
                '     Task 3: Write formulas into the FTE report
                '
                '     NOTE: The acReport_Run_WriteFormulas subroutine call 4 other subroutines with in it
                '
                '     1) Report_WriteFormula_Header_PxQ_Calculations
                '     2) Report_Populate_Header_AdjustmentValues
                '     3) Report_WriteFormula_LaborLoadFactor_PxQ_Calculations
                '     4) Report_WriteFormula_LaborLoadFactor_AdjustedbyYear
                '
            
                      acReport_Run_WriteFormulas xlWrkBk_CORP, _
                                              sWorksheet_Tab, _
                                              Pub_RowCount_Report_LineItems, _
                                              lStartRow_FTE_Header, _
                                              lStartRow_FTE_LineItems, _
                                              Pub_ColumnCount_Report_LineItems + lFIRST_COLUMN_REFERENCE_NUMBER, _
                                              Pub_ColumnCount_Report_LineItems - 36
                                              
            
                    'Add Row count
                     lStartRow_FTE_Header = lStartRow_FTE_Header + Pub_RowCount_Report_LineItems + lHEADER_ROWS
                     lStartRow_FTE_LineItems = lStartRow_FTE_LineItems + Pub_RowCount_Report_LineItems + lHEADER_ROWS
                     
                     
                End If
                
        Next
    
   Next
  
    
ProcExit:


    rsPUBLIC_Report_FTE_LineItems.Close
    Set rsPUBLIC_Report_FTE_LineItems = Nothing
    
    
Exit Sub

ProcErr:

  Select Case Err.Number
    
   Case 5
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
  
  Case 9
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
   ' Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next


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
   

End Sub