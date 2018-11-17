Private Sub acReport_Run_WriteFormulas(xlWrkBk_CORP As Excel.Workbook, sModelTab_WorksheetName As String, _
                                          lCountofRows_FTE_LineItems As Long, _
                                          lStartRow_Header As Long, _
                                          lStartRow_WriteFormula_PxQ As Long, _
                                          lStartColumn_WriteFormula_PxQ As Long, _
                                          lStartColumn_FTE_LineItems As Long)

  
 'Excel Object
  Dim xlWrkSht_FTE_Report As Excel.Worksheet
  
  Dim rngFirstCell As Range
  Dim rngLastCell As Range
  Dim rngPxQ As Range
  Dim rngHeader As Range
  
'Local Variable
  Dim lStartRow As Long
  Dim lStartColumn_WriteFormula_Adjustment As Long
  Dim lEndRow As Long
  Dim lEndColumn_WriteFormula_PxQ As Long
  
  Dim lNumberOfMonths_PreGoLive As Long
  Dim dContractSignDate_Year03 As Date
  Dim bCalculationManual As Boolean
  
  
 'Set FTE Report Worksheet variable
  Set xlWrkSht_FTE_Report = xlWrkBk_CORP.Worksheets("FTE Report")
  
 'Get following : 1) Number of Pre Go-Live months 2) Year 3 Contract sign date
  lNumberOfMonths_PreGoLive = CLng(xlWrkBk_CORP.Worksheets(sModelTab_WorksheetName).Cells(5, 22).Value)
  dContractSignDate_Year03 = CDate(xlWrkBk_CORP.Worksheets(sModelTab_WorksheetName).Cells(5, 48).Value)
  
 'Set Number of Pre Go-Live months
  Pub_dYear03_ContractSign = dContractSignDate_Year03
  Pub_lMonthsPreGoLive = lNumberOfMonths_PreGoLive
  
 'Debug.Print "Number of Months of Pre Go Live " & Pub_lMonthsPreGoLive
 'Debug.Print "Column Count " & lStartColumn_WriteFormula_PxQ
 

 '----------------------------------------------------------------------------------------------------------------
 '  Create Header and Year 3 through 11 adjustment matric
 '
 '  1) Run Year 3 to Year 11 Adjustment Header Report
 '  2) Run Create Header: Month 1 to 24 Header and Year 1 to Year 11
 
   '1) Write PxQ formula
  
    'Row Calculaltion the last row of the range from the first row of FTE line item plus the number of rows returned from the recorset minus 1
     lEndColumn_WriteFormula_PxQ = lStartColumn_WriteFormula_PxQ + 24

    'Set the first cell of the range (bottom left) and last cell (top right cell)
     Set rngFirstCell = xlWrkSht_FTE_Report.Cells(lStartRow_Header, lStartColumn_WriteFormula_PxQ)
     Set rngLastCell = xlWrkSht_FTE_Report.Cells(lStartRow_Header, lEndColumn_WriteFormula_PxQ)
 
    '---------------------------------------------------------------------------------------------------
    '  Create Year and Monthly header
    '  Calculate: Month 1 = Contract sign date - 24 Month - Number of Months of pre go-live
    
       Report_WriteFormula_Header_PxQ_Calculations rngFirstCell, rngLastCell, sModelTab_WorksheetName
     
        
    '----------------------------------------------------------------------------------------------------------------------
    '  Populate Header Matrix
    
       Report_Populate_Header_AdjustmentValues xlWrkBk_CORP, sModelTab_WorksheetName, , lStartRow_Header
       
       
    '----------------------------------------------------------------------------------------------------------------------
    '  >> Hide headers by using Group function <<
      
       xlWrkSht_FTE_Report.Rows("" & lStartRow_Header - 2 & ":" & lStartRow_Header + 8 & "").Group


 '-------------------------------------------------------------------------------------------------------------------------------------------------
 '  Adjusted Totals
 '
 '  1) Write PxQ formula that adjust or shifts Year 1 PxQ to the right by the number of Pre Go-Live Months
 '  2) Write Yearly Adjustment formula to adjusts the PxQ year total for Year 3 through 11 for Pre Go-Live adjust months
 

   '1) Write PxQ formula
 

    'Row Calculaltion the last row of the range from the first row of FTE line item plus the number of rows returned from the recorset minus 1
    'Debug.Print " lCountofRows_FTE_LineItems = " & lCountofRows_FTE_LineItems
    
     lEndRow = lStartRow_WriteFormula_PxQ + lCountofRows_FTE_LineItems - 1
     lEndColumn_WriteFormula_PxQ = lStartColumn_WriteFormula_PxQ + 24

    'Set the first cell of the range (bottom left) and last cell (top right cell)
     Set rngFirstCell = xlWrkSht_FTE_Report.Cells(lStartRow_WriteFormula_PxQ, lStartColumn_WriteFormula_PxQ)
     Set rngLastCell = xlWrkSht_FTE_Report.Cells(lEndRow, lEndColumn_WriteFormula_PxQ)
     
     
    '------------------------------------------------------------------------------------------------------------
    '   >>>   Write formulas   <<<
    
     Report_WriteFormula_LaborLoadFactor_PxQ_Calculations rngFirstCell, rngLastCell, lStartColumn_FTE_LineItems, lNumberOfMonths_PreGoLive, "lightgrey"
     
     
    '2) Write Yearly Adjustment
     Const lPxQ_WRITEFORMULA_YEAR02_YEAR11 = 11
     
    
    'Row Calculaltion the last row of the range from the first row of FTE line item plus the number of rows returned from the recorset minus 1
     lStartColumn_WriteFormula_Adjustment = lEndColumn_WriteFormula_PxQ + lPxQ_WRITEFORMULA_YEAR02_YEAR11
     lEndRow = lStartRow_WriteFormula_PxQ + lCountofRows_FTE_LineItems - 1
     lEndColumn = lStartColumn_WriteFormula_Adjustment + 11
     
    'Debug.Print vbCrLf & "lPxQ_WRITEFORMULA_YEAR02_YEAR11 = " & lPxQ_WRITEFORMULA_YEAR02_YEAR11
    'Debug.Print "lEndColumn_WriteFormula_PxQ = " & lEndColumn_WriteFormula_PxQ
    'Debug.Print "lStartColumn_WriteFormula_Adjustment = " & lStartColumn_WriteFormula_Adjustment
    'Debug.Print "lEndColumn = " & lEndColumn & vbCrLf

    'Set the first cell of the range (bottom left) and last cell (top right cell)
     Set rngFirstCell = xlWrkSht_FTE_Report.Cells(lStartRow_WriteFormula_PxQ, lStartColumn_WriteFormula_Adjustment)
     Set rngLastCell = xlWrkSht_FTE_Report.Cells(lEndRow, lEndColumn)
  
    '------------------------------------------------------------------------------------------------------------------------
    ' Write formula
    
     Report_WriteFormula_LaborLoadFactor_AdjustedbyYear rngFirstCell, rngLastCell, "blue"


    'Turn Automatic calculation back on
     If bCalculationManual = False Then
 
            Application.Calculation = xlCalculationAutomatic
 
     End If
 
 
    'Recalculate FTE worksheet
     xlWrkSht_FTE_Report.Calculate


End Sub
