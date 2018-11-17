Private Sub Sub_Test003()

'Excel Objects
  Dim xlWrkSht_FTE_Report As Excel.Worksheet
  Dim rngFirstCells As Range
  Dim rngLastCells As Range
  
  Dim lStartRow As Long
  Dim lStartColumn As Long
  Dim lEndRow As Long
  Dim lEndColumn As Long
  Dim lMonthsPreGoLive As Long
  
 'Set public variables
  Pub_RecordsetCount_FTE_LineItems = 10
  Pub_lMonthsPreGoLive = 4
  
  lStartRow = 16
  lStartColumn = 14
  lEndRow = lStartRow + Pub_RecordsetCount_FTE_LineItems
  lEndColumn = 96

  Set xlWrkSht_FTE_Report = ActiveWorkbook.Worksheets("FTE Report")
  
  Set rngFirstCells = xlWrkSht_FTE_Report.Cells(lStartRow, lStartColumn)
  Set rngLastCells = xlWrkSht_FTE_Report.Cells(lEndRow, lEndColumn)
  
  Report_WriteFormula_PxQ_LaborLoadFactor rngFirstCells, rngLastCells, Pub_lMonthsPreGoLive, "lightyellow"
  
  

End Sub
'
'*************************************************** SUBROUTINE BELOW / TEST DATA ABOVE *********************************************************
'

Public Sub Report_WriteFormula_PxQ_LaborLoadFactor(rngFirstCell As Range, rngLastCell As Range, Optional lMonthsPreGoLive As Long, Optional sColorCell As String)


   '*************************************************************************************************
   '
   '   NOTE:This will adjust formula for PreGoLive monthly adjustment for months 1 to 24
   '
       
   'Local Variables
    Dim lNumberOfRows As Long
    Dim lNumberOfColumns As Long
    Dim sFormula_PxQ_Year1_Months As String
    Dim sFormula_PxQ_Year2_Months As String
    
   'All PxQ Month 1 to Year 11 range
    Dim rngPxQ As Range
    
  'Year 1 range
    Dim rngPxQ_Months_1_12 As Range
    Dim rngPxQ_Year1 As Range
    
  'Adjustment Months
    Dim rngPxQ_Months_13_24_Year1_Adjustment As Range
    Dim rngPxQ_Months_13_24_Year2_Adjustment As Range
    
   'Year 2 through 11
    Dim rngPxQ_Year2 As Range
    Dim rngPxQ_Year3_Year11 As Range
    
    Dim rngSubTotals As Range
    

  'Row and Column Constant and variable value
  'Note: The determining factor in creating a constant is based on the formula that will be used in the range denoted by the column
    Const COLUMN_PxQ_Month_01 = 49
    Const COLUMN_PxQ_Year_01 = COLUMN_PxQ_Month_01 + 12
    Const COLUMN_PxQ_Month_13 = COLUMN_PxQ_Year_01 + 1
    Const COLUMN_PxQ_Year_02 = COLUMN_PxQ_Month_01 + 25
    Const COLUMN_PxQ_Year_03 = COLUMN_PxQ_Year_02 + 1
    Const COLUMN_PxQ_Year_11 = COLUMN_PxQ_Year_03 + 8
    
    lNumberOfRows = rngLastCell.Row + 1 - rngFirstCell.Row
    lNumberOfColumns = rngLastCell.Column - rngFirstCell.Column
    Debug.Print "Pre Go-Live Months " & lMonthsPreGoLive
    
   'Set range the Price x Quantity range
    Set rngPxQ = Range(rngFirstCell, rngLastCell)
    Set rngSubTotals = Range(Cells(rngLastCell.Row + 2, rngFirstCell.Column + COLUMN_PxQ_Month_01 - 1), Cells(rngLastCell.Row + 2, rngLastCell.Column))
    
   'Set range for first 12 Months and Year 1 total
    Set rngPxQ_Months_1_12 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Month_01), Cells(lNumberOfRows, COLUMN_PxQ_Month_01 + 11))
    Set rngPxQ_Year1 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Year_01), Cells(lNumberOfRows, COLUMN_PxQ_Year_01))

   'Set range Adjustment range for Months 13 through 24
    Set rngPxQ_Months_13_24_Year1_Adjustment = rngPxQ.Range(Cells(1, COLUMN_PxQ_Month_13), Cells(lNumberOfRows, COLUMN_PxQ_Month_13 + lMonthsPreGoLive - 1))
    Set rngPxQ_Months_13_24_Year2_Adjustment = rngPxQ.Range(Cells(1, COLUMN_PxQ_Month_13 + lMonthsPreGoLive), Cells(lNumberOfRows, COLUMN_PxQ_Month_13 + 11))
    
   'Set range for Year 1 total and Year 3 through 11 formula
    Set rngPxQ_Year2 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Year_02), Cells(lNumberOfRows, COLUMN_PxQ_Year_02))
    Set rngPxQ_Year3_Year11 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Year_03), Cells(lNumberOfRows, COLUMN_PxQ_Year_03 + 8))
    
    
   'Write formulas to range

    rngPxQ_Months_1_12.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC14*RC[-37])/12)" '"=IF(Y16="","",($N16*Y16)/12)"
    rngPxQ_Year1.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"

   'If Pre Go-Live months is 0 then set months 13 through 24 to Year2 Loaded Labor Cost
    If lMonthsPreGoLive = 0 Then
    
       'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year2_Adjustment.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC15*RC[-37])/12)" '=IF(AP16="","",($O16*AP16)/12)
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "iceblue"
    
    End If
    
    
    If lMonthsPreGoLive > 0 And lMonthsPreGoLive < 12 Then
    
      'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year1_Adjustment.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC14*RC[-37])/12)" '=IF(Y16="","",($N16*Y16)/12)
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "iceblue"
        
      'Year 2 Loaded Rate
        rngPxQ_Months_13_24_Year2_Adjustment.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC15*RC[-37])/12)" '=IF(AP16="","",($O16*AP16)/12)
        ColorCell rngPxQ_Months_13_24_Year2_Adjustment, "blue"
    
    End If
    
   'If Pre Go-Live months is 12 then set months 13 through 24 to Year 1 Loaded Labor Cost
    If lMonthsPreGoLive = 12 Then
    
  
       'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year1_Adjustment.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC14*RC[-37])/12)" '=IF(Y16="","",($N16*Y16)/12)
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "red"
    
    End If

    rngPxQ_Year2.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    rngPxQ_Year3_Year11.FormulaR1C1 = "=IF(RC[-37]="""","""",(RC[-72]*RC[-37])/12)"

    rngSubTotals.FormulaR1C1 = "=SUM(R[-12]C:R[-2]C)"

'    ColorCell rngPxQ_Months_13_24_Year1_Adjustment, sColorCell
'    ColorCell rngPxQ_Months_13_24_Year2_Adjustment, sColorCell
'     ColorCell rngFormula_Total_Year2, sColorCell

'    ColorCell rngSubTotals, sColorCell
'    ColorCell rngPxQ_Months_1_12, sColorCell
    ColorCell rngPxQ_Year1, "grey"
    ColorCell rngPxQ_Year2, "orange"
    ColorCell rngPxQ_Year3_Year11, "lightgreen"
        

  
End Sub