Public Sub Report_WriteFormula_LaborLoadFactor_PxQ_Calculations(rngFirstCell As Range, rngLastCell As Range, lStartColumn_FTE_LineItems, Optional lMonthsPreGoLive As Long, Optional sColorCell As String)

   '*************************************************************************************************
   '
   '   NOTE:This will adjust formula for PreGoLive monthly adjustment for months 1 to 24
   '
       
   'Local Variables
    Dim lNumberOfRows As Long
    Dim lNumberOfColumns As Long
    Dim lSubTotal_StartRow As Long
    Dim lStartColumn_PxQ_LineItems As Long
    Dim lColumnDifference_PxQ_FTE_Months_1_24 As Long
    Dim lColumnDifference_PxQ_FTE_Year3_Year11 As Long

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
    Dim rngHeader As Range
    

  'Row and Column Constant and variable value
  'Note: The determining factor in creating a constant is based on the formula that will be used in the range denoted by the column
    Const COLUMN_PxQ_Year_02 = 14 + 12
    Const COLUMN_PxQ_Year_03 = COLUMN_PxQ_Year_02 + 1
    
    lNumberOfRows = rngLastCell.Row + 1 - rngFirstCell.Row
    lNumberOfColumns = rngLastCell.Column - rngFirstCell.Column
    lStartColumn_PxQ_LineItems = rngFirstCell.Column
    lColumnDifference_PxQ_FTE_Months_1_24 = lStartColumn_PxQ_LineItems - lStartColumn_FTE_LineItems
    lColumnDifference_PxQ_FTE_Year3_Year11 = lColumnDifference_PxQ_FTE_Months_1_24 + 35
    
'     Debug.Print "lColumnDifference_PxQ_FTE_Year3_Year11 = " & lColumnDifference_PxQ_FTE_Year3_Year11 & " = " & lColumnDifference_PxQ_FTE_Months_1_24 & " + " & 35
'     Debug.Print "lColumnDifference_PxQ_FTE_Months_1_24 = " & lColumnDifference_PxQ_FTE_Months_1_24 & " = " & lStartColumn_PxQ_LineItems & " - " & lStartColumn_FTE_LineItems
   ' Debug.Print "Pre Go-Live Months " & lMonthsPreGoLive
    
   'Set range the Price x Quantity range Header and Subtotals
    Set rngPxQ = Range(rngFirstCell, rngLastCell)
    
   'Set Range Header and Subtotals
    Set rngSubTotals = Range(Cells(rngLastCell.Row + 2, rngFirstCell.Column), Cells(rngLastCell.Row + 2, rngLastCell.Column + 10))
    Set rngHeader = Range(Cells(rngFirstCell.Row - 1, rngFirstCell.Column), Cells(rngFirstCell.Row - 1, rngLastCell.Column + 10))

    
   'Set range for first 12 Months and Year 1 total
    Set rngPxQ_Months_1_12 = rngPxQ.Range(Cells(1, 1), Cells(lNumberOfRows, 12))
    Set rngPxQ_Year1 = rngPxQ.Range(Cells(1, 13), Cells(lNumberOfRows, 13))

   'Set range Adjustment range for Months 13 through 24
    Set rngPxQ_Months_13_24_Year1_Adjustment = rngPxQ.Range(Cells(1, 14), Cells(lNumberOfRows, 14 + lMonthsPreGoLive - 1))
    Set rngPxQ_Months_13_24_Year2_Adjustment = rngPxQ.Range(Cells(1, 14 + lMonthsPreGoLive), Cells(lNumberOfRows, 14 + 11))
    
   'Set range for Year 1 total and Year 3 through 11 formula
    Set rngPxQ_Year2 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Year_02), Cells(lNumberOfRows, COLUMN_PxQ_Year_02))
    Set rngPxQ_Year3_Year11 = rngPxQ.Range(Cells(1, COLUMN_PxQ_Year_03), Cells(lNumberOfRows, COLUMN_PxQ_Year_03 + 8))
    
    
   '---------------------------------------------------------------------------------------------------
   '    Write formulas to range
   '
    rngPxQ_Months_1_12.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC14*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "])/12)" '"=IF(Y16="","",($N16*Y16)/12)"
    rngPxQ_Months_1_12.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    
    rngPxQ_Year1.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    rngPxQ_Year1.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    
    ColorCell rngPxQ_Months_1_12, "iceblue"
    ColorCell rngPxQ_Year1, "none"

   '---------------------------------------------------------------------------------------------------
   '
   '   Make Pre Go-Live Monthly Adjusstments (use formulaR1C1 to write formulas)
   '
   
   'If Pre Go-Live months is 0 then set months 13 through 24 to Year2 Loaded Labor Cost
    If lMonthsPreGoLive = 0 Then
    
       'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year2_Adjustment.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC15*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "])/12)" '=IF(AP16="","",($O16*AP16)/12)
        rngPxQ_Months_13_24_Year2_Adjustment.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "iceblue"
    
    End If
    
   'If pre go-live months are between 1 and 11 adjust the monthly calculation by PxQ Year 1 and Year 2
    If lMonthsPreGoLive > 0 And lMonthsPreGoLive < 12 Then
    
      'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year1_Adjustment.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC14*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "])/12)" '=IF(Y16="","",($N16*Y16)/12)
        rngPxQ_Months_13_24_Year1_Adjustment.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "iceblue"
        
      'Year 2 Loaded Rate
        rngPxQ_Months_13_24_Year2_Adjustment.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC15*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "])/12)" '=IF(AP16="","",($O16*AP16)/12)
        rngPxQ_Months_13_24_Year2_Adjustment.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        ColorCell rngPxQ_Months_13_24_Year2_Adjustment, "lightblue"
    
    End If
    
   'If Pre Go-Live months is 12 then set months 13 through 24 to Year 1 Loaded Labor Cost
    If lMonthsPreGoLive = 12 Then
    
  
       'Year 1 Loaded Rate
        rngPxQ_Months_13_24_Year1_Adjustment.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC14*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "])/12)" '=IF(Y16="","",($N16*Y16)/12)
        rngPxQ_Months_13_24_Year1_Adjustment.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        ColorCell rngPxQ_Months_13_24_Year1_Adjustment, "lightyellow"
    
    End If


    rngPxQ_Year2.FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
    rngPxQ_Year2.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    rngPxQ_Year3_Year11.FormulaR1C1 = "=IF(RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]="""","""",(RC[-" & lColumnDifference_PxQ_FTE_Year3_Year11 & "]*RC[-" & lColumnDifference_PxQ_FTE_Months_1_24 & "]))"
    rngPxQ_Year3_Year11.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

   'Write Header labels Month 1 through 14 and Year 1 through Year 11
    For i = 1 To 12
    
        rngHeader.Cells(1, i).Value = "Month " & i
        'Debug.Print rngHeader.Cells(1, i).Address
    
    Next i
    
    rngHeader.Cells(1, 13).Value = "Year 1"
    
    For i = 13 To 24
    
        rngHeader.Cells(1, i + 1).Value = "Month " & i
        'Debug.Print rngHeader.Cells(1, i + 1).Address
    
    Next i
    
    For i = 25 To 34
    
        rngHeader.Cells(1, i + 1).Value = "Year " & i - 23
        'Debug.Print rngHeader.Cells(1, i + 1).Address
    
    Next i
    
    rngHeader.HorizontalAlignment = xlCenter
        
   'Color fill and text rngheader
    ColorCell rngHeader, "black"
    ColorCell rngHeader, "white", True
     
    
   'SUBTOTALS
    lSubTotal_StartRow = lNumberOfRows + 1
       
    rngSubTotals.FormulaR1C1 = "=SUM(R[-" & lSubTotal_StartRow & "]C:R[-2]C)"
    rngSubTotals.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
  
    With rngSubTotals.Borders(xlEdgeTop)
    
        .LineStyle = xlDouble
        .Weight = xlThick
        
    End With
    
    rngSubTotals.Font.Bold = True
    ColorCell rngSubTotals, sColorCell


'******************************** Color Range ********************************

'    ColorCell rngPxQ_Months_13_24_Year1_Adjustment, sColorCell
'    ColorCell rngPxQ_Months_13_24_Year2_Adjustment, sColorCell
'    ColorCell rngFormula_Total_Year2, sColorCell

    
'    ColorCell rngPxQ, "red"


'    ColorCell rngPxQ_Year1, "grey"
'    ColorCell rngPxQ_Year2, "orange"
'    ColorCell rngPxQ_Year3_Year11, "lightgreen"
        
  
End Sub
