Public Sub Report_WriteFormula_LaborLoadFactor_AdjustedbyYear(rngFirstCell As Range, rngLastCell As Range, Optional sColorCell As String)

  '*************************************************************************************************
  '
  '   NOTE:This will adjust formula for PreGoLive monthly adjustment for months 1 to 24
  '
       
    Dim xlWrkSht_FTE_Report As Excel.Worksheet
    
       
   'Local Variables
    Dim lNumberOfRows As Long
    Dim lNumberOfColumns As Long
    Dim lSubTotal_StartRow As Long
    
    Dim sFormula_Adjusted1 As String
    Dim sFormula_Adjusted2 As String
    Dim sFormula_Adjusted3 As String
    Dim sFormula_Adjusted_Year03 As String
    Dim sFormula_Adjusted As String
    
    Dim sMonthDate_Header_Address As String
    Dim sSum_1_24_Month_Address As String
    Dim sIndex_Month1_Year11_Address As String
    
   'All PxQ Month 1 to Year 11 range
    Dim rngAdjusted As Range
    
  'Year 1 and 2 range
    Dim rngAdjusted_Year01 As Range
    Dim rngAdjusted_Year02 As Range
    
  'Year 3 with Adjustment calculations
    Dim rngAdjusted_Year03_Esc1 As Range
    Dim rngAdjusted_Year03_Esc2 As Range
    Dim rngAdjusted_Year03_Esc3 As Range
    Dim rngAdjusted_Year03 As Range
    Dim rngAdjusted_Year04_Year11 As Range
    
  'Subtotal and header
    Dim rngHeader As Range
    Dim rngSubTotals As Range
    Dim rngModel_TotalLaborExpense As Range
    Dim rngDifference_Dollars As Range
    Dim rngDifference_Percent As Range
    
    Dim rngTemp As Range
    
   'Count of Number of Rows and Columns
    lNumberOfRows = rngLastCell.Row + 1 - rngFirstCell.Row
    lNumberOfColumns = rngLastCell.Column - rngFirstCell.Column
    
   'Set worksheet objecct
    Set xlWrkSht_FTE_Report = rngFirstCell.Parent

    
   '---------------------------------------------------------------------------------------------------
   '    Set ranges for Adjusted calculation
   '
   
   'Set range Line Items rows
    Set rngAdjusted = Range(rngFirstCell, rngLastCell)
    
   'Set Header
    Set rngHeader = Range(Cells(rngFirstCell.Row - 1, rngFirstCell.Column), Cells(rngFirstCell.Row - 1, rngLastCell.Column - 1))
    
   'Set Sutotal/Footer and Model Total Expense
    Set rngSubTotals = Range(Cells(rngLastCell.Row, rngFirstCell.Column), Cells(rngLastCell.Row, rngLastCell.Column - 1))   'Set range last row of line items
    Set rngModel_TotalLaborExpense = Range(Cells(rngLastCell.Row + 1, rngFirstCell.Column), Cells(rngLastCell.Row + 1, rngLastCell.Column - 1)) 'Set range 1 Row AFTER last row of line items
    Set rngDifference_Dollars = Range(Cells(rngLastCell.Row + 2, rngFirstCell.Column), Cells(rngLastCell.Row + 2, rngLastCell.Column - 1)) 'Set range 2 Rows AFTER last row of line items
    Set rngDifference_Percent = Range(Cells(rngLastCell.Row + 3, rngFirstCell.Column), Cells(rngLastCell.Row + 3, rngLastCell.Column - 1)) 'Set range 3 Rows AFTER last row of line items
    
   'Set range for years 1 through 11
    Set rngAdjusted_Year01 = rngAdjusted.Range(Cells(1, 1), Cells(lNumberOfRows, 1))
    Set rngAdjusted_Year02 = rngAdjusted.Range(Cells(1, 2), Cells(lNumberOfRows, 2))
    Set rngAdjusted_Year03 = rngAdjusted.Range(Cells(1, 3), Cells(lNumberOfRows, 3))
    Set rngAdjusted_Year04_Year11 = rngAdjusted.Range(Cells(1, 4), Cells(lNumberOfRows, 11))
    
    
   '---------------------------------------------------------------------------------------------------
   '     Write Formulas Year Adjustment for year 1 and year 2

   'Year 1
    rngAdjusted_Year01.FormulaR1C1 = "=SUM(RC[-35]:RC[-24])"
    rngAdjusted_Year01.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    
   'Year 2
    rngAdjusted_Year02.FormulaR1C1 = "=SUM(RC[-23]:RC[-12])"
    rngAdjusted_Year02.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    
    
   'Set range for Adjusted formulas
    sMonthDate_Header_Address = "R5C64:R5C88"
    sSum_1_24_Month_Address = "RC64:RC88"
    sIndex_Month1_Year11_Address = "RC64:RC98"

   
   '---------------------------------------------------------------------------------------------------
   '   Write Formulas Year Adjustment for year 3 through year 11

   '1st Adjustment formula
    sFormula_Adjusted1 = "=SUMIFS(" & sSum_1_24_Month_Address & "," & sMonthDate_Header_Address & ","">=""&R5C," & sMonthDate_Header_Address & ",""<""&R5C[1])"
   
   '2nd Adjustment formula
    sFormula_Adjusted2 = "+IFERROR(INDEX(" & sIndex_Month1_Year11_Address & ",1,R7C-13)*R8C/R9C,0)"
    
   '3rd Adjustment formula
    sFormula_Adjusted3 = "+IFERROR(INDEX(" & sIndex_Month1_Year11_Address & ",1,R10C-13)*R11C/R12C,0)"
    
   '1st + 2nd +3rd adjustment formulas combined
   'Debug.Print "sFormula_Adjusted " & sFormula_Adjusted
    sFormula_Adjusted = sFormula_Adjusted1 & sFormula_Adjusted2 & sFormula_Adjusted3
    
   
   'Write Year 3 Adjusted to FTE Report
    rngAdjusted_Year03.FormulaR1C1 = sFormula_Adjusted
    rngAdjusted_Year03.NumberFormat = NUMBER_FORMAT_CUSTOM_ACCOUNTING
    
   'Write Year 4 to Year 11 Adjusted to FTE Report
    rngAdjusted_Year04_Year11.FormulaR1C1 = sFormula_Adjusted
    rngAdjusted_Year04_Year11.NumberFormat = NUMBER_FORMAT_CUSTOM_ACCOUNTING
    
   'Format line items - cells fill color
    ColorCell rngAdjusted_Year03, "None"
    ColorCell rngAdjusted_Year04_Year11, "None"
    
    
   '----------------------------------------------------
   '   Format HEADER
    For i = 1 To 11
    
        rngHeader.Cells(1, i).Value = "Year " & i
    
    Next i
    
   'Format HEADER
    rngHeader.HorizontalAlignment = xlCenter
    rngHeader.Font.Bold = True


    ColorCell rngHeader, "black", True
    
    
   '-----------------------------------------------------------------------------------------------------------------------
   '    Calculated Labor Expense:
   '
   '    a)
   '    b) Format SUBTOTALS and write subtotal formuals
   
    lSubTotal_StartRow = lNumberOfRows - 1
    
   'Subtotals write formulas and format them
    rngSubTotals.FormulaR1C1 = "=SUM(R[-" & lSubTotal_StartRow & "]C:R[-1]C)"
    rngSubTotals.NumberFormat = NUMBER_FORMAT_CUSTOM_ACCOUNTING
    
   'Format Borders
    With rngSubTotals.Borders(xlEdgeTop)
    
        .LineStyle = xlDouble
        .Weight = xlThick
        
    End With
    
   'Format Subtotals - bold and will the cells blue
    rngSubTotals.Font.Bold = True
    ColorCell rngSubTotals, "lightblue"
    
    
  '-------------------------------------------------------------------------------------------------------------------------
  '   Reconciliation Calculations:
  '
  '   1)  Total Labor Expense: that comes from the model (see columns Y through BB) This IS NOT the calculated labor expense
  '   2)  The Dollar difference between Calculated Labor Expense - Total Labor Expense
  '   3)  The Percent difference between Calculated Labor Expense - Total Labor Expense
  '   4)  Total Difference for the model
  
  
  '1 Total Labor Expense: that comes from the model (see columns Y through BB) This IS NOT the calculated labor expense
     rngModel_TotalLaborExpense.Cells(1, 1).FormulaR1C1 = "=R[-1]C[-62]"
     Range(rngModel_TotalLaborExpense.Cells(1, 2), rngModel_TotalLaborExpense.Cells(1, 11)).FormulaR1C1 = "=R[-1]C[-50]"
     rngModel_TotalLaborExpense.NumberFormat = NUMBER_FORMAT_CUSTOM_ACCOUNTING
    
    
  '2 The Dollar difference between Calculated Labor Expense - Total Labor Expense
     rngDifference_Dollars.FormulaR1C1 = "=R[-2]C-R[-1]C"
     rngDifference_Dollars.NumberFormat = NUMBER_FORMAT_CUSTOM_ACCOUNTING
  
   'Format Borders
     With rngDifference_Dollars.Borders(xlEdgeTop)
    
        .LineStyle = xlDouble
        .Weight = xlThin
        
     End With


  '3  The Percent difference between Calculated Labor Expense - Total Labor Expense
      'rngDifference_Percent.FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C/R[-2]C)"
      rngDifference_Percent.FormulaR1C1 = "=IFERROR(R[-1]C/R[-2]C,0)"
      rngDifference_Percent.NumberFormat = NUMBER_FORMAT_CUSTOM_PERCENT
      
      
  '4  Total Difference for the modelS

      'Model Total - Adjust labor loaded load factor
       Set rngTemp = xlWrkSht_FTE_Report.Cells(rngSubTotals.Row, rngLastCell.Column)
      
'       Debug.Print rngTemp.Address
'       Debug.Print "=Sum(R" & rngSubTotals.Row & "C" & rngFirstCell.Column & ":" & rngSubTotals.Row & "C" & rngLastCell.Column - 1 & ")"
       rngTemp.FormulaR1C1 = "=Sum(R" & rngSubTotals.Row & "C" & rngFirstCell.Column & ":R" & rngSubTotals.Row & "C" & rngLastCell.Column - 1 & ")"
       rngTemp.Font.Bold = True
   
      'Model Total - Total Labor Expense
       Set rngTemp = xlWrkSht_FTE_Report.Cells(rngModel_TotalLaborExpense.Row, rngLastCell.Column)
       rngTemp.FormulaR1C1 = "=Sum(R" & rngModel_TotalLaborExpense.Row & "C" & rngFirstCell.Column & ":R" & rngModel_TotalLaborExpense.Row & "C" & rngLastCell.Column - 1 & ")"
   
      'Model Total - Difference betwwen Adjust labor loaded load factor - Total Labor Expense
       Set rngTemp = xlWrkSht_FTE_Report.Cells(rngDifference_Dollars.Row, rngLastCell.Column)
       
       'Debug.Print rngTemp.Address
       rngTemp.FormulaR1C1 = "=Sum(R" & rngDifference_Dollars.Row & "C" & rngFirstCell.Column & ":R" & rngDifference_Dollars.Row & "C" & rngLastCell.Column - 1 & ")"

       With rngTemp.Borders(xlEdgeTop)
    
         .LineStyle = xlDouble
         .Weight = xlThin
        
       End With

  
      'Model Total - Percent difference between Adjust labor loaded load factor - Total Labor Expense
       Set rngTemp = xlWrkSht_FTE_Report.Cells(rngDifference_Percent.Row, rngLastCell.Column)
       rngTemp.FormulaR1C1 = "=IFERROR(R" & rngDifference_Percent.Row - 1 & "C" & rngLastCell.Column & "/R" & rngDifference_Percent.Row - 2 & "C" & rngLastCell.Column & ",0)"
       rngTemp.NumberFormat = NUMBER_FORMAT_CUSTOM_PERCENT


End Sub
