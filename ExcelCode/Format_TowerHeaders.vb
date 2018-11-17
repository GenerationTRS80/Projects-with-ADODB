Public Sub Report_Format_TowerHeaders(xlWrkBk_CORP As Workbook, lStartRow As Long, lStartColumn As Long, lEndRow As Long, Optional lNumberOfRows As Long, Optional lEndColumn As Long, Optional sColorCell As String)

  
  Dim xlWrkSht As Excel.Worksheet
  Dim rngFirstCell As Range
  Dim rngLastCell As Range
  Dim rngHeader As Range
  Dim rngHeader_CMIpull As Range
  Dim rngHeader_Price_Salary_Escalation As Range
  Dim rngHeader_Quantity_FTEs As Range
  Dim rngHeader_PxQ_Calculations As Range
  

  Set xlWrkSht = xlWrkBk_CORP.Worksheets("FTE Report")
  
  
 'Set Header range
  Set rngFirstCell = xlWrkSht.Cells(lStartRow, lStartColumn)
  Set rngLastCell = xlWrkSht.Cells(lStartRow, lEndColumn)
  
  Set rngHeader = Range(rngFirstCell, rngLastCell)
  
  
 'Set Header for those column that are copied from the CMI Tab header in the Model tabs header on row 4
  Set rngFirstCell = xlWrkSht.Cells(lStartRow, lStartColumn)
  Set rngLastCell = xlWrkSht.Cells(lStartRow, lStartColumn + 10)
  
  Set rngHeader_CMIpull = Range(rngFirstCell, rngLastCell)
  

 'Set Header for Salary Escalation
  Set rngFirstCell = xlWrkSht.Cells(lStartRow, lStartColumn + 11)
  Set rngLastCell = xlWrkSht.Cells(lStartRow, lStartColumn + 11 + 10)
  
  Set rngHeader_Price_Salary_Escalation = Range(rngFirstCell, rngLastCell)
  
  
 'Set Header for FTE Qty (from CMI Pull)
  Set rngFirstCell = xlWrkSht.Cells(lStartRow, lStartColumn + 11 + 11)
  Set rngLastCell = xlWrkSht.Cells(lStartRow, lStartColumn + 11 + 11 + 38)
  
  Set rngHeader_Quantity_FTEs = Range(rngFirstCell, rngLastCell)

  
 'Set Header for the calculated escalation of annual salary
  Set rngFirstCell = xlWrkSht.Cells(lStartRow, lEndColumn)
  Set rngLastCell = xlWrkSht.Cells(lStartRow, lEndColumn)
  
  
 'Set  the first cell of the Header PxQ calculations
  Set rngHeader_PxQ_Calculations = Range(rngFirstCell, rngLastCell)

 'Set cell colors

  ColorCell rngHeader_CMIpull, "black", True
  ColorCell rngHeader_Price_Salary_Escalation, "lightyellow"
  ColorCell rngHeader_Quantity_FTEs, "black", True
  ColorCell rngHeader_PxQ_Calculations, "blue"
  
 'Set formattting
'  rngHeader_Price_Salary_Escalation.WrapText = True

  


End Sub
