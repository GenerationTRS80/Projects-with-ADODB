Private Sub SUB_PivotSQL_Spreadsheet_Formatting(xlWrkSht_TARGET As Excel.Worksheet, _
                                                sTarget_WorksheetName As String, _
                                                lStartRow As Long, _
                                                lStartColumn As Long, _
                                                lNumberFields As Long)


 'Local Variables
  Dim lNumberOfRows As Long
  Dim lNumberOfColumns As Long
  
  
 'Excel Range object range
  Dim rngFirstCell As Range
  Dim rngLastCell As Range
  Dim rngHeader As Range

  
'--------------------------------------------------------------------------------------------
'   NOTE:   You need to check to see if the worksheet is visisible. If NOT visible then
'           formatting and setting of the cursor will *Fail* and cause an error
'


'Format sheet
 xlWrkSht_TARGET.Activate
 xlWrkSht_TARGET.Cells.Select
 
Select Case sTarget_WorksheetName


Case "Pivot_WeeklyHours"

'  With Selection
'      .HorizontalAlignment = xlGeneral
'      .VerticalAlignment = xlBottom
'      .WrapText = False
'      .Orientation = 0
'      .AddIndent = True
'      .IndentLevel = 0
'      .ShrinkToFit = False
'      .ReadingOrder = xlContext
'      .MergeCells = False
'
'  End With


Case "Data Forecast"

  With Selection
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = True
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
      
  End With


 Case "Issue and Risks"



 End Select


End Sub