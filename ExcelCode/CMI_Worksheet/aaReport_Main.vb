Public Sub aaReport_Main()

   'Excel Variables
    Dim xlWrkBk_CORP As Excel.Workbook
    
   'Local variables
    Dim sWorksheet_Tab As String
    Dim lRowCount As Long

   On Error GoTo ProcErr
    
    Set xlWrkBk_CORP = Application.ActiveWorkbook
    

   'Clear FTE Report
    Report_ClearAll xlWrkBk_CORP, "FTE Report"
   
   'Set row counts to  = 0
    Pub_RecordsetCount_FTE_LineItems = 0
    lRowCount = TARGET_WORKSHEET_START_ROWNUM
    
 '-------------------------------------------------------------------------------
 '               >>>>> Copy Range To a Recordset <<<<<<<<
 '
   
    sWorksheet_Tab = "Model 1"
       
    If Report_CopyRangeXML_ToRecordset(xlWrkBk_CORP, sWorksheet_Tab) = False Then
 
       'Exit worksheet
         Debug.Print "*** Error in subReport_CopyRangeXML_ToRecordset Sub ***"
        'GoTo ProcExit
        
    Else
    
        Debug.Print "Sub ran successfully for " & sWorksheet_Tab
 
    End If
    

   '>>>>>>>>>>>>> Copy Recordset to FTE Reports <<<<<<<<<<<<<<<<
    If SUB_CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, "FTE Report", rsPUBLIC_FTEs_LineItems, _
                                      lRowCount, _
                                      TARGET_WORKSHEET_START_COLUMNNUM, False) = False Then

        Debug.Print "**** Error Exited CopyRecordset Sub ******"
        GoTo ProcExit

    End If
   
   'Add Row count
    lRowCount = lRowCount + Pub_RecordsetCount_FTE_LineItems + 3


 '-------------------------------------------------------------------------------
 '               >>>>> Copy Range To a Recordset <<<<<<<<
 '
   
    sWorksheet_Tab = "Model 2"
       
    If Report_CopyRangeXML_ToRecordset(xlWrkBk_CORP, sWorksheet_Tab) = False Then
 
       'Exit worksheet
         Debug.Print "*** Error in Report_CopyRangeXML_ToRecordset Sub ***"
        'GoTo ProcExit
        
    Else
    
        Debug.Print "Sub ran successfully for " & sWorksheet_Tab
 
    End If
    

   '>>>>>>>>>>>>> Copy Recordset to FTE Reports <<<<<<<<<<<<<<<<
    If SUB_CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, "FTE Report", rsPUBLIC_FTEs_LineItems, _
                                      lRowCount, _
                                      TARGET_WORKSHEET_START_COLUMNNUM, False) = False Then

        Debug.Print "**** Error Exited CopyRecordset Sub ******"
        GoTo ProcExit

    End If
   
   'Add Row count
    lRowCount = lRowCount + Pub_RecordsetCount_FTE_LineItems + 3
    
    
 '-------------------------------------------------------------------------------
 '               >>>>> Copy Range To a Recordset <<<<<<<<
 '
   
    sWorksheet_Tab = "Model 3"
       
    If Report_CopyRangeXML_ToRecordset(xlWrkBk_CORP, sWorksheet_Tab) = False Then
 
       'Exit worksheet
         Debug.Print "*** Error in Report_CopyRangeXML_ToRecordset Sub ***"
        'GoTo ProcExit
        
    Else
    
        Debug.Print "Sub ran successfully for " & sWorksheet_Tab
 
    End If
    

   '>>>>>>>>>>>>> Copy Recordset to FTE Reports <<<<<<<<<<<<<<<<
    If SUB_CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, "FTE Report", rsPUBLIC_FTEs_LineItems, _
                                      lRowCount, _
                                      TARGET_WORKSHEET_START_COLUMNNUM, False) = False Then

        Debug.Print "**** Error Exited CopyRecordset Sub ******"
        GoTo ProcExit

    End If
   
   'Add Row count
    lRowCount = lRowCount + Pub_RecordsetCount_FTE_LineItems + 3
    

 '-------------------------------------------------------------------------------
 '               >>>>> Copy Range To a Recordset <<<<<<<<
 '
   
    sWorksheet_Tab = "Model 4"
       
    If Report_CopyRangeXML_ToRecordset(xlWrkBk_CORP, sWorksheet_Tab) = False Then
 
       'Exit worksheet
         Debug.Print "*** Error in Report_CopyRangeXML_ToRecordset Sub ***"
        'GoTo ProcExit
        
    Else
    
        Debug.Print "Sub ran successfully for " & sWorksheet_Tab
 