Public Sub aaMainSub_CMIpull(xlWrkSht_Button As Excel.Worksheet, _
                             sTarget_WorksheetName As String, _
                             Optional sSelected_FilePathName = "NONE", _
                             Optional sInitial_FilePath As String)

 'Excel Objects
  Dim appXL As Excel.Application
  Dim xlWrkBk_CORP As Excel.Workbook
  Dim xlWrkBk_SP As Excel.Workbook
  Dim xlWrkBk_Add As Excel.Workbook
  
 'Local Variable
  'Dim sSelected_FilePathName As String
  Dim sSelected_FilePath As String
  Dim sFileName As String
  Dim sWorkbook_Added_Name As String
  
  Dim bModelNotFound_OnSharePoint As Boolean
  

 'ADO Objectsd
  Dim rsClone_CMIworksheet_LineItems As ADODB.Recordset
  Dim rsClone_Filtered_CMIworksheet_LineItems As ADODB.Recordset
  
 'Set Default
  bModelNotFound_OnSharePoint = False
  
  On Error GoTo ProcErr


 'Set workbook and then application
  Set xlWrkBk_CORP = xlWrkSht_Button.Parent
  Set appXL = xlWrkBk_CORP.Parent

 'Add a new workbook object
  With appXL

        .AskToUpdateLinks = False
        .Cursor = xlWait
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False

  End With
  

 '------------->>> This action takes the hyperlink and opens the Excel file on Sharepoint <<<---------------
 'Set appXL = New Excel.Application
  Set xlWrkBk_Add = appXL.Workbooks.Add
 
 'Open a Hyperlik
   xlWrkBk_Add.FollowHyperlink sSelected_FilePathName, , False
 ' Debug.Print "FollowYperlink " & sSelected_FilePathName


 '>>>>> If the file is not found and the user want to open on SP then run the abFileDialog_OpenWorkbook  <<<<<
  If bModelNotFound_OnSharePoint = True Then
  
      'Add Deal Document and Financial folder to initial file path
        sInitial_FilePath = sInitial_FilePath & "/Deal Documents/05. Financial/"
  
       If abFileDialog_OpenWorkbook(xlWrkSht_Button, sInitial_FilePath) = False Then
       
            Debug.Print "adFileDialog Failed"
            GoTo ProcExit
            
       End If
  
      'Pass the file path and model file name returned by the abFileDialog
      'via the Public variable Pub_sFilePathName to sSelected_FilePathName
       sSelected_FilePathName = Pub_sFilePathName
       
     'Open a Hyperlik with file path name provided by abFileDialog
       xlWrkBk_Add.FollowHyperlink sSelected_FilePathName, , False
       
  End If
  

 'Set the opened active workbook in Sharepoint to workbook object xlWrkBk_SP
  Set xlWrkBk_SP = appXL.ActiveWorkbook
 
 '------------------->>> Get the Selected FilePathName, FilePath and FileName <<<------------------------
 'sSelected_FilePathName = Trim(xlWrkBk_SP.FullName)
  sFileName = Trim(xlWrkBk_SP.Name)
  sSelected_FilePath = Trim(Left(sSelected_FilePathName, Len(sSelected_FilePathName) - Len(sFileName)))
 

 'Set to a Public Variable as FilePathName, FilePath and FileName
  Pub_sFilePath = sSelected_FilePath
  Pub_sFilePathName = sSelected_FilePathName
  Pub_sFileName = sFileName
 
 ' Debug.Print "New Sharepoint workbook name =  " & Pub_sFilePath & vbCrLf & "Name of File " & Pub_sFileName
 

 '------------->>>  Open DOM XML Document and copy it into EXCEL Workbook (XML to COM  object) <<<--------------
  If acOpenRecordset_from_ExcelWorksheet(xlWrkBk_SP, True) = False Then
    
       'Check to see if the worksheet is Cloud or AMS iF it is cloud or AMS then don't select the Start Page
        If sTarget_WorksheetName <> "Cloud - CMI" And sTarget_WorksheetName <> "AMS - CMI" Then
    
        'Select the Start Page
            With xlWrkSht_Button
         
                .Activate
                .Range("C3").Select
                 
            End With
         
        End If
    
       'Exit worksheet
        Debug.Print "*** Error in OpenRecordset_from_ExcelWorsheet Sub ***"
        GoTo ProcExit
 
 End If
 
 
'------------->>> POPULATE RECORDSET - CMIWorksheet Line Items<<<--------------
 If adPOPULATE_rsPUBLIC_CMIworksheet_LineItems(rsPUBLIC_CMIworksheet, PULL_WORKSHEET_START_ROWNUM) = False Then
  
   'Exit subroutine
    GoTo ProcExit
  
 End If
 
'---------------       >>> Copy all line items UNFILTERED data <<<     --------------
 Set rsClone_CMIworksheet_LineItems = rsPUBLIC_CMIworksheet_LineItems.Clone

 If CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, "Tab View - CMI", rsClone_CMIworksheet_LineItems, _
                                CMI_TABVIEW_START_ROWNUM, _
                                CMI_TABVIEW_START_COLUMNNUM, _
                                HEADER_YESNO, _
                                CMI_TABVIEW_ROWCOUNT, _
                                CMI_TABVIEW_COLUMNCOUNT, _
                                7, 7) = False Then

    Debug.Print "**** Error Exited CopyRecordset Sub ******"
    GoTo ProcExit

 End If
 
 
'-------------------------------------------------------------------------------------------------------------------------------------
'
'    Write Header data to Model Tab
'
 Set rsClone_CMIworksheet_LineItems = rsPUBLIC_CMIworksheet_LineItems.Clone

 If CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, sTarget_WorksheetName, rsClone_CMIworksheet_LineItems, _
                                CMI_TABVIEW_START_ROWNUM, _
                                CMI_TABVIEW_START_COLUMNNUM, _
                                HEADER_YESNO, _
                                CMI_TABVIEW_ROWCOUNT, _
                                CMI_TABVIEW_COLUMNCOUNT, _
                                7, 7) = False Then

    Debug.Print "**** Error Exited CopyRecordset Sub ******"
    GoTo ProcExit

 End If


 
'-------------------------------------------------------------------------------------------------------------------------------------
'
'    A) CREATE RECORDSET for
'               1) FTEs recordset
'               2) Expense recordset
'
'    B) COPY RECORDSETs FTE Line Items and Expense lines to the Model tab
'

    'Create Filtered Recordset on column BM criteria
      Set rsClone_CMIworksheet_LineItems = rsPUBLIC_CMIworksheet_LineItems.Clone
    
      If aeFILTER_CMIworksheet_LineItems(rsClone_CMIworksheet_LineItems) = False Then
      
       'Exit subroutine
        GoTo ProcExit
      
      End If
    
      Set rsClone_CMIworksheet_LineItems = rsPUBLIC_Filtered_CMIworksheet_LineItems.Clone

    '--------------------------------------------------------------------------------------
    '   Set the data types (integer, long, double etc) for each field
    '                   from recordsource
    '
    '   A) Create an individual recordset for: FTE line items
    '   B) Create an individual recordset for: Expense line items
    
     If afSET_DATATYPE_LineItems(rsClone_CMIworksheet_LineItems) = False Then
    
       'Exit subroutine
        GoTo ProcExit
    
     End If


    '----------------------------------------------------------------------------------------------------------
    '   1) CopyRecordset FTEs Line Items
    '
    
     Set rsClone_CMIworksheet_LineItems = rsPUBLIC_FTEs_LineItems.Clone
     
     If CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, sTarget_WorksheetName, rsClone_CMIworksheet_LineItems, _
                                    TARGET_WORKSHEET_START_ROWNUM, _
                                    TARGET_WORKSHEET_START_COLUMNNUM, _
                                    HEADER_YESNO, _
                                    TARGET_WORKSHEET_ROWCOUNT, _
                                    TARGET_WORKSHEET_COLUMNCOUNT, _
                                    7, 7) = False Then
    
        Debug.Print "**** Error Exited CopyRecordset Sub ******"
        GoTo ProcExit
    
     End If


    '----------------------------------------------------------------------------------------------------------
    '   2) CopyRecordset Expense Line Items
    '
    
     Set rsClone_CMIworksheet_LineItems = rsPUBLIC_Expense_LineItems.Clone
    
     If CopyRecordset_to_Spreadsheet(xlWrkBk_CORP, sTarget_WorksheetName, rsClone_CMIworksheet_LineItems, _
                                   TARGET_WORKSHEET_START_ROWNUM + Pub_RecordsetCount_FTE_LineItems, _
                                   TARGET_WORKSHEET_START_COLUMNNUM, _
                                   False, _
                                   TARGET_WORKSHEET_ROWCOUNT, _
                                   TARGET_WORKSHEET_COLUMNCOUNT) = False Then
    
       Debug.Print "**** Error Exited CopyRecordset Sub ******"
       GoTo ProcExit
    
    End If


'-------------------------------------------------------------------------------------------------------------------------------
'
'    Write filename to Pull Down List worksheet
'
     WriteFilename_PullDownLists_TabHeaders xlWrkBk_CORP, sTarget_WorksheetName, _
                                            Pub_DBUpload_FileName, _
                                            Pub_sFilePathName, _
                                            Pub_DBUpload_DealNumber, _
                                            Pub_DBUpload_TowerName, _
                                            Pub_DBUpload_TemplateNumber
     
     
    ''Get the workbook added name and Check the workbook added name to make sure it is correct
    ' sWorkbook_Added_Name = xlWrkBk_Add.Name
    ' Debug.Print "Workbook Added name = " & sWorkbook_Added_Name & vbCrLf & " Active WorkbookName " & appXL.ActiveWorkbook.Name


ProcExit:

 With appXL

    .AskToUpdateLinks = True
    .Cursor = xlDefault
    .EnableEvents = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True

 End With
 
'Close Added Workbook
 xlWrkBk_Add.Close False
 
 rsPUBLIC_CMIworksheet.Close
 Set rsPUBLIC_CMIworksheet = Nothing
  
 rsPUBLIC_CMIworksheet_LineItems.Close
 Set rsPUBLIC_CMIworksheet_LineItems = Nothing
 
 rsPUBLIC_Worksheet_Tab_DBUpload.Close
 Set rsPUBLIC_Worksheet_Tab_DBUpload = Nothing
 
 rsPUBLIC_Filtered_CMIworksheet_LineItems.Close
 Set rsPUBLIC_Filtered_CMIworksheet_LineItems = Nothing
 
 rsPUBLIC_FTEs_LineItems.Clone
 Set rsPUBLIC_FTEs_LineItems = Nothing
 
'Closer Recordset clone
 rsClone_CMIworksheet_LineItems.Close
 Set rsClone_CMIworksheet_LineItems = Nothing

 Debug.Print ">>>> Successfull Completed All Subroutines !!!! <<<<" & vbCrLf
 
 
 Exit Sub

ProcErr:

  Select Case Err.Number
   
  Case -2146697209
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next
    
  Case -2147467260 'Cancel button was hit oh Hyperlink destination selection
   Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
   Resume ProcExit

  Case -2146697201, -2146697210
  
   'Check the user to see if they want to select the file path to the cost model themselves on SharePoint site
    If MsgBox("Cost Model file was not found on in the Sharepoint folder where it was expected to be. (note the file could be in an Archive Folder" & vbCrLf & vbCrLf & _
             "Click yes to select another folder where the model might be located?", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
    
        bModelNotFound_OnSharePoint = True
        Resume Next
    
    Else
    
        Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
        Resume ProcExit
    
    End If
       
 'Check to see if you can connect to the Deal Site. If you can't then then tell the user of the error and have them contact ITOPursuitSite mailbox
  Case -2146697211
    MsgBox "Can not connect to the Sales Pursuit site" _
    & vbCrLf & vbCrLf & "Please, email the ITOPursuitSite mailbox with a description of this problem listed below!" _
    & vbCrLf & vbCrLf & """" & Err.Description & """", vbExclamation + vbOKOnly
    Resume ProcExit
  
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
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
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

End Sub​