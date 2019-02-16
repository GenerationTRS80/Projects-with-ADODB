Option Compare Database
Public Function fn_Append_tblForecastNumber(sTempVar_Name As String, sDeleteQuery As String, sAppendQuery As String)

 'Access Objects
  Dim qd As DAO.QueryDef
  
 'Local Variables
  Dim sInputValue As String
  Dim lMonthNumber As Long
  Dim tmpV_Forecat_Number As TempVar
  Dim sSQL As String

 On Error GoTo ProcErr
 
  DoCmd.SetWarnings False
 
  Set tmpV_Forecat_Number = TempVars(sTempVar_Name)
  lMonthNumber = tmpV_Forecat_Number.Value
  
 'Debug.Print tmpV_Forecat_Number.Name
  
 'Delete Rows SQL
  sSQL = sSQL & "DELETE tblForecast_" & lMonthNumber & ".*" & vbCrLf
  sSQL = sSQL & "FROM tblForecast_" & lMonthNumber & ";"
  
  'Debug.Print sSQL
 
 'Set query definition
  DoCmd.DeleteObject acQuery, sDeleteQuery
  Set qd = CurrentDb.CreateQueryDef(sDeleteQuery, sSQL)
 ' Set qd = CurrentDb.CreateQueryDef("", sSQL)

 'Append
  qd.Execute

 'Append SQL
  sSQL = "INSERT INTO tblForecast_" & lMonthNumber & vbCrLf
  sSQL = sSQL & "([Cost Center], [SLT Short], SLT, LT, Account, [Acct Description]" & vbCrLf
  sSQL = sSQL & ", Jun, Jul, Aug, Sep, Oct, Nov, [Dec], Jan, Feb, Mar, Apr, May, Total, Q1, Q2, Q3, Q4 )" & vbCrLf
  sSQL = sSQL & "SELECT Final_C.[Cost Center], Final_C.[SLT Short], Final_C.SLT, Final_C.LT, Final_C.Account, Final_C.[Acct Description]," & vbCrLf
  sSQL = sSQL & "Final_C.Jun, Final_C.Jul, Final_C.Aug, Final_C.Sep, Final_C.Oct, Final_C.Nov, Final_C.Dec, Final_C.Jan" & vbCrLf
  sSQL = sSQL & ", Final_C.Feb, Final_C.Mar, Final_C.Apr, Final_C.May, Final_C.Total, Final_C.Q1, Final_C.Q2, Final_C.Q3, Final_C.Q4 "
  sSQL = sSQL & "FROM [tblForecastConsolidation-C_Final] as Final_C;"
  sSQL = sSQL & ""

 ' Debug.Print sSQL
 
 'Set query definition
  DoCmd.DeleteObject acQuery, sAppendQuery
  Set qd = CurrentDb.CreateQueryDef(sAppendQuery, sSQL)
 
  'Set qd = CurrentDb.CreateQueryDef("", sSQL)

 'Append
  qd.Execute

  MsgBox "Completed appending from tblForecastConsolidationC-Final" & vbCrLf & vbCrLf & _
          "into tblForecast" & lMonthNumber, vbInformation + vbOKOnly, "Append into tblForecast" & lMonthNumber
  
ProcExit:

  DoCmd.SetWarnings True
 
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit
    
  Case 7874 'Object doesn't exist. Go to next step
    Resume Next

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit

End Function
Public Function fn_Macro_RunSaved_Export()

  
  Dim sFilePath As String
  
  sFilePath = TempVars("tmpV_FilePath").Value
  
  'Debug.Print sFilePath
  
 'Export tbl_LoadSAP_Forecast PRA and PRD
  DoCmd.TransferText acExportDelim, , "tbl_LoadSAP_Forecast_PRD_Final", sFilePath & "LoadSAP_Forecast_PRD_Final.csv"
  DoCmd.TransferText acExportDelim, , "tbl_LoadSAP_Forecast_PRA_Final", sFilePath & "LoadSAP_Forecast_PRA_Final.csv"
  
 
 
End Function
Public Function List_TempVars()

 Dim TempVar As TempVar
 Dim lCountTempVar As Long
 
 
 'Application.CurrentProject.AllMacros.
 lCountTempVar = TempVars.Count
 
 Debug.Print TempVars.Count
 
'List Variables
 For Each TempVar In TempVars
 
  Debug.Print "TempVar " & TempVar.Name & " = " & TempVar.Value & " Type is " & TypeName(TempVar.Value)
 
 Next
 
 Debug.Print TempVars("ForecastMonth_ID").Value

End Function
Public Function fn_Select_SLT_Column(sTempVar_Name As String)

  Dim qd As DAO.QueryDef
  Dim sSLT_DefaultValue As String
  Dim sInputValue As String
  Dim tmpV_Select_SLTColumn As TempVar

 On Error GoTo ProcErr
 
  Set tmpV_Select_SLTColumn = TempVars(sTempVar_Name)
  
  Debug.Print tmpV_Select_SLTColumn.Name
 
  Set qd = CurrentDb.CreateQueryDef("", "Select * from tbl_SLT_Default")
  sSLT_DefaultValue = qd.OpenRecordset.Fields(1).Value

 'Set Input box to temp var value
  sInputValue = InputBox("Enter the SLT Column you want to use SLT1 is the default. The column range is 1 through 5 ", _
                                                     "Select SLT Column", sSLT_DefaultValue)
                                                     
 'Check for null value
  If Len(Trim(sInputValue)) = 0 Then
  
    MsgBox "A blank was entered" & vbcrfl & "The default value " & sSLT_DefaultValue & " will be used", vbOKOnly + vbInformation
    sInputValue = sSLT_DefaultValue
    
    GoTo ProcExit
    
  End If
  

 'If the default SLT value is different than the SLT value entered. Then ask the user if they want to update the default value with the current value
  If sSLT_DefaultValue <> sInputValue Then
  
    If MsgBox("You have enter a different value " & sInputValue & " Than the default value of " & sSLT_DefaultValue & vbCrLf & vbCrLf & _
          "Do you want to update the default value to the current value you have entered?", _
          vbInformation + vbYesNo + vbDefaultButton2, "Update Default Value") = vbYes Then
         
          Set qd = CurrentDb.CreateQueryDef("", "UPDATE tbl_SLT_Default SET tbl_SLT_Default.SLT_Default =""" & sInputValue & """;")
             
         'Update the tbl_SLT_Default table
          qd.Execute
             
     End If
     
   End If
   
 'Set Temp Var
  Debug.Print "Input value " & sInputValue
  tmpV_Select_SLTColumn.Value = sInputValue

 
ProcExit:
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit
    
  Case 7874 'Table doesn't exist. It does not need to be deleted
    Resume ProcExit

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit

End Function
Public Function fn_macro_Delete_Table(sTableName As String)

 On Error GoTo ProcErr
  
    DoCmd.DeleteObject acTable, sTableName
    
ProcExit:
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit
    
  Case 7874 'Table doesn't exist. It does not need to be deleted
    Resume ProcExit

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit

End Function

Public Function Run_Macro_Create_tblForecastConsolidation_D_Final()


  fn_macro_Delete_Table "tblForecastConsolidation-D_Final"
  
  fn_macro_OpenQuery_Recompile "qry_tblImport_WorksheetForecast_RemoveZeroValue"
  
  fn_macro_RunQuery "mkt_tblForecastConsolidation-D_Final"
  

End Function
Public Function fn_macro_RunQuery(sQuery_Name As String)

 'This will remove the warnings from deleting or appending query
 
 'Access Objects
  Dim db As DAO.Database
  Dim qd As DAO.QueryDef
  
 'local variables
  Dim sSQL As String
  
 On Error GoTo ProcErr
  
  Set db = Application.CurrentDb
  Set qd = db.QueryDefs(sQuery_Name)
  
  DoCmd.SetWarnings False
  
  
 'Execute query
  qd.Execute
  
'  sSQL = db.QueryDefs(sQuery_Name).Sql
'  DoCmd.RunSQL sSQL

  'db.Execute sSQL, dbFailOnError
  
' 'Refresh Tables and Queries
' 'Note: Refresh tables needs to be done to update the table data in the Make Table SQL statement
'  db.TableDefs.Refresh


ProcExit:

  DoCmd.SetWarnings True
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit
    
  Case 3010 'Table already exists
   MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description, vbOKOnly + vbInformation, "Error handled by CostCenter Forecast Database"
    Resume ProcExit


  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit
End Function
Public Function fn_macro_OpenQuery_Recompile(sQuery_Name As String)

'This function will Open,Save then close a query to recompile it

 On Error GoTo ProcErr
 
 'This will remove the warnings from deleting or appending query
  DoCmd.SetWarnings False
  
 'Recompile query
  DoCmd.OpenQuery sQuery_Name, acViewNormal
  DoCmd.Save acQuery, sQuery_Name
  DoCmd.Close acQuery, sQuery_Name


ProcExit:

  DoCmd.SetWarnings True
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit


  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit
 
End Function
Public Function fn_macro_FindText_in_QueryDefs()

 '****************************************************************************
 '*
 '*  Author: Phil Seiersen
 '*  Date: Nov 9 2018
 '*
 '*  NOTES: Take the sql from queryDefs and turn it into a string
 '*

  Dim db As DAO.Database
  Dim qd As DAO.QueryDef
  Dim sMsg As String
  Dim sListQueries As String
  Dim iCount As Integer
  Dim sFindText As String
 
 On Error GoTo ProcErr
 
 'NOTE: Application.CurrentProject
  Set db = Application.CurrentDb

 'Count query def
  iCount = 0
  
 'Get Text to search with from dialog box in macro
  sFindText = TempVars("tmpV_FindText").Value
  sListQueries = "Text searched term " & sFindText & vbCrLf & vbCrLf
  
 'Get SQL for
  For Each qd In db.QueryDefs
  
   If InStr(1, qd.Sql, sFindText, vbTextCompare) Then
    
      iCount = iCount + 1
      sListQueries = sListQueries & Format(iCount, "00") & " " & qd.Name & vbCrLf
      sMsg = sMsg & qd.Name & vbCrLf & qd.Sql & vbCrLf
      
   End If

  Next qd

  Debug.Print sListQueries
  MsgBox sListQueries, , "List Query Definitions"
  
ProcExit:
  Exit Function

ProcErr:
  Select Case Err.Number

  Case 13 'Cancel button hit on input box
    Resume ProcExit

  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit

  End Select

 Resume ProcExit
End Function


Public Function fn_Macro_CreateTable(sQueryDefinition_Name As String, Optional sTableName As String)


'Check procedural erros with VBA
 On Error GoTo ProcErr


  Select Case sQueryDefinition_Name
  
  Case "CreateTable_tblImport_WorksheetForecast"
  
    DoCmd.DeleteObject acTable, sTableName
    fnSub_CreateTable sQueryDefinition_Name
  
  Case "CreateTable_tblImport_WorksheetDelta_Comments"
  
    DoCmd.DeleteObject acTable, sTableName
    fnSub_CreateTable sQueryDefinition_Name
  
  Case "CreateTable_tblImport_WorksheetFTERoster"
  
    DoCmd.DeleteObject acTable, sTableName
    fnSub_CreateTable sQueryDefinition_Name
    
  Case "CreateTable_tblImport_WorksheetETWRoster"
  
    DoCmd.DeleteObject acTable, sTableName
    fnSub_CreateTable sQueryDefinition_Name
    
  Case "CreateTable_tblImport_WorksheetOther"
  
    DoCmd.DeleteObject acTable, sTableName
    fnSub_CreateTable sQueryDefinition_Name
    
  Case "CreateTable_tblImport_WorksheetDepreciation"

  DoCmd.DeleteObject acTable, sTableName
  fnSub_CreateTable sQueryDefinition_Name
       
  Case Else
  
    GoTo ProcExit
  
 End Select


ProcExit:

  Exit Function

ProcErr:
  Select Case Err.Number

  Case 7874 'Table to delete was not found
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume Next
    
  Case -2147217900
    'MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical, "Error handled by FPA Database"
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit


  Case Else
    MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Stop
    Resume Next

  End Select

Resume ProcExit

End Function
Private Function fnSub_CreateTable(sQueryDefinition_Name As String, Optional sTableName As String)

  Dim sSQL As String

  Select Case sQueryDefinition_Name
  
  Case "CreateTable_tblImport_WorksheetForecast"
  
   'NOTE: to set a datetime field default value it needs to be done using ADO
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql

   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)

   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
  Case "CreateTable_tblImport_WorksheetFTERoster"
   
   'NOTE: to set a datetime field default value it needs to be done using ADO
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
    
   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)
   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
  Case "CreateTable_tblImport_WorksheetETWRoster"
   
   'NOTE: to set a datetime field default value it needs to be done using ADO
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
    
   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)
   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
  Case "CreateTable_tblImport_WorksheetDepreciation"
   
   'NOTE: to set a datetime field default value it needs to be done using ADO
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
    
   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)
   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
  Case "CreateTable_tblImport_WorksheetDelta_Comments"
  
   'NOTE: to set a datetime field default value it needs to be done using ADO
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
    
   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)
   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
    
  Case "CreateTable_tblPLRollup_Structure"
  
   'SQL does not need to be modified at this time 12/10/2018 Phil Seiersen
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
  
  Case Else
  
  'Create the ForecastMonth table name by concatenating tbl_ with the number entered in the input box
    sSQL = CurrentDb.QueryDefs(sQueryDefinition_Name).Sql
  
   'Remove trailing Parenthese
    sSQL = Left(sSQL, Len(sSQL) - 1)
   'Add default value to Update Upload_Forecast_Date
    sSQL = Trim(sSQL) & " DEFAULT NOW() NOT NULL"
    sSQL = sSQL & ")"
    
  End Select
  
  'Debug.Print sSQL
  
 ' >>>>> EXECUTE QUERY  <<<<<<<
  CurrentProject.Connection.Execute sSQL

  
End Function
Private Function fn_CreateTable_Set_Field_DefaultValue(sTableName As String, sColumnName As String, sDataType As String, sDefaultValue As String, Optional bNOTNULL As Boolean = False)

  Dim sSQL As String

  sSQL = "ALTER TABLE " & sTableName
  sSQL = sSQL & "ADD COLUMN " & sColumnName
  sSQL = sSQL & " " & sDateType
  sSQL = sSQL & " DEFAULT " & sDefaultValue

  If bNOTNULL Then
  
    sSQL = sSQL & " NOT NULL"

  End If

 ' >>>>> EXECUTE QUERY  <<<<<<<
  CurrentProject.Connection.Execute sSQL
  
  

End Function