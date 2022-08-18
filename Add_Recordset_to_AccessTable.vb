Public Sub Add_Recordset_From_Form(CtlSubmitBtn As Control, sTableName)

'--------------------------ADO objects--------------------------
 Dim RS As ADODB.Recordset, rsPrev As ADODB.Recordset
 Dim Conn As ADODB.Connection
 Dim Errors As ADODB.Errors
 Dim Fld As ADODB.Field
 Dim Cmd As ADODB.Command
 Dim Prm As ADODB.Parameter

'-------------------------Access objects------------------------
 Dim Ctl As Control
 Dim Frm As Form

'---------------------------Variables---------------------------
 Dim sStr As String, sFieldName As String
 Dim i As Integer
 Dim lCount As Long

'----------------------Instantiate objects----------------------
 Set Conn = New ADODB.Connection
 Set RS = New ADODB.Recordset
 Set rsPrev = New ADODB.Recordset
 Set Cmd = New ADODB.Command
 Set Prm = New ADODB.Parameter
 Set Frm = CtlSubmitBtn.Parent

'Check and loop through all ADO connection errors
 On Error GoTo AdoErr

'OLEDB connection string to Access's Jet DB
  Conn.Open StrConnectDB

'Use Jet Server
  Conn.CursorLocation = adUseServer

''Disconnect Recordset
' Conn.CursorLocation = adUseClient

'Check procedural erros with VBA
 On Error GoTo ProcErr

'Open Recordset
  RS.Open "SELECT * FROM " & sTableName, Conn, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect

''Get the last record entered into the database
  rsPrev.Open "SELECT * FROM " & sTableName & " ORDER BY " & sTableName & ".ID;", Conn, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
  rsPrev.MoveLast

  If rsPrev.RecordCount > 0 Then

    'Check for changes of values between the current Trade being entered and the last trade entered
    'Get values from the Data Entry form and create parameter from those values
     For Each Ctl In Frm.Section(0).Controls

        If Ctl.ControlType = acComboBox Or Ctl.ControlType = acTextBox Then

            sFieldName = Trim(Right(Ctl.Name, Len(Ctl.Name) - 4))

            'MsgBox "Control Name " & Ctl.NAME & " = " & TypeName(Ctl.VALUE) & vbCrLf & _
                "Field Name  " & rsPrev.Fields(sFieldName).NAME & " = " & TypeName(rsPrev.Fields(sFieldName).VALUE)

          'Don't check Update_Import_Date or ID field
            If Not Ctl.Value = rsPrev.Fields(sFieldName).Value Then

                'List the name of the field updated
                sStr = sStr & "Name " & sFieldName & " = " & Ctl.Value & _
                        " Previous Value = " & rsPrev.Fields(sFieldName).Value & vbCrLf

            End If

        End If


      Next

    Else

        sStr = "There are currently ZERO deals entered in the table!!!"

    End If

'Add New Record to table
'If the trade entered is exactly like the previous traded exit the procedure it will produce a zero string length
  If Len(sStr) > 0 Then

'      'Check to see if you want to add the trade
        If MsgBox("These are the field that have changed from the last Trade entered " & vbCrLf & vbCrLf & sStr & vbCrLf & "Do you want to add this trade" & vbCrLf, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then

            'ADD NEW RECORD
             RS.AddNew

            'Get values from the Data Entry Form and put them into each field in the new recordset you just created
             For Each Ctl In Frm.Section(0).Controls

                If Ctl.ControlType = acComboBox Or Ctl.ControlType = acTextBox Then



                    'Get Name of Table field from control name
                    'NOTE You need to have the control name based on the table field name you are inputting to
                    '     You could equally use the record set field name, but the recordset would return all field in the table
                    '      not the one just being inputted to from the form
                    sFieldName = Trim(Right(Ctl.Name, Len(Ctl.Name) - 4))

                    'Create parameter and append to cmd object
                    'NOTE This ensure that the data coming from the Form is the right data type and right data size

                    Set Prm = Cmd.CreateParameter(, RS(sFieldName).Type, adParamInput, RS(sFieldName).DefinedSize, Ctl.Value)
                    'Set Prm = Cmd.CreateParameter(sStr, rs(sStr).TYPE, adParamInput, rs(sStr).DefinedSize, Ctl.value)

                    'Append each parameter created
                    '  Cmd.Parameters.Append Prm

                    'Set Parameter values to record set
                    RS(sFieldName).Value = Prm.Value

                End If

            Next Ctl

          'THIS IS WHERE YOU UPDATE THE RECORDSET
            RS.UpdateBatch


          'BACK UP THE RECORDSET JUST ENTERED
            Add_to_Data_BackUp CtlSubmitBtn

        Else

          'Exit Sub
            Frm.Controls("cbo_Asset").SetFocus
            GoTo ProcExit

        End If
  Else

    MsgBox "This is a Duplicate Trade" & vbCrLf & vbCrLf & "This Trade hasn't been entered!", vbCritical

   'Exit Sub
    Frm.Controls("cbo_Asset").SetFocus
    GoTo ProcExit

  End If

  ''Set ActiveConnection to complete disconnecting the recordset
  ' rs.ActiveConnection = Nothing

ProcExit:

  DoCmd.Hourglass False

'Close Recordset Object
  RS.Close
  Set RS = Nothing
  rsPrev.Close
  Set rsPrev = Nothing

'Close Connection object
  Conn.Close
  Set Conn = Nothing

Exit Sub

AdoErr:

Dim ErrLoop As Error

Set Errors = Conn.Errors
    For Each ErrLoop In Errors
        MsgBox "Description " & ErrLoop.Description & vbCrLf & "The Error # is " & ErrLoop.Number & vbCrLf & "The source " & ErrLoop.Source, vbCritical
    Next
Resume ProcExit

ProcErr:
  DoCmd.Hourglass False

  Select Case Err.Number

  Case -2147217900 'Missing SQL Statement
    Resume ProcExit

  Case 3021 'BOF or EOF not found
    Resume Next

  Case 3625 'The array of field being input dont match the recordset field names
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Resume Next

  Case 3704 'Recordset empty End program to stop more errors
    Resume Next

  Case Else
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    Resume Next
  End Select
Resume ProcExit

End Sub