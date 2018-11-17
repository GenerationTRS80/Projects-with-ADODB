Option Compare Database
'**************************************************************************************************************
'*  Rev 4
'*  Date 9/16/03
'*  By Phil Seiersen
'*
'**************************************************************************************************************

Private Sub RunCode()
Dim dBeginDate As Date, dEndDate As Date

    dBeginDate = #3/1/2003#
    dEndDate = #3/28/2003#

'    BuildSQL_TITLEPTS dBeginDate, dEndDate

'    Export_CSV BuildSQL_TITLEPTS(dBeginDate, dEndDate), "C:\GMS_Reporting\TitlePts.xls"
'    Export_CSV BuildSQL_TITLECHARGE(dBeginDate, dEndDate), "C:\GMS_Reporting\TitleCharge.xls"
'    Export_CSV BuildSQL_CUSTODY(dBeginDate, dEndDate), "C:\GMS_Reporting\Custody.xls"
'    Export_CSV BuildSQL_STATCONT(dBeginDate, dEndDate), "C:\GMS_Reporting\Statcont.xls"
'    Export_CSV BuildSQL_ACTUALPTS(dBeginDate, dEndDate), "C:\GMS_Reporting\Actual.xls"
'    Export_CSV BuildSQL_NOMTIE(dBeginDate, dEndDate), "C:\GMS_Reporting\Nomtie.xls"
'    Export_CSV BuildSQL_IMBAL(dBeginDate, dEndDate), "C:\GMS_Reporting\Imbal.xls"
    Export_CSV BuildSQL_MARGIN(dBeginDate, dEndDate), "C:\GMS_Reporting\Margin.xls"
    
    
    
'    BuildSQL_CUSTODY dBeginDate, dEndDate
'    BuildSQL_STATCONT dBeginDate, dEndDate
'    BuildSQL_IMBAL dBeginDate, dEndDate
'    BuildSQL_MARGIN dBeginDate, dEndDate
'    BuildSQL_NOMTIE dBeginDate, dEndDate
'    BuildSQL_ACTUALPTS dBeginDate, dEndDate


End Sub
Public Function BuildSQL_TITLEPTS(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String, sBeginMonth As String, sEndMonth As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "TITLEPTS"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptTitlePTsTemp
  Set qd = CurrentDb.QueryDefs("rptTitlePTsTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.POINTID, A.DEALID, A.PIPELINE, A.METER, A.ENTITY, A.PRODUCER,A.FIRSTPURCHASE, A.FIRSTPURCHASER, " & vbCrLf
        sSQL = sSQL & "A.DK" & i & " QTY, A.PRICE" & i & " PRICE, A.KCONNECT, A.VOLUMESTATUS, A.BILLINGSTATUS, A.COMPANY, A.PLANNINGGROUP, A.NOMINATIONGROUP, " & vbCrLf
        sSQL = sSQL & "A.DELIVERABILITY, A.FLOWUNIT, A.CURRENCY, A.CURRENCYUNIT, A.MONTHLYQUANTITYIND, A.MTHTOTAL, A.MONTHLYAMOUNTIND, " & vbCrLf
        sSQL = sSQL & "A.DETAILTAXINFO , A.INITIALOWNER, A.LASTCHANGE, A.LASTCHANGEUSER, A.ADDDATE, to_date(sysdate) CURRENTDATE" & vbCrLf
 
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

 'Pass SQL string to another procedure
    BuildSQL_TITLEPTS = sSQL
 
End Function

Public Function BuildSQL_ACTUALPTS(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "ACTUAL"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptActualTemp
  Set qd = CurrentDb.QueryDefs("rptActualTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.PIPELINE, A.CONTRACT, A.POINTOFVIEW, A.RECMETER, " & vbCrLf
        sSQL = sSQL & "A.DK" & i & " QTY, A.DKTOTAL, A.RECENTITY, A.DELMETER, A.DELENTITY, A.NOMTRANSACTIONTYPE, A.STATUS, " & vbCrLf
        sSQL = sSQL & "A.TPORTSTATEMENTDATE, A.TPORTSTATEMENTID, A.OVERRIDEFUEL, A.FUELPERCENT, A.FUELPOINTOFVIEW, A.NOMUNIT, A.CURRENCYUNIT, A.CURRENCY, A.ACTUALIZEUNIT, A.CALCHVFACTOR, A.TPORTCRSTATEMENTID, " & vbCrLf
        sSQL = sSQL & "A.TPORTCRSTATEMENTDATE, A.RELATEDRECORDLINK, A.PACKAGEID, A.FUELDAILYMTHLYIND, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE" & vbCrLf
 
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_ACTUALPTS = sSQL
 
End Function

Public Function BuildSQL_NOMTIE(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "NOMTIE"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptActualTemp
  Set qd = CurrentDb.QueryDefs("rptNomtieTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
       
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.NOMTRANSACTIONTYPE, A.PIPELINE, A.CONTRACT,  " & vbCrLf
        sSQL = sSQL & "A.RECMETER, A.RECENTITY, A.DELMETER, A.DELENTITY, A.USCONTRACT, A.USENTITY, A.DSCONTRACT, A.DSENTITY, A.POINTOFVIEW, A.PACKAGEID, "
        sSQL = sSQL & "A.DK" & i & " QTY, A.USDSRANK" & i & " USDRANK, A.STATUS" & i & " STATUS, A.DIFF" & i & " DIFF, A.DIFFSTAT" & i & " DIFFSTAT, A.ACTUAL" & i & " ACTUAL, " & vbCrLf
        sSQL = sSQL & "A.CAPACITYTYPEIND, A.TRACKINGNUMBER, A.ACTIVITYNUMBER, A.CONFTRACKINGNUMBER, A.GISBMODELTYPE, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE" & vbCrLf

    'Change SQL per day
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_NOMTIE = sSQL
 
 
End Function
 
Public Function BuildSQL_MARGIN(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "MARGIN"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptMarginTemp
  Set qd = CurrentDb.QueryDefs("rptMarginTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.USERTYPE, A.PRODMONTH, A.PIPELINE, A.CONTRACT, A.RECENTITY, A.DELENTITY, A.RECMETER, A.DELMETER, A.POINTOFVIEW, A.NOMTRANSACTIONTYPE, " & vbCrLf
        sSQL = sSQL & "A.DK" & i & " QTY, A.DKTOTAL, A.AMOUNT" & i & " AMOUNT, A.AMOUNTTOTAL, A.TRANSTYPE, A.FUELPERCENT, A.PARTNER, A.OVERRIDEFUEL, A.OVERRIDERATE, A.AVGPRICE, " & vbCrLf
        sSQL = sSQL & "A.SUBTYPE, A.BOOKINGUNIT, A.BOOKINGCURRENCY, A.RATEUTILIZED, A.LASTCHANGE, A.LASTCHANGEUSER, A.ENTITY, to_date(sysdate) CURRENTDATE" & vbCrLf

        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_MARGIN = sSQL
 
End Function

Public Function BuildSQL_CUSTODY(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "CUSTODY"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptCustodyTemp
  Set qd = CurrentDb.QueryDefs("rptCustodyTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.PIPELINE,  A.CONTRACT, A.POINTOFVIEW, " & vbCrLf
        sSQL = sSQL & "A.RECMETER, A.RECENTITY, A.DELMETER, A.DELENTITY, A.NOMTRANSACTIONTYPE, A.OVERRIDERATE, A.CAPACITYTYPE, A.ACTIVITYNUMBER, A.OVERRIDEFUEL, A.FUELPERCENT, " & vbCrLf
        sSQL = sSQL & "A.DK" & i & " QTY, A.DIFF" & i & " DIFF, A.GISBMODELTYPE, A.ROUTEID, A.NOMUNIT, A.IMBALUNIT, "
        sSQL = sSQL & "A.NOMENTRYUNIT, A.ACCTSTATUS, A.TRACKINGNUMBER, A.SHIPPERTIMESTAMP, A.PIPELINETIMESTAMP, A.PACKAGEID, A.BIDTRANSPORTATIONRATE, A.RELATEDRECORDLINK, " & vbCrLf
        sSQL = sSQL & "A.FUELDAILYMTHLYIND, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE" & vbCrLf

    'Change SQL per day
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_CUSTODY = sSQL
 
 
End Function
Public Function BuildSQL_STATCONT(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "STATCONT"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptCustodyTemp
  Set qd = CurrentDb.QueryDefs("rptStatContTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.PIPELINE,  A.CONTRACT, A.POINTOFVIEW, A.RECMETER, A.DELMETER, " & vbCrLf
        sSQL = sSQL & "A.RemDK" & i & " VOLUME, A.ORIGINALCONTRACT, A.ORIGINALRECMETER, A.SEGMENTRELEASE, A.RATEUNIT, A.NOMUNIT, A.PATHTYPE, "
        sSQL = sSQL & "A.PATHPRIORITY, A.RECSTATIONTYPE, A.DELSTATIONTYPE, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE" & vbCrLf

    'Change SQL per day
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL
 
    BuildSQL_STATCONT = sSQL
 
End Function
Public Function BuildSQL_TITLECHARGE(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "TITLECHARGE"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptCustodyTemp
  Set qd = CurrentDb.QueryDefs("rptTitleChargeTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.PIPELINE,  A.CONTRACT, A.POINTOFVIEW, " & vbCrLf
        sSQL = sSQL & "A.RECMETER, A.DELMETER, A.RECENTITY, DELENTITY, A.NOMTRANSACTIONTYPE, " & vbCrLf
        sSQL = sSQL & "A.DK" & i & " VOLUME,  A.AVGPRICE, A.STATEMENTDATE, A.STATEMENTID, A.ACCTYEAR, A.ACCTMONTH, A.CURRENCYUNIT, A.CURRENCY, " & vbCrLf
        sSQL = sSQL & "A.TITLEID, A.ENTITY, A.STATUS, A.STATEMENTGROUP, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE" & vbCrLf

    'Change SQL per day
        Select Case i

        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_TITLECHARGE = sSQL
 
End Function
Public Function BuildSQL_IMBAL(dBeginMonth As Date, dEndMonth As Date, Optional iFirstDay As Integer = 1, Optional iLastDay As Integer = 31) As String
  Dim sSelect As String, sSQL As String, sWhere As String
  Dim i As Integer
  Dim qd As DAO.QueryDef
  
'Define constants
  Const S_TABLE_NAME As String = "IMBAL"
  
  sBeginMonth = CStr(Year(dBeginMonth) & Format(Month(dBeginMonth), "00"))
  sEndMonth = CStr(Year(dEndMonth) & Format(Month(dEndMonth), "00"))

'Set QueryDef Object to rptImbalTemp
  Set qd = CurrentDb.QueryDefs("rptImbalTemp")
  
  qd.SQL = "A"
  
    For i = iFirstDay To iLastDay - 2
  
        sSelect = sSelect & "Select * from(" & vbCrLf
 
    Next i
 
    For i = iFirstDay To iLastDay
    
        sSQL = sSQL & "Select to_date('" & i & "-' || Decode(A.PRODMONTH,1,'JAN',2,'FEB',3,'MAR',4,'APR',5,'MAY',6,'JUN',7,'JUL',8,'AUG',9,'SEP',10,'OCT',11,'NOV',12,'DEC') ||'-'|| A.PRODYEAR) DeliveryDate, " & vbCrLf
        sSQL = sSQL & "A.PRODYEAR, A.PRODMONTH, A.PIPELINE, A.CONTRACT, A.BEGPIPELINE, A.ACCTYEAR, A.ACCTMONTH, A.CURRPIPELINE, " & vbCrLf
        sSQL = sSQL & "A.DAILYIMBAL" & i & " IMBALANCE, A.PMAPIPELINE, A.ACCUMPIPELINE, A.BEGUSER, A.CURRUSER, A.PMAUSER, A.ACCUMUSER, " & vbCrLf
        sSQL = sSQL & "A.CASHOUTUSER, A.TRADEUSER, A.CURRENTVALUE, A.ACCUMVALUE, A.CASHOUTVALUE, A.TRADEVALUE, A.PMAVALUE, " & vbCrLf
        sSQL = sSQL & "A.AVGPRICE, A.TPORTSTATEMENTDATE, A.EXPORTDATE, A.TPORTSTATEMENTID, A.BEGVALUE, A.ICSTATEMENTID, "
        sSQL = sSQL & "A.CURRENCY, A.IMBALUNIT, A.TPORTCRSTATEMENTID, A.TPORTCRSTATEMENTDATE, A.LASTCHANGE, A.LASTCHANGEUSER, to_date(sysdate) CURRENTDATE " & vbCrLf

    'Change SQL per day
        Select Case i
        
        Case 29
        
          'Calculate leap year
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2, DECODE(A.PRODYEAR,2000,'02',2004,'02',2008,'02',2012,'02',2016,'02',2020,'02',2024,'02',2028,'02',2032,'02',2036,'02',2040,'02',''),3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
                   
        Case 30
        
          'Remove end paranthese
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
        
        Case 31
        
          'Remove UNION ALL from end of statment
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE( A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) >=to_number(" & sBeginMonth & ") " & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',3,'03',5,'05',7,'07',8,'08',10,'10',12,'12','')) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        Case Else
        
          'All the other days than 1, 29, 30, 31
            sSQL = sSQL & "From  " & S_TABLE_NAME & " A " & vbCrLf
            sSQL = sSQL & "WHERE to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) >=to_number(" & sBeginMonth & ")" & vbCrLf
            sSQL = sSQL & "AND to_number(A.PRODYEAR || DECODE(A.PRODMONTH,1,'01',2,'02',3,'03',4,'04',5,'05',6,'06',7,'07',8,'08',9,'09',A.PRODMONTH)) <=to_number(" & sEndMonth & ")" & vbCrLf
            
        End Select
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i <> iLastDay And i <> iFirstDay Then
        
            sSQL = sSQL & ") Union ALL" & vbCrLf
        
        End If
        
      'If last day is reached before the 28th then don't add the SQL string below
        If i = iFirstDay And i <> iLastDay Then
        
            sSQL = sSQL & " Union ALL" & vbCrLf
                       
        End If
        
    Next i
  
  'Add where statement
    sSQL = sSelect & sSQL
    
    Debug.Print sSQL
 
 'Update Querydef with SQL string
    qd.SQL = sSQL

    BuildSQL_IMBAL = sSQL
 
End Function
Public Sub ADO_ImportExcel()
'NOTE NEED UPDATE BY IN PRICECURVE FILE

'Append Excel Sheet data to Access Table

'ADO objects
 Dim Conn As ADODB.Connection
 Dim Errors As ADODB.Errors
 Dim sExcelPathFile As String
 
'Variables
 Dim sSQL As String, sWorksheet As String
 Dim ErrLoop As Error
 
 sExcelPathFile = "C:\GMS_Reporting\TitlePts.xls"

'Check procedural erros with VBA
 On Error GoTo ProcErr
 
''Instantiate objects
 Set Conn = New ADODB.Connection
 
''Check and loop through all ADO connection errors
' On Error GoTo AdoErr
 
' MsgBox CurrentProject.AccessConnection
 
'Open Connection StrConnectDB is an OLEDB connection to the current DB
 Conn.Open StrConnectDB  'Access intrinsic connection to current database: Conn.Open CurrentProject.AccessConnection

''Connect to Spreadsheet
' sWorksheet = "[Excel 8.0;DATABASE=" & sExcelPathFile & ";HDR=YES;IMEX=1].[" & sWorksheetName & "$A1:BZ600]"

'Use SQL to Append Prices rptTitlePTsTemp
 sSQL = "SELECT * INTO [Excel 8.0;DATABASE=" & sExcelPathFile & ";HDR=YES;IMEX=1].[SHEET1] FROM rptTemp "

 DoCmd.Hourglass True
 
 Conn.Execute sSQL

ProcExit:
  
  DoCmd.Hourglass False
  
'Close Connection object
  Conn.Close
  Set Conn = Nothing
  
Exit Sub

'AdoErr:
'
'Set Errors = Conn.Errors
'
'    For Each ErrLoop In Errors
'        MsgBox "Description " & ErrLoop.Description & vbCrLf & "The Error # is " & ErrLoop.Number & vbCrLf & "The source " & ErrLoop.Source, vbCritical
'    Next
'
'Resume ProcExit

ProcErr:

  DoCmd.Hourglass False
  Select Case Err.Number
  Case -2147217900 'Missing SQL Statement
   MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    'MsgBox " The Columns in the Access Gas Curves.xls" & vbCrLf & "don't match the columns in the tbl_PriceCurves table", vbCritical
    Resume ProcExit
    
  Case -2147467259 'Invalid Path
    MsgBox "Description " & Err.Description, vbExclamation
    Resume ProcExit
    
  Case -2147217865 'Wrong Excel Sheet Name
    MsgBox "Description " & Err.Description
    'MsgBox "Excel worksheet called Price_Curves not found!", vbCritical
    Resume ProcExit
    
  Case 3625 'The array of field being input dont match the recordset field names
    MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
    Resume Next
  Case 3704 'Recordset empty End program to stop more errors
    Resume Next
  Case Else
    MsgBox "The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & "The source " & Err.Source, vbCritical
    Stop
    Resume Next
  End Select
Resume ProcExit

End Sub
​