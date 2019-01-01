Private Function FN_Write_SQLUnion_Pivot(dStart_WeekDay As Date, _
                                lStart_WeekNumber As Long, _
                                iNumberWeeks_Ahead As Integer, _
                                sTblEmployee_Opp As String, _
                                sTblEmployee As String, _
                                sTblOpportunity As String, _
                                sTblClient As String) As String

    Dim sWeekYearNumber As String
    Dim dWeekDay As Date
    Dim lYear As Long
    Dim lYearColumnAdjustment
    Dim sExcel_ColumnRef As String
    
   'Table Vars
    Dim sTblEmployee_SD As String
    Dim sTblEmployeeTeam As String
    Dim sTblHourCategory As String
    Dim sTblStatus As String
    

    sTblEmployee_SD = sTblEmployee  'Solution Director Tables
    sTblEmployee_BidMgr = sTblEmployee 'Bid Manager table based off  employee table
    sTblEmployeeTeam = "[Employee_Team$A3:V13]"
    sTblHourCategory = "[Hour_Category$A3:C20]"
    sTblStatus = "[New_Status$A3:B13]"
    sTblSkill_CategoryPrimary = "[Skill_Category$A3:N103]"
    sTblSkill_CategorySecondary = "[Skill_Category$A3:N103]"
    sTblTower = "[Tower$A3:M100]"
    sTblEmployee_OrgUnit = "[Employee_Org_Unit$A3:H15]"
    
    For i = 0 To iNumberWeeks_Ahead
    
        dWeekDay = DateAdd("d", 7 * i, dStart_WeekDay)
        lYear = Year(dStart_WeekDay)
        sWeekYearNumber = Format(i + lStart_WeekNumber, "00") & lYear
        
          
      '*** Calculate the column to reference for weekly hours ***
        If lYear = 2016 Then
        
            lYearColumnAdjustment = -22
        
        Else
        
           'Make adjustment in columns for Year 2017
           'NOTE: In the series of Weekly hour columns there is a non Weekly hour column on the 21st column in the series of column
            If i + lStart_WeekNumber < 21 Then
            
                lYearColumnAdjustment = 30
                
            Else
            
                lYearColumnAdjustment = 31
            
            End If
        
        End If
        
        
        sExcel_ColumnRef = "F" & i + lStart_WeekNumber + lYearColumnAdjustment
        
        
       ' >>>>>>>>>>>>   Write SQL statement   <<<<<<<<<<<<
        If i = 0 Then

             sSQL = sSQL & "Select * From ("
             sSQL = sSQL & "Select #" & dWeekDay & "# as [Week Date]" & vbCrLf

        Else

            sSQL = sSQL & vbCrLf & "Union " & vbCrLf
            sSQL = sSQL & "Select * From ("
            sSQL = sSQL & "Select #" & dWeekDay & "# as [Week Date]" & vbCrLf

        End If

           sSQL = sSQL & ", tblPivot.[F1] as [PivotTbl Key]" & vbCrLf
           sSQL = sSQL & ", tblPivot.[F2] as [Employee_Opp Desc]" & vbCrLf
           sSQL = sSQL & ", IIf(ISNULL(tblPivot.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblPivot.[" & sExcel_ColumnRef & "],""0"")))  as [Hours per Week]" & vbCrLf
       '    sSQL = sSQL & ", tblPivot.[F6] as OppID, Clng(Format(tblPivot.[" & sExcel_ColumnRef & "],""0"")) as [Hours per Week]" & vbCrLf
          
          'Filtered Employee Opportunity records
         
          'Employee Fields
           sSQL = sSQL & ", tblEmp.[F2] AS [Full Name], tblEmp.[F3] AS [Last Name], tblEmp.[F5] AS [First Name]" & vbCrLf
          
          
          'Employee Skills
           sSQL = sSQL & ", tblSkill_Primary.[F2] as [Primary Skill]" & vbCrLf
           sSQL = sSQL & ", tblSkill_Secondary.[F2] as [Secondary Skill]" & vbCrLf
           sSQL = sSQL & ", tblTower.[F2] as [Tower]" & vbCrLf
         
          'Employee Team Fields
           sSQL = sSQL & ", tblOrgUnit.[F2] AS [Organization Unit], tblEmpTeam.[F2] AS [Team Name], tblEmpTeam.[F7] AS [Team Sort]" & vbCrLf
          
          'Employee Filter
           sSQL = sSQL & ", tblEmp.[F12] AS [Scott Archer Report Filter], tblEmp.[F13] AS [David L Report Filter] , tblEmp.[F14] AS [Jennifer Shea Report Filter]" & vbCrLf

          'Opportunity Status
           sSQL = sSQL & ", tblStatus.[F2] as [Status]" & vbCrLf

          'Opportunity Fields
           sSQL = sSQL & ", tblOpp.[F2] as [Nessie ID]" & vbCrLf
           sSQL = sSQL & ", tblOpp.[F3] as [Opp Name]" & vbCrLf
           sSQL = sSQL & ", tblOpp.[F5] as [Opp Desc]" & vbCrLf
           sSQL = sSQL & ", tblOpp.[F41] as [Opportunity Updated by Employee]" & vbCrLf

          'Client
           sSQL = sSQL & ", tblClient.[F2] as [Client Name]" & vbCrLf
           sSQL = sSQL & ", tblClient.[F7] as [Client Account Group]" & vbCrLf
           
          'Solution Directors and  Bid Manager Opportunities
           sSQL = sSQL & ", tblEmp_SolutionDir.[F2] as [Solution Dir Opportunities List]" & vbCrLf
           sSQL = sSQL & ", tblEmp_BidMgr.[F2] as [Bid Manager Opportunities List]" & vbCrLf

          'Hour Category
           sSQL = sSQL & ", tblHourCat.[F2] as [Hour Category]" & vbCrLf
           sSQL = sSQL & ", tblHourCat.[F3] as [Hour Caterory Sort]" & vbCrLf

          'Opportunity Comment Fields
           sSQL = sSQL & ", tblOpp.[F28] as [Opp Issues Risks]" & vbCrLf
           sSQL = sSQL & ", tblOpp.[F29] as [Opp Status NextStep]" & vbCrLf
          
         
          'Employee Opportunity fields
'           sSQL = sSQL & ", tblPivot.[F77] as [Archive Date] " & vbCrLf
           sSQL = sSQL & ", tblPivot.[F82] as [SharePoint Editor]" & vbCrLf
           sSQL = sSQL & ", tblPivot.[F84] as [SharePoint Modified Date]" & vbCrLf
           sSQL = sSQL & ", tblPivot.[F85] as [SharePoint Created Date]" & vbCrLf
'           sSQL = sSQL & ", IIf(ISNULL(tblPivot.[F77]),FALSE,TRUE) as [Statging Update YesNo]" & vbCrLf
'           sSQL = sSQL & ", IIf([Statging Update YesNo]" & vbCrLf

           sSQL = sSQL & ", IIf(IIf(ISNULL(tblPivot.[F77]),FALSE,TRUE)" & vbCrLf
           sSQL = sSQL & ", IIf([SharePoint Modified Date]<DateSerial(2017,2,10)" & vbCrLf
           sSQL = sSQL & ", tblPivot.[F77],[SharePoint Modified Date])"
           sSQL = sSQL & ", [SharePoint Modified Date]) as [Forecast Update]"
          
         '*From Clause
         
          sSQL = sSQL & " From " & sTblEmployee_OrgUnit & " as tblOrgUnit" & vbCrLf
           
          sSQL = sSQL & " INNER JOIN (" & sTblTower & " as tblTower" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblSkill_CategorySecondary & " as tblSkill_Secondary" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblSkill_CategoryPrimary & " as tblSkill_Primary" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblEmployee_BidMgr & " as tblEmp_BidMgr" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblEmployee_SD & " as tblEmp_SolutionDir" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblStatus & " as tblStatus" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblClient & " as tblClient" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblHourCategory & " as tblHourCat" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblOpportunity & " as tblOpp" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblEmployeeTeam & " as tblEmpTeam" & vbCrLf
          sSQL = sSQL & " INNER JOIN (" & sTblEmployee & " as tblEmp" & vbCrLf
         
          sSQL = sSQL & " INNER JOIN " & sTblEmployee_Opp & " as tblPivot" & vbCrLf
          sSQL = sSQL & " ON tblEmp.[F1]=tblPivot.[F5])" & vbCrLf
          sSQL = sSQL & " ON tblEmpTeam.[F1]=tblEmp.[F9])" & vbCrLf
          sSQL = sSQL & " ON tblOpp.[F1]=tblPivot.[F6])" & vbCrLf
          sSQL = sSQL & " ON tblHourCat.[F1]=tblOpp.[F35])" & vbCrLf
          sSQL = sSQL & " ON tblClient.[F1]=tblOpp.[F21])" & vbCrLf
          sSQL = sSQL & " ON tblStatus.[F1]=tblOpp.[F23])" & vbCrLf
          sSQL = sSQL & " ON tblEmp_SolutionDir.[F1]=tblOpp.[F22])" & vbCrLf
          sSQL = sSQL & " ON tblEmp_BidMgr.[F1]=tblOpp.[F25])" & vbCrLf
          sSQL = sSQL & " ON tblSkill_Primary.[F1]=tblEmp.[F15])" & vbCrLf
          sSQL = sSQL & " ON tblSkill_Secondary.[F1]=tblEmp.[F16])" & vbCrLf
          sSQL = sSQL & " ON tblTower.[F1]=tblEmp.[F17])" & vbCrLf
          
          sSQL = sSQL & " ON tblOrgUnit.[F1]=tblEmp.[F18]" & vbCrLf

         'Where Clause
'          sSQL = sSQL & " ) as tblPivotHours"
         
          sSQL = sSQL & " Where tblPivot.[" & sExcel_ColumnRef & "] >0) as tblPivotHours"
'          sSQL = sSQL & " Where tblPivot.[" & sExcel_ColumnRef & "] Is Not Null ) as tblPivotHours"

 '         sSQL = sSQL & " WHERE tblPivotHours.[Hour Category]= ""Deal Hours""  ;"
 '         sSQL = sSQL & " WHERE tblPivotHours.[Status]= ""DEAD"" ;"
 '         sSQL = sSQL & " WHERE tblPivotHours.[OppID]= 29 ;"
          
          sSQL = sSQL & " ORDER By tblPivotHours.[Week Date], tblPivotHours.[Full Name], tblPivotHours.[Hour Caterory Sort] "
          sSQL = sSQL & ",tblPivotHours.[Status] ,tblPivotHours.[Client Name] , tblPivotHours.[Opp Name]"


    Next i
    
    FN_Write_SQLUnion_Pivot = sSQL

End Function