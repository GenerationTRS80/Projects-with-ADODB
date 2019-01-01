Private Function FN_Write_SQL_IssueRisks(sTblEmployee_Opp As String, _
                                sTblEmployee As String, _
                                sTblOpportunity As String, _
                                sTblClient As String) As String

    Dim sWeekYearNumber As String
    Dim dWeekDay As Date
    Dim sFormat_WeekDay As String
    Dim lYear As Long
    Dim lYearColumnAdjustment
    Dim sExcel_ColumnRef As String
    Dim sSQL As String
    Dim sSQL_Selected_HoursPerWeek
    
   'Table Vars
    Dim sTblEmployee_SDirector As String
    Dim sTblEmployee_BidMgr As String
    Dim sTblEmployeeTeam As String
    Dim sTblHourCategory As String
    Dim sTblStatus As String
    

    sTblEmployee_SDirector = sTblEmployee  'Solution Director Tables
    sTblEmployee_BidMgr = sTblEmployee 'Bid Manager table based off  employee table
    sTblEmployee_SharePoint_ID = sTblEmployee
    sTblEmployeeTeam = "[Employee_Team$A3:N13]"
    sTblHourCategory = "[Hour_Category$A3:R16]"
    sTblStatus = "[New_Status$A3:B13]"
    sTblSkill_CategoryPrimary = "[Skill_Category$A3:N103]"
    sTblSkill_CategorySecondary = "[Skill_Category$A3:N103]"
    sTblTower = "[Tower$A3:M100]"
    sTblEmployee_OrgUnit = "[Employee_Org_Unit$A3:H15]"
    
 
   ' >>>>>>>>>>>>   Write SQL statement   <<<<<<<<<<<<

   sSQL = sSQL & "Select * From (Select "
   
  'Opportunity Key
   sSQL = sSQL & "tblOpp.[F1] as [Opp Key]" & vbCrLf
   
  'Opportunity Table Modified Date
  'NOTE Opportunity Last Updated By is the Full Name field from the employee table joined by SharePoint Editor
   sSQL = sSQL & ", tblOpp.[F43] as [Update Opportunity]" & vbCrLf

    
  'Solution Director
   sSQL = sSQL & ", tblEmp_SolutionDir.[F2] as [Solution Director]" & vbCrLf

  'Opportunity Status
   sSQL = sSQL & ", tblStatus.[F2] as [Status]" & vbCrLf
   
  'Opportunity Forecast ID
   sSQL = sSQL & ", tblOpp.[F2] as [Forecast ID]" & vbCrLf

  'Client
   sSQL = sSQL & ", tblClient.[F2] as [Client Name]" & vbCrLf
      
'  'Account group from Client table
'   sSQL = sSQL & ", tblClient.[F7] as [Client Account Group]" & vbCrLf

  'Opportunity Fields
   sSQL = sSQL & ", tblOpp.[F3] as [Opportunity Name]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F5] as [Opportunity Description]" & vbCrLf

  'Opportunity Comment Fields
   sSQL = sSQL & ", Trim(tblOpp.[F29]) as [Status and NextStep]" & vbCrLf
   sSQL = sSQL & ", Trim(tblOpp.[F28]) as [Issues and Risks]" & vbCrLf

  'Opportunity ARR, TCV, Prob, Startegy, Solution, Offer
   sSQL = sSQL & ", IIf(IsNull(tblOpp.[F6]),Null,IIf(tblOpp.[F6]>1000000,FormatCurrency(tblOpp.[F6]/1000000,0),FormatCurrency(tblOpp.[F6],0))) as [Opp ARR]" & vbCrLf
   sSQL = sSQL & ", IIf(IsNull(tblOpp.[F7]),Null,IIf(tblOpp.[F7]>1000000,FormatCurrency(tblOpp.[F7]/1000000,0),FormatCurrency(tblOpp.[F7],0))) as [Opp TCV]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F8] as [Opp Terms]" & vbCrLf
   sSQL = sSQL & ", IIf(IsNull(tblOpp.[F9]),Null,IIf(tblOpp.[F9]>1,FormatPercent(tblOpp.[F9]/100,0),FormatPercent(tblOpp.[F9],0))) as [Opp Prob]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F10] as [Opp Strategy], tblOpp.[F11] as [Opp Solution]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F12] as [Opp Offer Review], tblOpp.[F13] as [Opp Orals]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F14] as [Opp Downselect], tblOpp.[F15] as [Opp Due Diligence]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F16] as [Opp BAFO], tblOpp.[F17] as [Opp Negotiation]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F18] as [Opp Close], tblOpp.[F19] as [Opp Handover]" & vbCrLf

  'Solution Directors and  Bid Manager Opportunities
   sSQL = sSQL & ", tblEmp_BidMgr.[F2] as [Bid Manager Opportunities List]" & vbCrLf
  
   
  ' "Updated Dated by" fields
  'These are the field that list the last person to Update the Opportunity and their Team and Sort order of their team
   sSQL = sSQL & ", tblEmp_SharePointID.[F2] as [Opportunity Updated By]" & vbCrLf
   
  '*Sort Fields
   sSQL = sSQL & ", tblEmpTeam.[F2] AS [Opportunity Updated By Team Name]" & vbCrLf
   
  'Calculation: IF Issue and Risks and Status Next Dtep are Null, then set sort to 99 Else sort by Employee Sort
  ' sSQL = sSQL & ", IIf(IsNull(tblOpp.[F29]) And IsNull(tblOpp.[F28]),1111,tblEmpTeam.[F7]) AS [Opportunity Updated By Team Sort]" & vbCrLf
   sSQL = sSQL & ", IIf(IsNull(tblOpp.[F29]) And IsNull(tblOpp.[F28]),2,1) AS [Opportunity Updated By Team Sort]" & vbCrLf
   
   
  '*From Clause
   sSQL = sSQL & " From " & sTblEmployee_OrgUnit & " as tblOrgUnit" & vbCrLf
   
   sSQL = sSQL & " INNER JOIN (" & sTblEmployeeTeam & " as tblEmpTeam" & vbCrLf
   sSQL = sSQL & " INNER JOIN (" & sTblEmployee_SharePoint_ID & " as tblEmp_SharePointID" & vbCrLf

   sSQL = sSQL & " INNER JOIN (" & sTblEmployee_BidMgr & " as tblEmp_BidMgr" & vbCrLf
   sSQL = sSQL & " INNER JOIN (" & sTblEmployee_SDirector & " as tblEmp_SolutionDir" & vbCrLf
   sSQL = sSQL & " INNER JOIN (" & sTblStatus & " as tblStatus" & vbCrLf
   sSQL = sSQL & " INNER JOIN (" & sTblClient & " as tblClient" & vbCrLf
   sSQL = sSQL & " INNER JOIN (" & sTblHourCategory & " as tblHourCat" & vbCrLf
   
   sSQL = sSQL & " INNER JOIN " & sTblOpportunity & " as tblOpp" & vbCrLf

   sSQL = sSQL & " ON tblHourCat.[F1]=tblOpp.[F35])" & vbCrLf
   sSQL = sSQL & " ON tblClient.[F1]=tblOpp.[F21])" & vbCrLf
   sSQL = sSQL & " ON tblStatus.[F1]=tblOpp.[F23])" & vbCrLf
   sSQL = sSQL & " ON tblEmp_SolutionDir.[F1]=tblOpp.[F22])" & vbCrLf
   sSQL = sSQL & " ON tblEmp_BidMgr.[F1]=tblOpp.[F25])" & vbCrLf
   
   sSQL = sSQL & " ON tblEmp_SharePointID.[F19]=tblOpp.[F41])" & vbCrLf
   sSQL = sSQL & " ON tblEmpTeam.[F1]=tblEmp_SharePointID.[F9])" & vbCrLf
   sSQL = sSQL & " ON tblOrgUnit.[F1]=tblEmp_SolutionDir.[F18]" & vbCrLf


 'Where Clause
 'NOTE: tblEmp.[18]=2 is equal Employee Org= "PRE-Sales" and  tblOpp.[F35]=5 is equal Hour Category= Deal Hours
 '      tblOpp.[F23] is equal Status not equal "Dead" tblOpp.[F22] is equal Team Member equal "Solution Director"
   sSQL = sSQL & " Where tblEmp_SolutionDir.[F18]=2 and tblOpp.[F35]=5 and tblOpp.[F23]<>5 and tblOpp.[F22]<>1) as tblIssueRisks"

 'Order By clause
  'sSQL = sSQL & " ORDER By tblIssueRisks.[Opportunity Updated By Team Sort], tblIssueRisks.[Status], tblIssueRisks.[Update Opportunity] Desc, tblIssueRisks.[Solution Director]"
  
   sSQL = sSQL & " ORDER By tblIssueRisks.[Opportunity Updated By Team Sort], tblIssueRisks.[Update Opportunity] Desc"
   sSQL = sSQL & " , tblIssueRisks.[Client Name], tblIssueRisks.[Opportunity Name]"


 '*** Set SQL to return value ***
   FN_Write_SQL_IssueRisks = sSQL

End Function