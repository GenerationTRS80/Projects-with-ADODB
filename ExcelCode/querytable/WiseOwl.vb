Public Sub Import_Loop_WiseOwlCourses()

  Dim xlWrkSht_Download As Excel.Worksheet
  Dim xlRange_Destination As Excel.Range
  Dim qtbl As Excel.QueryTable
  
  
 'String variables
  Dim sURL As String
  Dim sConn As String
  Dim sDestination As String
  
  
  sURL = "https://www.wiseowl.co.uk/courses/"
  sConn = "URL;" & sURL
  
 'Set xlWrkSht_Download Add
  Set xlWrkSht_Download = Worksheets.Add
  
  Set xlRange_Destination = xlWrkSht_Download.Range("A1")
  
 'Create QueryTable
  Set qtbl = xlWrkSht_Download.QueryTables.Add(sConn, xlRange_Destination)
  
 'Set QueryTable properties
  With qtbl
    .RefreshOnFileOpen = True
    .RefreshPeriod = 0 'Disable automatic refresh
    .Name = "WiseOwlCourse"
    .WebFormatting = xlWebFormattingRTF 'Don't get the url link in the text. Just values
    .WebSelectionType = xlSpecifiedTables
    .WebTables = "2,3"
    .Refresh
  End With
  

' 'Execute the Query in the QueryTable by refresing
'  qtbl.Refresh
  
End Sub
Public Sub Import_WiseOwlCourses()

  Dim xlWrkSht_Download As Excel.Worksheet
  Dim xlRange_Destination As Excel.Range
  Dim qtbl As Excel.QueryTable
  
  
 'String variables
  Dim sURL As String
  Dim sConn As String
  Dim sDestination As String
  
  
  sURL = "https://www.wiseowl.co.uk/courses/"
  sConn = "URL;" & sURL
  
 'Set xlWrkSht_Download Add
  Set xlWrkSht_Download = Worksheets.Add
  
  Set xlRange_Destination = xlWrkSht_Download.Range("A1")
  
 'Create QueryTable
  Set qtbl = xlWrkSht_Download.QueryTables.Add(sConn, xlRange_Destination)
  
 'Set QueryTable properties
  With qtbl
    .RefreshOnFileOpen = True
    .RefreshPeriod = 0 'Disable automatic refresh
    .Name = "WiseOwlCourse"
    .WebFormatting = xlWebFormattingRTF 'Don't get the url link in the text. Just values
    .WebSelectionType = xlSpecifiedTables
    .WebTables = "2,3"
    .Refresh
  End With
  

' 'Execute the Query in the QueryTable by refresing
'  qtbl.Refresh
  
End Sub

Public Sub Import_Xrates()
  Dim xlWrkSht_Download As Excel.Worksheet
  Dim xlRange_Destination As Excel.Range
  Dim qtbl As Excel.QueryTable
  
  
 'String variables
  Dim sURL As String
  Dim sConn As String
  Dim sDestination As String
  
  
  sURL = "https://www.x-rates.com/table/?from=USD&amount=1"
  sConn = "URL;" & sURL
  
 'Set xlWrkSht_Download Add
  Set xlWrkSht_Download = Worksheets.Add
  
  Set xlRange_Destination = xlWrkSht_Download.Range("A5")
  
 'Create QueryTable
  Set qtbl = xlWrkSht_Download.QueryTables.Add(sConn, xlRange_Destination)
  
 'Set QueryTable properties
  With qtbl
    .RefreshOnFileOpen = True
    .RefreshPeriod = 0 'Disable automatic refresh
    .Name = "Xrates"
    .WebFormatting = xlWebFormattingRTF 'Don't get the url link in the text. Just values
    .WebSelectionType = xlSpecifiedTables
    .WebTables = "1"
    .MaintainConnection = True
    .Refresh
  End With
  

' 'Execute the Query in the QueryTable by refresing
'  qtbl.Refresh
  
End Sub
