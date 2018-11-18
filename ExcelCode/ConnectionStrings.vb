Option Compare Database

Option Explicit

Public Function StrConnectDB() As String
'This connects to the local database using OLEDB connection provider
'http://www.connectionstrings.com/

'OLEDB connection string to Access's Jet DB
     StrConnectDB = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & CurrentDb.Name & ";" & _
                    "User Id=admin;" & _
                    "Password="

End Function
Public Function strConnect(sDbPathName As String) As String
'This connects to the local database using OLEDB connection provider
'http://www.connectionstrings.com/
'OLEDB connection string to Access's Jet DB

     strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & sDbPathName & ";" & _
                    "User Id=admin;" & _
                    "Password="

End Function

Public Function StrConnectCurve()
'This connects to the the Gas Curve Spread sheet
'http://www.connectionstrings.com/

    StrConnectCurve = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                      "Data Source=S:\PPM\Menconi\Storage\Gas_Curves\Access Gas Curve.xls;" & _
                      "Extended Properties=""Excel 8.0;HDR=Yes"""

End Function
Public Function sConn_Oracle_ZaiNetEOD() As String

    Dim File_Path_FILEDSN As String
    
    File_Path_FILEDSN = "\\porfiler02\shared\PPM\Asset Management - Wind\PhilS_AssetManagementWind\Development\Database\FILEDSN_Oracle_ZNP01\"
    
''Note  **** the TNS service Name is >>>  ZNP01  <<<
'
               sConn_Oracle_ZaiNetEOD = "FILEDSN=" & File_Path_FILEDSN & "Oracle_ZNP01.dsn;" & _
                                "User Id=p20441;" & _
                                "Password=nike1234;"
                                
                     Debug.Print sConn_Oracle_ZaiNetEOD

'
'       sConn_Oracle_ZaiNetEOD = "DSN=My_Data_Name;" & _
'                                "User Id=p20441;" & _
'                                "Password=nike1234;"
                                
       
       
''DNS-Less connection
'        sConn_Oracle_ZaiNetEOD = "Driver={Oracle in OraHome92};" & _
'                                 "Server=ZNP01;" & _
'                                 "Uid=p20441;" & _
'                                 "Pwd=nike1234;"
                                
                                
 '        sConn_Oracle_ZaiNetEOD = "Provider=OraOLEDB.Oracle;" & _
                                 "Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP))) " & _
                                 "(CONNECT_DATA=(SID=MyOracleSID)(SERVER=DEDICATED)));" & _
                                 "(HOST=myHost)(PORT=myPort)))" & _
                                 "Server=ZNP01;" & _
                                 "User Id=p20441;" & _
                                 "Password=nike1234;"

''OraOLEDB
'            sConn_Oracle_ZaiNetEOD = "Provider=OraOLEDB.Oracle;" & _
'                                    "Data Source=ZNP01;" & _
'                                    "User Id=p20441;" & _
'                                    "Password=test1234;"
                                              
End Function
Public Function sWebTrader_OLEDB()
'Note: This is an OLDEDB connection

        sWebTrader_OLEDB = "Provider=sqloledb;" & _
                    "Data Source=208.254.145.33;" & _
                    "Initial Catalog=ppm_operational;" & _
                    "User Id=PPMClient;" & _
                    "Password=P5EPheyu;"


        'Provider=sqloledb;Data Source=208.254.145.33;Initial Catalog=myDataBase;User Id=myUsername;Password=myPassword;
        ' Provider=sqloledb;Data Source=208.254.145.33;Initial Catalog=ppm_operational;Password=P5EPheyu;


End Function

Public Function sWebTrader_ODBC_DSN_Less()
    'This example comes from DataFast consulting webpage
    'http://www.amazecreations.com/datafast/ShowArticle.aspx?File=Articles/odbctutor01.htm
    Dim strConnect As String
    
'    strConnect = "ODBC;DRIVER={SQL Server}" _
'               & ";SERVER=" & strserver _
'               & ";DATABASE=" & strDatabase _
'               & ";UID=" & strUID _
'               & ";PWD=" & strPWD & ";"


    sWebTrader_ODBC_DSN_Less = "ODBC;Driver={SQL Server}" & _
                                ";Server=208.254.145.33;" & _
                                ";Database=ppm_operational" & _
                                ";Uid=PPMClient" & _
                                ";Pwd=P5EPheyu;"
           
End Function

Public Function sPI_OLEDB()

        'sWebTrader_ODBC_DSN_Less = "ODBC;Driver={SQL Server}" & _
                                ";Server=208.254.145.33;" & _
                                ";Database=ppm_operational" & _
                                ";Uid=PPMClient" & _
                                ";Pwd=P5EPheyu;"


            sPI_OLEDB = "Provider = PIOLEDB" & vbCrLf & _
                      "; Data Source = 172.16.54.162" & _
                      "; User ID =p20441" & _
                      "; Password=Phil!" & _
                      "; Timestamp Interval Start = True" & _
                      "; Keep Default Ordering = False" & _
                      "; Time Zone = Server"

End Function