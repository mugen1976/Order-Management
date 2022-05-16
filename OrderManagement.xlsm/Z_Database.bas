Attribute VB_Name = "Z_Database"
Option Explicit

Public Cn  As ADODB.Connection
Public rs  As ADODB.Recordset
Public SQL As String
Dim SERVER$, DBNAME$, USERNAME$, PASSWORD$  'As String

Public Sub Connect_DB()
    
    SERVER = "서버주소"         'IP 어드레스(서버주소)
    DBNAME = "DB이름"           'DB
    USERNAME = "DB아이디"       'ID
    PASSWORD = "DB패스워드"     'PW
    
    Set Cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'MySQL
    Call ConnectMySQL
    'MSSQL
'    Call ConnectMSSQL
    Cn.Open
End Sub

Private Sub ConnectMSSQL()
    
    '// 방식1 : DSN 이용안함, 대신 Server명, Port등을 지정해 주어야 한다.
    Cn.ConnectionString = "Provider = SQLOLEDB;Data Source = " & SERVER & ";Initial Catalog = " & DBNAME & ";User ID = " & USERNAME & ";Password =" & PASSWORD & ";"
    
    '// 방식2 : DSN을 이용하여 연결(상세한 연결Option은 DSN을 만들때 지정할 수 있다)
    'Cn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & a & "; Initial Catalog=" & b & ";Integrated Security=SSPI;"
    
End Sub

Public Sub ConnectMySQL()
    
    '// 방식1 : DSN을 이용하여 연결(상세한 연결Option은 DSN을 만들때 지정할 수 있다)
    'Cn.ConnectionString = "DSN=" & SERVER & ";Uid=" & USERNAME & ";Pwd=" & PASSWORD & ";Option=2;"
    
    '// 방식2 : DSN 이용안함, 대신 Server명, Port등을 지정해 주어야 한다.
    Cn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};Server=" & SERVER & ";Port=3306;Database=" & DBNAME & ";User=" & USERNAME & ";Password=" & PASSWORD & ";Option=2;"

End Sub
