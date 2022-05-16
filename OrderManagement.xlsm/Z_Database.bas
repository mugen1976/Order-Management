Attribute VB_Name = "Z_Database"
Option Explicit

Public Cn  As ADODB.Connection
Public rs  As ADODB.Recordset
Public SQL As String
Dim SERVER$, DBNAME$, USERNAME$, PASSWORD$  'As String

Public Sub Connect_DB()
    
    SERVER = "�����ּ�"         'IP ��巹��(�����ּ�)
    DBNAME = "DB�̸�"           'DB
    USERNAME = "DB���̵�"       'ID
    PASSWORD = "DB�н�����"     'PW
    
    Set Cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'MySQL
    Call ConnectMySQL
    'MSSQL
'    Call ConnectMSSQL
    Cn.Open
End Sub

Private Sub ConnectMSSQL()
    
    '// ���1 : DSN �̿����, ��� Server��, Port���� ������ �־�� �Ѵ�.
    Cn.ConnectionString = "Provider = SQLOLEDB;Data Source = " & SERVER & ";Initial Catalog = " & DBNAME & ";User ID = " & USERNAME & ";Password =" & PASSWORD & ";"
    
    '// ���2 : DSN�� �̿��Ͽ� ����(���� ����Option�� DSN�� ���鶧 ������ �� �ִ�)
    'Cn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & a & "; Initial Catalog=" & b & ";Integrated Security=SSPI;"
    
End Sub

Public Sub ConnectMySQL()
    
    '// ���1 : DSN�� �̿��Ͽ� ����(���� ����Option�� DSN�� ���鶧 ������ �� �ִ�)
    'Cn.ConnectionString = "DSN=" & SERVER & ";Uid=" & USERNAME & ";Pwd=" & PASSWORD & ";Option=2;"
    
    '// ���2 : DSN �̿����, ��� Server��, Port���� ������ �־�� �Ѵ�.
    Cn.ConnectionString = "DRIVER={MySQL ODBC 5.3 Unicode Driver};Server=" & SERVER & ";Port=3306;Database=" & DBNAME & ";User=" & USERNAME & ";Password=" & PASSWORD & ";Option=2;"

End Sub
