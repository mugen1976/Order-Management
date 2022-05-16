VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLogin 
   Caption         =   "�α���"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5450
   OleObjectBlob   =   "FormLogin.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

CancelButtonInactive.Visible = True
OKButtonInactive.Visible = True

End Sub

Sub CancelButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button appear Green when hovered on

CancelButtonInactive.Visible = False
OKButtonInactive.Visible = True

End Sub

Sub OKButtonInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make OK Button appear Green when hovered on

  CancelButtonInactive.Visible = True
  OKButtonInactive.Visible = False

End Sub

Private Sub txtID_Enter()

    If Me.txtID.Value = "���̵� �Է��ϼ���" Then
        Me.txtID.Value = ""
        Me.txtID.ForeColor = RGB(51, 51, 51) ' HEX: #33333
    End If
    
End Sub

Private Sub txtID_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.txtID.Value = "" Then
        Me.txtID.Value = "���̵� �Է��ϼ���"
        Me.txtID.ForeColor = RGB(128, 128, 128) ' HEX: #808080
    End If

End Sub

Private Sub txtPW_Enter()

    If Me.txtPW.Value = "��й�ȣ�� �Է��ϼ���" Then
        Me.txtPW.Value = ""
        Me.txtPW.ForeColor = RGB(51, 51, 51) ' HEX: #33333
        Me.txtPW.PasswordChar = "*"
    End If
    
End Sub

Private Sub txtPW_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '����Ű ��������
    If KeyCode = vbKeyReturn Then
        Call Login_Verification
    End If
End Sub

Private Sub txtPW_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.txtPW.Value = "" Then
        Me.txtPW.Value = "��й�ȣ�� �Է��ϼ���"
        Me.txtPW.ForeColor = RGB(128, 128, 128) ' HEX: #808080
        Me.txtPW.PasswordChar = ""
    End If

End Sub

Private Sub OkButton_Click()

    Call Login_Verification
    
End Sub

Private Sub CancelButton_Click()
    
    Unload Me

End Sub

Sub Login_Verification()

    Dim myUser As String, myPass As String
    Dim myPerm As String
    
    myUser = Me.txtID.Value
    myPass = Me.txtPW.Value
    
    ' ����ó��
    If myUser = "���̵� �Է��ϼ���" Or myUser = "" Then MsgBox "���̵� �Է����ּ���.": Exit Sub
    If myPass = "��й�ȣ�� �Է��ϼ���" Or myPass = "" Then MsgBox "��й�ȣ�� �Է����ּ���.": Exit Sub


    Call Connect_DB
    Call Connect_Table("user", myUser)   '���̺��, username

    rs.Find ("username LIKE '" & myUser & "'")
    If rs.EOF Then   '//ã�°��� �������
        MsgBox "��ϵ� ���̵� �����ϴ�.                  ", vbInformation
        Me.txtID.SetFocus
        Exit Sub
    End If
    
    rs.Find ("password LIKE '" & myPass & "'")
    If rs.EOF Then   '//ã�°��� �������
        MsgBox "��й�ȣ�� Ʋ���ϴ�.                  ", vbInformation
        Me.txtPW.SetFocus
        Exit Sub
    End If
    
    USERNAME = rs.Fields("username")
    rs.Close
    
    Unload Me
    
End Sub

Sub Connect_Table(tblName As String, sqlWhere As String)
    Set rs = Nothing
    Set rs = New Recordset
    
    If sqlWhere <> "" Then
        tblName = tblName & " WHERE username LIKE '" & sqlWhere & "'"
    End If
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM " & tblName, Cn, adOpenStatic, adLockOptimistic
End Sub


