VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLogin 
   Caption         =   "로그인"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5450
   OleObjectBlob   =   "FormLogin.frx":0000
   StartUpPosition =   1  '소유자 가운데
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

    If Me.txtID.Value = "아이디를 입력하세요" Then
        Me.txtID.Value = ""
        Me.txtID.ForeColor = RGB(51, 51, 51) ' HEX: #33333
    End If
    
End Sub

Private Sub txtID_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.txtID.Value = "" Then
        Me.txtID.Value = "아이디를 입력하세요"
        Me.txtID.ForeColor = RGB(128, 128, 128) ' HEX: #808080
    End If

End Sub

Private Sub txtPW_Enter()

    If Me.txtPW.Value = "비밀번호를 입력하세요" Then
        Me.txtPW.Value = ""
        Me.txtPW.ForeColor = RGB(51, 51, 51) ' HEX: #33333
        Me.txtPW.PasswordChar = "*"
    End If
    
End Sub

Private Sub txtPW_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '엔터키 눌렀을때
    If KeyCode = vbKeyReturn Then
        Call Login_Verification
    End If
End Sub

Private Sub txtPW_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.txtPW.Value = "" Then
        Me.txtPW.Value = "비밀번호를 입력하세요"
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
    
    ' 오류처리
    If myUser = "아이디를 입력하세요" Or myUser = "" Then MsgBox "아이디를 입력해주세요.": Exit Sub
    If myPass = "비밀번호를 입력하세요" Or myPass = "" Then MsgBox "비밀번호를 입력해주세요.": Exit Sub


    Call Connect_DB
    Call Connect_Table("user", myUser)   '테이블명, username

    rs.Find ("username LIKE '" & myUser & "'")
    If rs.EOF Then   '//찾는값이 없을경우
        MsgBox "등록된 아이디가 없습니다.                  ", vbInformation
        Me.txtID.SetFocus
        Exit Sub
    End If
    
    rs.Find ("password LIKE '" & myPass & "'")
    If rs.EOF Then   '//찾는값이 없을경우
        MsgBox "비밀번호가 틀립니다.                  ", vbInformation
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


