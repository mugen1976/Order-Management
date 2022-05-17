VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEditClient 
   Caption         =   "UserForm1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10340
   OleObjectBlob   =   "FormEditClient.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "FormEditClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'저장
Private Sub btnSave_Click()
    If Me.Caption = "매출처 등록" Then
        Call Add_Client
    Else
        Call Edit_Client
    End If
    
    Unload Me
End Sub

'닫기
Private Sub btnClose_Click()
    Unload Me
End Sub

'등록
Private Sub Add_Client()

    If Me.txtclientName.Value = "" Then
        MsgBox "매출처명을 입력해 주세요.", vbCritical, "입력오류"
        Exit Sub
    End If
    
    Call Connect_DB
    
    '중복 매입처 체크
    SQL = "SELECT * FROM client WHERE clientName LIKE '" & Me.txtclientName.Value & "'"
    
    rs.CursorLocation = adUseClient '★★★★★★★★★★★★★★★RecordCount를 뽑아내기위해 반드시 필요함
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        MsgBox "이미 등록되어있는 매출처 입니다.", vbCritical, "오류"
        rs.Close
        Cn.Close
        Exit Sub
    Else
        rs.Close
    End If
    
    '매입처 추가 SQL문
    SQL = "INSERT INTO client (clientName, licenseNumber, address, businessConditions, businessCategory)"
    SQL = SQL + " VALUES ('" & Me.txtclientName.Value & "', '" & Me.txtlicenseNumber.Value & "', '" & Me.txtaddress.Value & "', '" & Me.txtbusinessConditions.Value & "', '" & Me.txtbusinessCategory.Value & "')"
    rs.Open SQL, Cn
    Cn.Close

End Sub

'수정
Private Sub Edit_Client()

    Call Connect_DB
    
    '매입처 수정 SQL문
    SQL = "UPDATE client SET clientName = '" & Me.txtclientName.Value & "'"
    SQL = SQL + ", licenseNumber = '" & Me.txtlicenseNumber.Value & "'"
    SQL = SQL + ", address = '" & Me.txtaddress.Value & "'"
    SQL = SQL + ", businessConditions = '" & Me.txtbusinessConditions.Value & "'"
    SQL = SQL + ", businessCategory = '" & Me.txtbusinessCategory.Value & "'"
    SQL = SQL + " WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    Cn.Close

End Sub
'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnSave_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnSave
End Sub

Private Sub btnSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnSave
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnClose
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnClose
End Sub

'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnSave, btnClose" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
 For Each vList In vLists
 If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
 Next
Next
End Sub
'커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
Private Sub OnHover_Css(lbl As Control): With lbl: .BackColor = RGB(211, 240, 224): .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl: .BackColor = &H8000000E: .BorderColor = -2147483638: End With: End Sub



