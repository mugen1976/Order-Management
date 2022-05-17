VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormClient 
   Caption         =   "매출처 관리"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13410
   OleObjectBlob   =   "FormClient.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "FormClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'폼 열릴때 실행
Private Sub UserForm_Initialize()
    
    Call SetColumnHeaders
    Call Connect_DB
    Call Load_Client_DB
    Cn.Close
    
    Me.txtSearch.SetFocus

End Sub

'리스트뷰 헤더 설정
Private Sub SetColumnHeaders()
    
    With Me.ListClient.ColumnHeaders
        .Add Text:="번호", Width:=0, Alignment:=lvwColumnLeft
        .Add Text:="매출처명", Width:=150, Alignment:=lvwColumnCenter
        .Add Text:="사업자등록번호", Width:=100, Alignment:=lvwColumnCenter
        .Add Text:="주소", Width:=200, Alignment:=lvwColumnCenter
        .Add Text:="업태", Width:=80, Alignment:=lvwColumnCenter
        .Add Text:="종목", Width:=80, Alignment:=lvwColumnCenter
    End With

End Sub

'DB 불러오기
Private Sub Load_Client_DB(Optional SerchWord As String)
    Dim i, j As Integer
    Dim LstItem As ListItem
    
    Me.ListClient.ListItems.Clear
    
    '매입처 검색
    SQL = "SELECT idx, clientName, licenseNumber, address, businessConditions, businessCategory FROM client"
    
    If SerchWord <> "" Then
        SQL = SQL + " WHERE clientName LIKE '%" + SerchWord + "%'"
    End If
    
    SQL = SQL + " ORDER BY clientName"

    rs.CursorLocation = adUseClient '★★★★★★★★★★★★★★★RecordCount를 뽑아내기위해 반드시 필요함
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly

    '자료가 없을경우 종료
    If rs.RecordCount = 0 Then GoTo ex:
    
    With Me.ListClient
        rs.MoveFirst
        For j = 1 To rs.RecordCount   '레코드 수만큼 입력
            Set LstItem = .ListItems.Add(, , CStr(rs.Fields(0).Value))
            For i = 1 To 5
                If Not IsNull(rs.Fields(i)) Then
                    LstItem.SubItems(i) = CStr(rs.Fields(i).Value)
                End If
            Next i
            rs.MoveNext    '다음레코드로 이동
        Next j
    End With

ex:
    rs.Close
End Sub

'listview 아이템 선택
Private Sub ListClient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtIdx = ListClient.SelectedItem.Text
End Sub

'등록
Private Sub btnRegister_Click()
    With FormEditClient
        .Caption = "매출처 등록"
        .Show
    End With
    
    Call Connect_DB
    Call Load_Client_DB
    Cn.Close
End Sub

'수정
Private Sub btnEdit_Click()
    If Me.txtIdx = "" Then
        MsgBox "수정할 매출처를 선택해 주세요.", vbCritical, "오류"
        Exit Sub
    End If
    
    With FormEditClient
        .Caption = "매출처 수정"
        .txtIdx = ListClient.SelectedItem.Text
        .txtclientName = ListClient.SelectedItem.SubItems(1)
        .txtaddress = ListClient.SelectedItem.SubItems(2)
        .txtbusinessCategory = ListClient.SelectedItem.SubItems(3)
        .txtbusinessConditions = ListClient.SelectedItem.SubItems(4)
        .Show
    End With
    
    Call Connect_DB
    Call Load_Client_DB
    Cn.Close
End Sub

'삭제
Private Sub btnDelete_Click()

    If Me.txtIdx = "" Then
        MsgBox "삭제할 매출처를 선택해 주세요.", vbCritical, "오류"
        Exit Sub
    End If

    Dim YN As VbMsgBoxResult
    
    YN = MsgBox("선택하신 매입처 정보를 삭제하시겠습니까?", vbYesNo)
    If YN = vbNo Then Exit Sub

    Call Connect_DB
    
    SQL = " DELETE FROM client WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Client_DB
    Cn.Close
    
    Me.txtIdx = ""
    MsgBox "선택하신 품명 정보가 삭제되었습니다..", vbInformation

End Sub

'닫기
Private Sub btnClose_Click()
    Unload Me
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnRegister_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnRegister
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEdit
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnDelete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnDelete
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
Dim btnList As String: btnList = "btnDelete, btnEdit, btnClose, btnRegister" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
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


