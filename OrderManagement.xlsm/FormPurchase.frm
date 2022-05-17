VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPurchase 
   Caption         =   "매입처관리"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12260
   OleObjectBlob   =   "FormPurchase.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "FormPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'폼 열릴때 실행
Private Sub UserForm_Initialize()
    
    Call SetColumnHeaders
    Call Connect_DB
    Call Load_Purchase_DB
    Cn.Close
    
    Me.txtSearch.SetFocus

End Sub

'리스트뷰 헤더 설정
Private Sub SetColumnHeaders()
    
    With Me.ListProduct.ColumnHeaders
        .Add Text:="번호", Width:=0, Alignment:=lvwColumnLeft
        .Add Text:="매입처명", Width:=190, Alignment:=lvwColumnLeft
        .Add Text:="구분", Width:=80, Alignment:=lvwColumnCenter
        .Add Text:="발주구분", Width:=80, Alignment:=lvwColumnCenter
    End With

End Sub

'DB 불러오기
Private Sub Load_Purchase_DB(Optional SerchWord As String)
    Dim i, j As Integer
    Dim LstItem As ListItem
    
    Me.ListProduct.ListItems.Clear
    
    '매입처 검색
    SQL = "SELECT idx, purchaseName, sortation, orderSortation FROM purchase"
    
    If SerchWord <> "" Then
        SQL = SQL + " WHERE purchaseName LIKE '%" + SerchWord + "%'"
    End If
    
    SQL = SQL + " ORDER BY purchaseName"

    rs.CursorLocation = adUseClient '★★★★★★★★★★★★★★★RecordCount를 뽑아내기위해 반드시 필요함
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly

    '자료가 없을경우 종료
    If rs.RecordCount = 0 Then GoTo ex:
    
    With Me.ListProduct
        rs.MoveFirst
        For j = 1 To rs.RecordCount   '레코드 수만큼 입력
            Set LstItem = .ListItems.Add(, , CStr(rs.Fields(0).Value))
            For i = 1 To 3
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
Private Sub ListProduct_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With ListProduct.SelectedItem
        Me.txtIdx = .Text
        Me.txtPurchase = .SubItems(1)
        Me.txtSortation = .SubItems(2)
        Me.txtOrdersort = .SubItems(3)
    End With
End Sub

'listview 머리글 영역 클릭시 실행
Private Sub ListProduct_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListProduct
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

'검색창
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '엔터키를 눌렀을때
    If KeyCode = vbKeyReturn Then
        If Me.txtSearch.Value = "" Then
            MsgBox "검색어를 입력해주세요.", vbCritical
            Exit Sub
        End If
        
        Call Connect_DB
        Load_Purchase_DB (Me.txtSearch.Value)
        Cn.Close
    End If
End Sub

'등록
Private Sub btnRegister_Click()

    If txtPurchase.Value = "" Then
        MsgBox "매입처명을 입력해 주세요.", vbCritical, "입력오류"
        Exit Sub
    End If
    
    Call Connect_DB
    
    '중복 매입처 체크
    SQL = "SELECT * FROM purchase WHERE purchaseName LIKE '" & txtPurchase.Value & "'"
    
    rs.CursorLocation = adUseClient '★★★★★★★★★★★★★★★RecordCount를 뽑아내기위해 반드시 필요함
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        MsgBox "이미 등록되어있는 매입처 입니다.", vbCritical, "오류"
        rs.Close
        Cn.Close
        Exit Sub
    Else
        rs.Close
    End If
    
    '매입처 추가 SQL문
    SQL = "INSERT INTO purchase (purchaseName, sortation, orderSortation) VALUES ('" & txtPurchase.Value & "', '" & txtSortation.Value & "', '" & txtOrdersort.Value & "')"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close

End Sub

'수정
Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Call Connect_DB
    
    '매입처 수정 SQL문
    SQL = "UPDATE purchase SET purchaseName = '" & Me.txtPurchase.Value & "'"
    SQL = SQL + ", sortation = '" & Me.txtSortation.Value & "'"
    SQL = SQL + ", orderSortation = '" & Me.txtOrdersort.Value & "'"
    SQL = SQL + " WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close

End Sub

'삭제
Private Sub btnDelete_Click()
    Dim YN As VbMsgBoxResult
    
    YN = MsgBox("선택하신 매입처 정보를 삭제하시겠습니까?", vbYesNo)
    If YN = vbNo Then Exit Sub

    Call Connect_DB
    
    SQL = " DELETE FROM purchase WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close
    
    MsgBox "선택하신 품명 정보가 삭제되었습니다..", vbInformation
    Call reset_Textbox

End Sub

'초기화
Private Sub btnReset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call reset_Textbox
End Sub

'닫기
Private Sub btnClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Unload Me
End Sub

Private Sub reset_Textbox()
    With Me
        .txtIdx.Value = ""
        .txtPurchase = ""
        .txtSortation = ""
        .txtOrdersort = ""
    End With
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
Private Sub btnReset_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnReset
End Sub

Private Sub btnReset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnReset
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
Dim btnList As String: btnList = "btnReset, btnDelete, btnEdit, btnClose, btnRegister" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
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

