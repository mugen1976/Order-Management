VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPurchase 
   Caption         =   "����ó����"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12260
   OleObjectBlob   =   "FormPurchase.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "FormPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�� ������ ����
Private Sub UserForm_Initialize()
    
    Call SetColumnHeaders
    Call Connect_DB
    Call Load_Purchase_DB
    Cn.Close
    
    Me.txtSearch.SetFocus

End Sub

'����Ʈ�� ��� ����
Private Sub SetColumnHeaders()
    
    With Me.ListProduct.ColumnHeaders
        .Add Text:="��ȣ", Width:=0, Alignment:=lvwColumnLeft
        .Add Text:="����ó��", Width:=190, Alignment:=lvwColumnLeft
        .Add Text:="����", Width:=80, Alignment:=lvwColumnCenter
        .Add Text:="���ֱ���", Width:=80, Alignment:=lvwColumnCenter
    End With

End Sub

'DB �ҷ�����
Private Sub Load_Purchase_DB(Optional SerchWord As String)
    Dim i, j As Integer
    Dim LstItem As ListItem
    
    Me.ListProduct.ListItems.Clear
    
    '����ó �˻�
    SQL = "SELECT idx, purchaseName, sortation, orderSortation FROM purchase"
    
    If SerchWord <> "" Then
        SQL = SQL + " WHERE purchaseName LIKE '%" + SerchWord + "%'"
    End If
    
    SQL = SQL + " ORDER BY purchaseName"

    rs.CursorLocation = adUseClient '�ڡڡڡڡڡڡڡڡڡڡڡڡڡڡ�RecordCount�� �̾Ƴ������� �ݵ�� �ʿ���
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly

    '�ڷᰡ ������� ����
    If rs.RecordCount = 0 Then GoTo ex:
    
    With Me.ListProduct
        rs.MoveFirst
        For j = 1 To rs.RecordCount   '���ڵ� ����ŭ �Է�
            Set LstItem = .ListItems.Add(, , CStr(rs.Fields(0).Value))
            For i = 1 To 3
                If Not IsNull(rs.Fields(i)) Then
                    LstItem.SubItems(i) = CStr(rs.Fields(i).Value)
                End If
            Next i
            rs.MoveNext    '�������ڵ�� �̵�
        Next j
    End With

ex:
    rs.Close
End Sub

'listview ������ ����
Private Sub ListProduct_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With ListProduct.SelectedItem
        Me.txtIdx = .Text
        Me.txtPurchase = .SubItems(1)
        Me.txtSortation = .SubItems(2)
        Me.txtOrdersort = .SubItems(3)
    End With
End Sub

'listview �Ӹ��� ���� Ŭ���� ����
Private Sub ListProduct_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListProduct
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

'�˻�â
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '����Ű�� ��������
    If KeyCode = vbKeyReturn Then
        If Me.txtSearch.Value = "" Then
            MsgBox "�˻�� �Է����ּ���.", vbCritical
            Exit Sub
        End If
        
        Call Connect_DB
        Load_Purchase_DB (Me.txtSearch.Value)
        Cn.Close
    End If
End Sub

'���
Private Sub btnRegister_Click()

    If txtPurchase.Value = "" Then
        MsgBox "����ó���� �Է��� �ּ���.", vbCritical, "�Է¿���"
        Exit Sub
    End If
    
    Call Connect_DB
    
    '�ߺ� ����ó üũ
    SQL = "SELECT * FROM purchase WHERE purchaseName LIKE '" & txtPurchase.Value & "'"
    
    rs.CursorLocation = adUseClient '�ڡڡڡڡڡڡڡڡڡڡڡڡڡڡ�RecordCount�� �̾Ƴ������� �ݵ�� �ʿ���
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        MsgBox "�̹� ��ϵǾ��ִ� ����ó �Դϴ�.", vbCritical, "����"
        rs.Close
        Cn.Close
        Exit Sub
    Else
        rs.Close
    End If
    
    '����ó �߰� SQL��
    SQL = "INSERT INTO purchase (purchaseName, sortation, orderSortation) VALUES ('" & txtPurchase.Value & "', '" & txtSortation.Value & "', '" & txtOrdersort.Value & "')"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close

End Sub

'����
Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Call Connect_DB
    
    '����ó ���� SQL��
    SQL = "UPDATE purchase SET purchaseName = '" & Me.txtPurchase.Value & "'"
    SQL = SQL + ", sortation = '" & Me.txtSortation.Value & "'"
    SQL = SQL + ", orderSortation = '" & Me.txtOrdersort.Value & "'"
    SQL = SQL + " WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close

End Sub

'����
Private Sub btnDelete_Click()
    Dim YN As VbMsgBoxResult
    
    YN = MsgBox("�����Ͻ� ����ó ������ �����Ͻðڽ��ϱ�?", vbYesNo)
    If YN = vbNo Then Exit Sub

    Call Connect_DB
    
    SQL = " DELETE FROM purchase WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Purchase_DB
    Cn.Close
    
    MsgBox "�����Ͻ� ǰ�� ������ �����Ǿ����ϴ�..", vbInformation
    Call reset_Textbox

End Sub

'�ʱ�ȭ
Private Sub btnReset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call reset_Textbox
End Sub

'�ݱ�
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

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnRegister_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnRegister
End Sub

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEdit
End Sub

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnDelete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnDelete
End Sub

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnReset_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnReset
End Sub

Private Sub btnReset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnReset
End Sub

'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnClose
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnClose
End Sub

'�Ʒ� �ڵ带 �������� �߰��� ��, "btnXXX, btnYYY"�� ��ư�̸��� ��ǥ�� ������ ������ �����մϴ�.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnReset, btnDelete, btnEdit, btnClose, btnRegister" ' ��ư �̸��� ��ǥ�� �����Ͽ� �Է��ϼ���.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
 For Each vList In vLists
 If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
 Next
Next
End Sub
'Ŀ�� �̵��� ��ư ������ �����ϴ� ������ɹ��� �������� �߰��մϴ�.
Private Sub OnHover_Css(lbl As Control): With lbl: .BackColor = RGB(211, 240, 224): .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl: .BackColor = &H8000000E: .BorderColor = -2147483638: End With: End Sub

