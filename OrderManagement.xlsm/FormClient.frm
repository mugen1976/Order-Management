VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormClient 
   Caption         =   "����ó ����"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13410
   OleObjectBlob   =   "FormClient.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "FormClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'�� ������ ����
Private Sub UserForm_Initialize()
    
    Call SetColumnHeaders
    Call Connect_DB
    Call Load_Client_DB
    Cn.Close
    
    Me.txtSearch.SetFocus

End Sub

'����Ʈ�� ��� ����
Private Sub SetColumnHeaders()
    
    With Me.ListClient.ColumnHeaders
        .Add Text:="��ȣ", Width:=0, Alignment:=lvwColumnLeft
        .Add Text:="����ó��", Width:=150, Alignment:=lvwColumnCenter
        .Add Text:="����ڵ�Ϲ�ȣ", Width:=100, Alignment:=lvwColumnCenter
        .Add Text:="�ּ�", Width:=200, Alignment:=lvwColumnCenter
        .Add Text:="����", Width:=80, Alignment:=lvwColumnCenter
        .Add Text:="����", Width:=80, Alignment:=lvwColumnCenter
    End With

End Sub

'DB �ҷ�����
Private Sub Load_Client_DB(Optional SerchWord As String)
    Dim i, j As Integer
    Dim LstItem As ListItem
    
    Me.ListClient.ListItems.Clear
    
    '����ó �˻�
    SQL = "SELECT idx, clientName, licenseNumber, address, businessConditions, businessCategory FROM client"
    
    If SerchWord <> "" Then
        SQL = SQL + " WHERE clientName LIKE '%" + SerchWord + "%'"
    End If
    
    SQL = SQL + " ORDER BY clientName"

    rs.CursorLocation = adUseClient '�ڡڡڡڡڡڡڡڡڡڡڡڡڡڡ�RecordCount�� �̾Ƴ������� �ݵ�� �ʿ���
    rs.Open SQL, Cn, adOpenStatic, adLockReadOnly

    '�ڷᰡ ������� ����
    If rs.RecordCount = 0 Then GoTo ex:
    
    With Me.ListClient
        rs.MoveFirst
        For j = 1 To rs.RecordCount   '���ڵ� ����ŭ �Է�
            Set LstItem = .ListItems.Add(, , CStr(rs.Fields(0).Value))
            For i = 1 To 5
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
Private Sub ListClient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtIdx = ListClient.SelectedItem.Text
End Sub

'���
Private Sub btnRegister_Click()
    With FormEditClient
        .Caption = "����ó ���"
        .Show
    End With
    
    Call Connect_DB
    Call Load_Client_DB
    Cn.Close
End Sub

'����
Private Sub btnEdit_Click()
    If Me.txtIdx = "" Then
        MsgBox "������ ����ó�� ������ �ּ���.", vbCritical, "����"
        Exit Sub
    End If
    
    With FormEditClient
        .Caption = "����ó ����"
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

'����
Private Sub btnDelete_Click()

    If Me.txtIdx = "" Then
        MsgBox "������ ����ó�� ������ �ּ���.", vbCritical, "����"
        Exit Sub
    End If

    Dim YN As VbMsgBoxResult
    
    YN = MsgBox("�����Ͻ� ����ó ������ �����Ͻðڽ��ϱ�?", vbYesNo)
    If YN = vbNo Then Exit Sub

    Call Connect_DB
    
    SQL = " DELETE FROM client WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    
    Call Load_Client_DB
    Cn.Close
    
    Me.txtIdx = ""
    MsgBox "�����Ͻ� ǰ�� ������ �����Ǿ����ϴ�..", vbInformation

End Sub

'�ݱ�
Private Sub btnClose_Click()
    Unload Me
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
Private Sub btnClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnClose
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnClose
End Sub

'�Ʒ� �ڵ带 �������� �߰��� ��, "btnXXX, btnYYY"�� ��ư�̸��� ��ǥ�� ������ ������ �����մϴ�.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnDelete, btnEdit, btnClose, btnRegister" ' ��ư �̸��� ��ǥ�� �����Ͽ� �Է��ϼ���.
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


