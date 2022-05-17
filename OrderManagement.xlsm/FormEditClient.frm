VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEditClient 
   Caption         =   "UserForm1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10340
   OleObjectBlob   =   "FormEditClient.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "FormEditClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'����
Private Sub btnSave_Click()
    If Me.Caption = "����ó ���" Then
        Call Add_Client
    Else
        Call Edit_Client
    End If
    
    Unload Me
End Sub

'�ݱ�
Private Sub btnClose_Click()
    Unload Me
End Sub

'���
Private Sub Add_Client()

    If Me.txtclientName.Value = "" Then
        MsgBox "����ó���� �Է��� �ּ���.", vbCritical, "�Է¿���"
        Exit Sub
    End If
    
    Call Connect_DB
    
    '�ߺ� ����ó üũ
    SQL = "SELECT * FROM client WHERE clientName LIKE '" & Me.txtclientName.Value & "'"
    
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
    SQL = "INSERT INTO client (clientName, licenseNumber, address, businessConditions, businessCategory)"
    SQL = SQL + " VALUES ('" & Me.txtclientName.Value & "', '" & Me.txtlicenseNumber.Value & "', '" & Me.txtaddress.Value & "', '" & Me.txtbusinessConditions.Value & "', '" & Me.txtbusinessCategory.Value & "')"
    rs.Open SQL, Cn
    Cn.Close

End Sub

'����
Private Sub Edit_Client()

    Call Connect_DB
    
    '����ó ���� SQL��
    SQL = "UPDATE client SET clientName = '" & Me.txtclientName.Value & "'"
    SQL = SQL + ", licenseNumber = '" & Me.txtlicenseNumber.Value & "'"
    SQL = SQL + ", address = '" & Me.txtaddress.Value & "'"
    SQL = SQL + ", businessConditions = '" & Me.txtbusinessConditions.Value & "'"
    SQL = SQL + ", businessCategory = '" & Me.txtbusinessCategory.Value & "'"
    SQL = SQL + " WHERE idx = '" & Me.txtIdx.Value & "'"
    rs.Open SQL, Cn
    Cn.Close

End Sub
'�������� �߰��� ��ư�� ������ŭ �Ʒ� ��ɹ��� �������� �߰��� ��, btnClose �� ��ư �̸����� �����մϴ�.
Private Sub btnSave_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnSave
End Sub

Private Sub btnSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnSave
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
Dim btnList As String: btnList = "btnSave, btnClose" ' ��ư �̸��� ��ǥ�� �����Ͽ� �Է��ϼ���.
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



