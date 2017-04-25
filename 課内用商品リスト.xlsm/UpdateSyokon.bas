Attribute VB_Name = "UpdateSyokon"
Option Explicit

Type Syokon

    Code As String
    Jan As String
    VendorCode As String
    
End Type

Sub SKU��JAN������6�P�^�Œu������()

'���̃G�N�Z���u�b�N���A���V�[�g�͈͖̔͂���w��̂���
'�C�~�f�B�G�C�g�ŁAWorkbooks(1).name�Ń��[�N�u�b�N�����m�F�ł���B
Dim Rng As Range
Set Rng = Workbooks("6�P�^-JAN���X�g0309.xlsx").Sheets(1).Range("A2:A50200")

Dim r As Range
For Each r In Rng

    'Debug.Assert r.Row < 1000

    Dim sy As Syokon
    
    'ToDo 5�n�܂�R�[�h�́A�擪0�𗎂Ƃ�
    
    sy.Code = r.Value
    sy.Jan = r.Offset(0, 2)
    
    '9�n�܂�A1�n�܂�6�P�^�͎��ށE�Y��̂��ߔ�΂�
    If sy.Code Like "09#####" Or sy.Code Like "01#####" Then
        
        GoTo continue
            
    End If

    Call UpdateJan(sy)
    
continue:

Next

ThisWorkbook.Close SaveChanges:=True

End Sub


Private Sub UpdateJan(Syokon As Syokon)

Dim c As Range

'A��̊Y��JAN��T��
With Workbooks("�����p���i���.xlsm").Worksheets("���i���").Columns(1)

'���S��v��
Set c = .Find(what:=Syokon.Jan, LookIn:=xlValues, LookAt:=xlWhole)

If c Is Nothing Then Exit Sub
If Cells(c.Row, 2).Value Like "######" Then Exit Sub
'�ŏ��̃Z���̃A�h���X���T����
Dim FirstAddress As String
FirstAddress = c.Address

'�J�Ԃ��������A�����𖞂������ׂẴZ������������
Do

    Dim SkuCell As Range, Sku As String
    Set SkuCell = c.Offset(0, 1)
    Sku = SkuCell.Value
    
    '�����ă��[�}���ϐ��ɂ���AF�񌩏o���̌��̈Ӑ}���s���Ȃ̂�
    Dim KijunSku As Range
    Set KijunSku = Cells(c.Row, 6)
    
    'B��Ƀn�C�t�����Ȃ��A6�P�^�ł��Ȃ���΁A6�P�^�ŏ㏑��
    If Not Sku = Syokon.Code And InStr(Sku, "-") < 1 Then
         c.Offset(0, 1).Value = Syokon.Code
    End If
    
    'F��͑S�ď㏑��
    If KijunSku.Value <> Syokon.Code Then
        KijunSku.Value = Syokon.Code
    End If
    
    '���̌������Z�b�g
    Set c = .FindNext(c)
    If c Is Nothing Then Exit Do

Loop Until c.Address = FirstAddress

End With

End Sub

Sub �����̎d����ɍ��킹��()

Dim FinalRow As Long, i As Long
FinalRow = Worksheets("���i���").UsedRange.Rows.Count

For i = 2 To FinalRow

    Call UpdateVendor(i)

Next


End Sub

Private Sub UpdateVendor(ByVal Row As Long)

Dim CurrentVendor As String
CurrentVendor = Cells(Row, 4).Value

Dim CurrentCode As String
CurrentCode = Cells(Row, 2).Value

'�d���於����Ȃ�A���i�}�X�^�Ɋ�Â����d���於������
If CurrentVendor = "" Then
    Dim NewVendorName
    NewVendorName = GetVendorName(GetSyokonVendor(CurrentCode))
    
    If NewVendorName <> "" Then
    
        Cells(Row, 4).Value = NewVendorName
    
    End If
    
    Exit Sub
End If

'��z���쐬-�d����V�[�g�̎d����R�[�h���擾
Dim CurrentVendorCode As String
CurrentVendorCode = GetVendorCodeFromPurBook(CurrentVendor)

'���i�}�X�^�̎d����R�[�h���擾
Dim SyokonVendorCode As String
SyokonVendorCode = GetSyokonVendor(CurrentCode)
If SyokonVendorCode = "" Then Exit Sub

'���i�}�X�^�̎d����R�[�h�ƈ�v���邩�`�F�b�N
'�s��v�Ȃ�΁A�d����V�[�g�̖��̂ŏ㏑������
If CurrentVendorCode <> SyokonVendorCode Then
    Cells(Row, 4).Value = GetVendorName(SyokonVendorCode)
End If

End Sub

Private Function GetVendorCodeFromPurBook(ByVal VendorName As String) As String
On Error Resume Next

GetVendorCodeFromPurBook = WorksheetFunction.VLookup(VendorName, Worksheets("�d����").Range("B2:AA490"), 26, False)

End Function

Private Function GetSyokonVendor(ByVal Code As String) As String
On Error Resume Next

GetSyokonVendor = WorksheetFunction.VLookup(Code, Workbooks("���i�}�X�^.xlsx").Worksheets(1).Range("A2:C11384"), 3, False)

End Function

Private Function GetVendorName(ByVal VendorCode As String) As String
On Error Resume Next

GetVendorName = WorksheetFunction.VLookup(VendorCode, Worksheets("�d����").Range("AA2:AB490"), 2, False)

End Function
