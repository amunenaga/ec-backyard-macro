Attribute VB_Name = "UpdateVendorCode"
Option Explicit

Sub �d����R�[�h�ǋL()

With Worksheets("���i���")
    .Activate
    
    Dim EndRow As Long
    EndRow = .UsedRange.Rows.Count

End With

Dim i As Long

For i = 2 To EndRow

    Call UpdateVendorCode(i)

Continue:
Next

End Sub

Sub UpdateVendorCode(ByVal CurrentRow As Double)

Dim VenderTerm As String, Terms As Variant, v As Variant
VenderTerm = Cells(CurrentRow, 4).Value

'�d���於��Split
Terms = Split(VenderTerm, " ")

'��������������z��A�e�X�ɂ��Ďd����V�[�g�ƈ�v����d���於�����邩
For Each v In Terms

    Dim TmpVendorCode As String
    TmpVendorCode = GetVendorCode(CStr(v))
    
    '�d����R�[�h���擾�ł��Ȃ���Ύ��̗v�f��
    If TmpVendorCode <> "" Then
    
        Cells(CurrentRow, 32).NumberFormatLocal = "@"
        Cells(CurrentRow, 32).Value = TmpVendorCode
        
        '���L�t�B�[���h���X�V���āA�ۗ��t���O�𗧂ĂĂ���
        Dim VendorName As String, ExclamationNote As String
        VendorName = v
        
        '�������̒��ӏ��������́A�d���悩��d���於�̂��폜�����c��̕���
        ExclamationNote = Trim(Replace(VenderTerm, VendorName, ""))
        
        If ExclamationNote <> "" Then
            If IsEmpty(Cells(CurrentRow, 35).Value) Then Cells(CurrentRow, 35).Value = ExclamationNote
            Cells(CurrentRow, 34).Value = 1
        End If

    End If

Next

End Sub

Function GetVendorCode(ByVal VendorTerm As String)

On Error Resume Next
    
Dim VendorCode As String
VendorCode = WorksheetFunction.VLookup(VendorTerm, Worksheets("�d����").Range("B2:AA500"), 26, 0)

GetVendorCode = VendorCode

End Function
