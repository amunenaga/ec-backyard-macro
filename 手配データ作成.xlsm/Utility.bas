Attribute VB_Name = "Utility"
Option Explicit

Sub PrepareSheet(ByRef TargetSheet As Variant)

If Not IsEmpty(TargetSheet.Range("A2").Value) Then
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 0).Delete Shift:=xlShiftUp
End If

End Sub

Function FetchWorkBook(path As String) As Workbook

'�����œn���ꂽ�p�X�̃u�b�N���J���܂��B�����̃u�b�N���J���Ă���΁A���̃u�b�N��߂�l�ɂ��܂��B

Dim WorkBookName As String
WorkBookName = Dir(path)

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = WorkBookName Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(path)

ret:
Set FetchWorkBook = wb

End Function

