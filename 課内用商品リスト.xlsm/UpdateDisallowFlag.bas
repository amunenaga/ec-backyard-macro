Attribute VB_Name = "UpdateDisallowFlag"
Option Explicit
Sub �s���R�]�L()

Dim Arr As Variant
Arr = Array("�p��", "�s��", "�s��")

'JAN�ŁA�s���R���󗓂��܂��t�B���^�[
With Worksheets("���i���").Range("A1").CurrentRegion
    '.AutoFilter Field:=2, Criteria1:="???????*"
    .AutoFilter Field:=16, Criteria1:="="
End With

'�ݒ肵���������ꂼ��ŁA�t�B���^�[���Ď�z�s���R�֋L�������s
Dim s As Variant
For Each s In Arr
    Call InputReason(s)
Next

Worksheets("���i���").Range("A1").AutoFilter


End Sub

Sub InputReason(ByVal Str As String)
Attribute InputReason.VB_ProcData.VB_Invoke_Func = " \n14"
'�p�ԃt���O�̖��]�L���t�B���^�[

Worksheets("���i���").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:="*" & Str & "*"

Dim r As Range, TargetRange As Range
Set TargetRange = Intersect(Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible), Range("P2:P300000"))

If TargetRange Is Nothing Then Exit Sub

For Each r In TargetRange
    r.Offset(0, -1).Value = 1
    r.Offset(0, 1).Value = Date
    r.Value = Str
Next

End Sub
