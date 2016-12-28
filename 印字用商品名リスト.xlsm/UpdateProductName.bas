Attribute VB_Name = "UpdateProductName"
Sub Update()

'�X�V���f�[�^�̃u�b�N�E�V�[�g���w��
Dim SourceSheet As Worksheet
Set SourceSheet = Workbooks(1).Worksheets(1)

Worksheets("�ŏI").Activate

'�u�ŏI�v�V�[�g�Ɋ��ɓ����Ă��鏤�i�R�[�h�̃����W
Dim CodeRange As Range
Set CodeRange = Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1))

Dim i As Long, Code As String, Name As String
i = 2

'�J���Ă��郏�[�N�u�b�N�̓�ڂȂ̂ŁA����w�肵�����K�v������B
Do
    Code = SourceSheet.Cells(i, 1).Value

    Dim HitRow As Long
    HitRow = WorksheetFunction.Match(Code, CodeRange, 0)
    
    Cells(HitRow, 2).Value = SourceSheet.Cells(i, 4).Value
    
    i = i + 1
    
Loop Until IsEmpty(SourceSheet.Cells(i, 1).Value)

End Sub

