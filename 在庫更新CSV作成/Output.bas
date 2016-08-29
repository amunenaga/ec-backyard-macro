Attribute VB_Name = "Output"
Option Explicit

Sub AppendQtyCsv()

Dim startTime As Long
startTime = Timer

'����
Call FetchSecondInventry
'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Call SetRangeName

'�����f�[�^����S�A�C�e���ɍ݌ɂ��Z�b�g
Call SetQuantity

With yahoo6digit

    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "�I�Ȃ��ɗL", "�I�Ȃ�����" _
            ), Operator:=xlFilterValues
    
    '�t�B���^�[���������W���Z�b�g
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'�����o���pCSV�V�[�g��p��
Worksheets("CSV").Cells.Clear

'�w�b�_�[����������
Dim header As Variant
header = Array("code", "quantity", "allow-overdraft")

Worksheets("CSV").Range("A1:C1") = header

Worksheets("���t�[�f�[�^").Activate

Dim colQuantity As Long, colAllow As Long
colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column

Dim i As Long
i = 2
'�R�[�h�����W�ɑ΂��āAr.row�ōs�ԍ������o���ē����s��Quantity/Allow�̒l���擾����
Dim r As Range
For Each r In CodeRange

    Dim Code As String
    Code = r.Value

    Dim qty As Long, pur As String
    qty = Cells(r.Row, colQuantity).Value
    pur = Cells(r.Row, colAllow).Value

    Worksheets("CSV").Range("A" & i & ":C" & i) = Array(Code, qty, pur)

    i = i + 1

Next

Worksheets("CSV").Activate

'CSV�ǋL����
Dim FSO As New FileSystemObject
Dim Csv As Object

'�ǋL���[�h ForAppending �Ńt�@�C�����J��
Set Csv = FSO.OpenTextFile(Filename:=ThisWorkbook.Path & "\" & "���t�[�݌ɍX�V" & Format(Date, "mmdd") & ".csv", IOMode:=8)

For i = 2 To Worksheets("CSV").UsedRange.Rows.Count
    
    With Worksheets("Csv")
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next

'�I���������i�[
Dim endTime As Long
endTime = Timer

'���O�V�[�g�֏������Ԃ��L�^
Call ApendProcessingTime(endTime - startTime)

End Sub

