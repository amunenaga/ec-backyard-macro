Attribute VB_Name = "Module1"
Sub AppendQtyCsv()

Const CSV_SH_NAME As String = "�A�b�v���[�h�p�݌�"

Application.ScreenUpdating = False

'CSV�ǋL����
Dim FSO As New FileSystemObject
Dim Csv As Object

Dim CsvFileName As String
CsvFileName = "�݌ɍX�V" & Format(Date, "mmdd") & ".csv"

'�ǋL���[�h ForAppending �Ńt�@�C�����J��
Set Csv = FSO.OpenTextFile(FileName:=ThisWorkbook.Path & "\" & CsvFileName, IOMode:=8)

Dim LastRow As Long
LastRow = Worksheets(CSV_SH_NAME).UsedRange.Rows.Count

For i = 2 To LastRow
    
    With Worksheets(CSV_SH_NAME)
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next
End Sub

