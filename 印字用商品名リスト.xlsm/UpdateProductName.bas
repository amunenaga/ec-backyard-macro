Attribute VB_Name = "UpdateProductName"
Sub Update()

'更新元データのブック・シートを指定
Dim SourceSheet As Worksheet
Set SourceSheet = Workbooks(1).Worksheets(1)

Worksheets("最終").Activate

'「最終」シートに既に入っている商品コードのレンジ
Dim CodeRange As Range
Set CodeRange = Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1))

Dim i As Long, Code As String, Name As String
i = 2

'開いているワークブックの二つ目なので、毎回指定し直す必要がある。
Do
    Code = SourceSheet.Cells(i, 1).Value

    Dim HitRow As Long
    HitRow = WorksheetFunction.Match(Code, CodeRange, 0)
    
    Cells(HitRow, 2).Value = SourceSheet.Cells(i, 4).Value
    
    i = i + 1
    
Loop Until IsEmpty(SourceSheet.Cells(i, 1).Value)

End Sub

