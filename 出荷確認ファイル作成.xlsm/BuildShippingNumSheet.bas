Attribute VB_Name = "BuildShippingNumSheet"
Option Explicit
Sub 佐川ヤマト_シート作成()

'ボタン削除
Worksheets("トップ").Shapes(1).Delete

'TSV/CSVファイルパス指定
Dim Path As Collection
Set Path = GetCsvPath()

'データ読み込み
Call LoadAmazon(Path.Item("Amazon"))
Call LoadRakuten(Path.Item("Rakuten"))
Call LoadYahoo(Path.Item("Yahoo"))

'運送会社別にシートへコピー
Call SortByCarrier("佐川急便")
Call SortByCarrier("ヤマト運輸")

'列幅調整
Dim i As Long
For i = 1 To Worksheets.Count
    Worksheets(i).Range("A1").CurrentRegion.Columns.AutoFit
Next i

'後処理、データリンク削除、セルの「名前」削除
Dim qt As QueryTable
For Each qt In Worksheets("トップ").QueryTables
    qt.Delete
Next qt

Dim nm As Name
For Each nm In ActiveWorkbook.Names
    nm.Delete
Next nm

'ファイル保存
Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\出荷確認_" & Format(Date, "yyyyMMdd") & ".xlsx", FileFormat:=xlWorkbookDefault
Application.DisplayAlerts = True

End Sub

Sub LoadAmazon(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("$B$2")) 'パスは動的に、書き出し先はB2固定。Amazonから取り込むので
    .Name = "Amazon"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 4
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 9, 9, 9, 9, 2, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Amazon")

End Sub

Sub LoadRakuten(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("B1").End(xlDown).Offset(1, 0)) 'パスと書き出し先は動的に決める
    .Name = "楽天"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 2
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 9, 2, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("楽天")

End Sub

Sub LoadYahoo(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("B1").End(xlDown).Offset(1, 0)) 'パスと書き出し先は動的に決める
    .Name = "yahoo"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 932
    .TextFileStartRow = 2
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 2, 9, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Yahoo")

End Sub

Private Sub FillMallName(ByVal MallName As String)
'CSV読み込み後にA列をモール名で埋めます。

Dim StartRow As Double, EndRow As Double, i As Double
StartRow = IIf(Range("A2").Value = "", 2, Range("A1").End(xlDown).Row + 1)
EndRow = Range("B1").End(xlDown).Row

For i = StartRow To EndRow
    Cells(i, 1).Value = MallName
Next i

End Sub

Sub SortByCarrier(ByVal CarrierName As String)
'運送会社名を受け取って、運送会社毎のシートへ送り状番号をコピー

'運送会社とフィルター条件のマッピング
Dim Criteria As Variant

Select Case CarrierName
    
    Case "佐川急便"
        Criteria = "4031*"
    
    Case "ヤマト運輸"
        Criteria = Array("7645*", "3046*")

End Select

'送り状番号をフィルターしてコピー
With Range("A1").CurrentRegion
    .AutoFilter Field:=3, Criteria1:=Criteria, Operator:=xlFilterValues
    .Copy Worksheets(CarrierName).Range("A1")
    .AutoFilter 'オートフィルター解除
End With

End Sub

Function GetCsvPath() As Collection
'送り状番号CSVパスを取得、3ファイルまで同時指定可能

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .Filters.Clear
    .Filters.Add "Amazon,楽天,Yahoo!", "*.tsv; *.csv"
    .InitialFileName = "\\Server02\商品部\ネット販売関連\出荷通知"

    .Show
    
    If .SelectedItems.Count >= 4 Then
       MsgBox "ファイル指定が3つを超えています。"
       End
    End If
    
    Dim Paths As Collection, CurrentPath As String, i As Long
    Set Paths = New Collection
    
    For i = 1 To 3
        CurrentPath = fd.SelectedItems.Item(i)
        Select Case True
            Case CurrentPath Like "*amazon*"
                Paths.Add Item:=CurrentPath, Key:="Amazon"
            
            Case CurrentPath Like "*楽天*"
                Paths.Add Item:=CurrentPath, Key:="Rakuten"
                
            Case CurrentPath Like "*yahoo*"
                Paths.Add Item:=CurrentPath, Key:="Yahoo"
            
        End Select
    Next

End With

Set GetCsvPath = Paths

End Function
