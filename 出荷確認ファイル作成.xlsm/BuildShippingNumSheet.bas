Attribute VB_Name = "BuildShippingNumSheet"
Option Explicit
Sub 佐川ヤマト_シート作成()

'ファイル保存
Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\出荷確認_" & Format(Date, "yyyyMMdd") & ".xlsx", FileFormat:=xlWorkbookDefault
Application.DisplayAlerts = True


'TSV/CSVファイルパス指定
Dim Path As Collection
Set Path = GetCsvPath()

'Amazon、楽天、YahooのCSVパスがあるか確認
If Path.Count = 0 Then
    MsgBox Prompt:="CSVファイルが指定されていません、処理を終了します。", Buttons:=vbCritical
    Exit Sub
End If

'ボタン削除
'Worksheets("トップ").Shapes(1).Delete

On Error Resume Next

    Dim ErrorMall As String 'シートへ読み込めなかったモールを追記する
    
    'Try
    Call LoadAmazon(Path.Item("Amazon"))
    'catch
    If Err Then
        Err.Clear '明示的にクリアーしていないと、次のIf Err構文でTrueとなる
        ErrorMall = ErrorMall & "Amazon" & vbLf
    End If
    
    'Try
    Call LoadRakuten(Path.Item("Rakuten"))
    'catch
    If Err Then
        Err.Clear
        ErrorMall = ErrorMall & "楽天" & vbLf
    End If

    'Try
    Call LoadYahoo(Path.Item("Yahoo"))
    'catch
    If Err Then
        Err.Clear
        ErrorMall = ErrorMall & "ヤフー" & vbLf
    End If
    
On Error GoTo 0

'運送会社別にシートへコピー
'運送会社毎の送り状番号冒頭5ケタは、SortByCarrierプロシージャにてハードコーディング
Call SortByCarrier("佐川急便")
Call SortByCarrier("ヤマト運輸")

'列幅調整
Dim i As Long
For i = 1 To Worksheets.Count
    Worksheets(i).Range("A1").CurrentRegion.Columns.AutoFit
Next i

'後処理  データリンク削除＆セルの「名前」削除
Dim qt As QueryTable
For Each qt In Worksheets("トップ").QueryTables
    qt.Delete
Next qt

Dim nm As Name
For Each nm In ActiveWorkbook.Names
    nm.Delete
Next nm

'振分後の保存と完了メッセージ
ThisWorkbook.Save

If ErrorMall = "" Then
    MsgBox Prompt:="処理完了", Buttons:=vbInformation
Else
    MsgBox Prompt:="処理完了" & vbLf & vbLf & ErrorMall & "データが読み込めませんでした。", Buttons:=vbExclamation
End If

End Sub

Sub LoadAmazon(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=GetDestRange())
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
    "TEXT;" & Path, Destination:=GetDestRange())
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
    "TEXT;" & Path, Destination:=GetDestRange()) 'パスと書き出し先は動的に決める
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
'引数でコピー先シートを指定するので、シート名と運送会社を合わせること。
'Select文内に、運送会社-送り状番号冒頭5ケタの組み合わせをコーディングしている。
'採番が変わった際は、Case文内の絞り込み用文字列を変更すること。

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

'ファイルダイアログにてAmazon・楽天・ヤフーのTSV/CSVファイルを複数同時指定してもらう
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    'ファイル選択ダイアログの設定
    .Filters.Clear
    .Filters.Add "Amazon,楽天,Yahoo!", "*.tsv; *.csv"
    .InitialFileName = "\\Server02\商品部\ネット販売関連\出荷通知"

    'ダイアログ表示
    .Show
    
    'ファイル選択後の処理
    If .SelectedItems.Count = 0 Then
    
        MsgBox "ファイル指定がキャンセルされました。"
        End
    
    ElseIf .SelectedItems.Count >= 4 Then
        MsgBox "ファイル指定が3つを超えています。"
        End
    
    End If
    
    Dim Paths As Collection, CurrentPath As String, i As Long
    Set Paths = New Collection
    
    '選択されたファイルパスを、Amazon・楽天・ヤフーをキーとしたコレクションに入れ直す
    For i = 1 To .SelectedItems.Count
        CurrentPath = .SelectedItems.Item(i)
        
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

Private Function GetDestRange() As Range

'書き出し先セルを決める、B2が空の時にEndコマンドでは1,048,576行まで飛んでしまうので。
Dim r As Range
If IsEmpty(Range("B2")) Then
    Set r = Range("B2")
Else
    Set r = Range("B1").End(xlDown).Offset(1, 0)
End If

Set GetDestRange = r

End Function
