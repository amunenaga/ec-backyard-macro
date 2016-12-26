Attribute VB_Name = "Module1"
Option Explicit
Sub ICOKURI連結()

    Call ConcatenateICOKURI
    
    '実行PCのデスクトップをフルパスで
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
        
    Dim PutFolder As String
    PutFolder = CStr(wsh.SpecialFolders("desktop")) & "\ICOKURI連結済み\"
        
    '実行したPCのデスクトップに「ICOKURI連結済みフォルダ」が無い場合
    If Dir(PutFolder) = "" Then
    
        PutFolder = Replace(PutFolder, "ICOKURI連結済み\", "")
    
    End If
    
    '連結したCSVをxlsx形式で保存
    Application.DisplayAlerts = False
    
        ThisWorkbook.SaveAs Filename:=PutFolder & "ICOKURI" & Format(Date, "MMdd") & ".xlsx"
     
    Application.DisplayAlerts = True

    'マクロ起動ボタンを削除
    Sheet1.Shapes(1).Delete
    
End Sub

Private Sub ConcatenateICOKURI()
Const B2_FOLDER As String = "\\Server02\商品部\ネット販売関連\梱包室データ\B2ヤマトデータ\"

Dim FolderName(2) As String
FolderName(0) = "アマゾン"
FolderName(1) = "楽天"
FolderName(2) = "ヤフー"

Dim v As Variant

For Each v In FolderName

    Dim B2CsvPath As String
    B2CsvPath = FindTodaysCSV(B2_FOLDER & v)
    
    If B2CsvPath <> "" Then
        Call ImportICOKURI(B2CsvPath)
    End If

Next

End Sub

Sub ImportICOKURI(ByVal CsvPath)
Attribute ImportICOKURI.VB_ProcData.VB_Invoke_Func = " \n14"

'最終行のRange、書き出し先セルの特定
Dim LastRow As Long
LastRow = Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row

Dim PutStartCell As Range

If LastRow = 1 Then
    Set PutStartCell = Range("A1")
Else
    Set PutStartCell = Cells(LastRow + 1, 1)
End If

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=PutStartCell _
    )
    .Name = "ICOKURI"
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
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
    2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

ActiveWorkbook.Connections(1).Delete

End Sub

Function FindTodaysCSV(ByVal CsvFolderPath As String) As String

'CSVフォルダのパス指定、最後\マークかチェック
If Not Right(CsvFolderPath, 1) = "\" Then
    CsvFolderPath = CsvFolderPath & "\"
End If

'実行時バインディングでファイルオブジェクト
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, TodayCSV As Object
      
'指定フォルダー内のFileNameを含むファイル名を調べて、本日 日付ファイルを一つ取得する

For Each f In FSO.GetFolder(CsvFolderPath).Files

    If DateDiff("D", f.DateLastModified, DateValue(Date)) = 0 Then
    
        Set TodayCSV = f
        Exit For
        
    End If

Next

If f Is Nothing Then
    
      MsgBox Prompt:=Replace(Mid(CsvFolderPath, InStr(CsvFolderPath, "B2ヤマトデータ")), "\", " ") & vbLf & "ICOKURIデータなし"

      Exit Function
      
End If

FindTodaysCSV = CsvFolderPath & f.Name

End Function
