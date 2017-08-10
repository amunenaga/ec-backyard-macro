Attribute VB_Name = "InsertFormula"
Option Explicit

Sub 入荷日の算出式を入力()
'算出式が未入力の行を調べる
'A列からK列の商品名列までは、手配データ作成から入力される

ThisWorkbook.Worksheets("納期リスト").Activate

Dim InsertEndCell As Range, InsertStartCell As Range, TargetRange As Range

'最終行のL列に数式が入っていれば、日付算出式は入力済なのでとして、プロシージャ終了
If Cells(Range("A1").End(xlDown).Row, 12).Formula <> "" Then Exit Sub

'最終行のH列のセル
Set InsertEndCell = Cells(Range("A1").End(xlDown).Row, 8)

'L列で式が入っている最終行から一行下のH列のセル
Set InsertStartCell = Cells(InsertEndCell.Row, 12).End(xlUp).Offset(1, -4)

Set TargetRange = Range(InsertStartCell, InsertEndCell)

'セットしたレンジに対して式を入れる
Dim r As Variant

For Each r In TargetRange

    'H列、I列はメーカーシートの入荷に関する文言
    On Error Resume Next
        r.Offset(0, 0).Value = WorksheetFunction.VLookup(Cells(r.Row, 3).Value, Worksheets("メーカー").Range("B3:D1000"), 2, False)
        r.Offset(0, 1).Value = WorksheetFunction.VLookup(Cells(r.Row, 3).Value, Worksheets("メーカー").Range("B3:D1000"), 3, False)
    On Error GoTo 0

    'K , L列は､W列の日付から入荷日を算出する式を入れる、行番号を＠と置いて置換して式の文字列を作る
    r.Offset(0, 3).Formula = Replace("=IFERROR(VALUE(IF($J@="""",IF($I@=""当日"",$E@,IF($I@=""翌日"",$E@+1,IF($I@=""翌々日"",$E@+2,""""))),$J@+1)),"""")", "@", r.Row)
    r.Offset(0, 4).Formula = Replace("=IF($K@="""","""",IF(MOD($K@,7)=0,$K@+2,IF(MOD($K@,7)=1,$K@+1,$K@)))", "@", r.Row)

Next

End Sub
