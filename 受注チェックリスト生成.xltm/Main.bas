Attribute VB_Name = "Main"
Option Explicit

Sub 受注チェックリスト生成()

'CSV読込、作業シートへコピー
Importer.CSV読込
Transfer.作業シートへデータ抽出

'作業シートでのデータ修正処理
Worksheets("作業シート").Activate

SetParser.セット分解
Transfer.住所結合
Transfer.JAN転記
Transfer.楽天商品名修正


Transfer.提出用シートへ転記

'取込、作業シート込みのエクセルファイルを保存
Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:="受注チェックリスト" & Format(Date, "MMdd") & ".xlsx"
    
    '提出シートのみファイルを保存
    'Sheets("提出シート").Copy

    'ActiveWorkbook.SaveAs FileName:="提出" & Format(Date, "MMdd") & ".xlsx"

Application.DisplayAlerts = True

'セット商品リストブックを閉じる
Dim w As Workbook
For Each w In Workbooks
    If w.Name = "ｾｯﾄ商品ﾘｽﾄ.xls" Then w.Close False
Next

MsgBox "シート生成 完了"

End Sub

Sub 生成のみ実行()

Transfer.作業シートへデータ抽出

'作業シートでのデータ修正処理
Worksheets("作業シート").Activate

SetParser.セット分解
Transfer.住所結合
Transfer.JAN転記
Transfer.楽天商品名修正

Transfer.提出用シートへ転記

Dim w As Workbook
For Each w In Workbooks
    If w.Name = "ｾｯﾄ商品ﾘｽﾄ.xls" Then w.Close False
Next

MsgBox "シート作成 完了" & vbLf & "ファイル名を指定して保存して下さい。"

End Sub
