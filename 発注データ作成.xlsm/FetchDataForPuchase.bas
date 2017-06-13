Attribute VB_Name = "FetchDataForPuchase"
Option Explicit

Sub CreateQuantitySheet()

Call LoadAllPicking
Call SumPuchaseRequest

'発注に必要な情報をデータベース・Excelファイルから取得
Call FetchSyokonData
Call FetchExcellForPurchase

Call CalcPurchaseQuantity

End Sub

Sub FetchSyokonData()
'商魂商品マスターのデータ取得
'実際に読みに行くテーブルは、受注詳細確認用に毎朝レプリケーションを作るServer3のEC課用マスタ

'接続のためのオブジェクトを定義、DB接続設定をセット
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'商品コードのレンジをセット、1セルずつSQL実行
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim sql As String, Code As String
    Code = r.Value
        
    sql = "SELECT 商品コード, 取扱区分, ロット数, 仕入原価, 仕入先, 仕入先マスタ.仕入先略称 " & _
          "FROM 商品マスタ JOIN 仕入先マスタ ON 商品マスタ.仕入先 = 仕入先マスタ.仕入先コード " & _
          "WHERE 商品コード = " & Code & "OR JANコード = '" & Code & "'"
    
    Set DbRs = DbCnn.Execute(sql)

    If Not DbRs.EOF Then
        Cells(r.Row, 3).Value = DbRs("ロット数")
        Cells(r.Row, 4).Value = DbRs("仕入先")
        Cells(r.Row, 5).Value = DbRs("仕入先略称")
        Cells(r.Row, 10).Value = DbRs("仕入原価")
        Cells(r.Row, 2).Value = GetKubun(DbRs("取扱区分"))
        
        'JAN受注分の商品コード置換、主にAmazon卸用
        If Len(r.Value) > 6 Then
            r.NumberFormatLocal = "@"
            r.Value = IIf(Len(DbRs("商品コード")) = 5, "0" & DbRs("商品コード"), DbRs("商品コード"))
        End If
    
    End If

Next

End Sub

Sub FetchExcellForPurchase()
'発注用商品情報のデータ取得

Dim DataBook As Workbook, DataSheet As Worksheet, PurDataCodeRange As Range, PurDataJanRange As Range
Set DataSheet = OpenPurDataBook().Worksheets("商品情報")
Set PurDataJanRange = DataSheet.Range(Cells(1, 1), Cells(DataSheet.UsedRange.Rows.Count, 1))
Set PurDataCodeRange = DataSheet.Range(Cells(1, 2), Cells(DataSheet.UsedRange.Rows.Count, 2))

ThisWorkbook.Activate
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange

    Dim Code As String, HitRow As Double
        
    Code = r.Value

    On Error Resume Next
        HitRow = WorksheetFunction.Match(Code, PurDataCodeRange, 0)

        If Err Then
            Err.Clear
            HitRow = WorksheetFunction.Match(Code, PurDataJanRange, 0)
            
            If Err Or IsEmpty(Cells(r.Row, 4).Value) Then '仕入先コードが商魂にもエクセルにもなければ、発注できないので注意書きを入れる
                Cells(r.Row, 2).Value = "発注用商品情報 データなし"
                GoTo Continue
            End If
        
        End If
    
    On Error GoTo 0
        
    '手配時注意、メーカーロット、仕入先名
    Cells(r.Row, 2).Value = Cells(r.Row, 2).Value & DataSheet.Cells(HitRow, 35).Value '手配時注意
    Cells(r.Row, 11).Value = DataSheet.Cells(HitRow, 5).Value '発注用商品情報のロット数
    Cells(r.Row, 12).Value = DataSheet.Cells(HitRow, 4).Value '仕入先名
    
    '仕入先コード、原価、仕入先名は6ケタにない時のみ入れる
    If IsEmpty(Cells(r.Row, 4).Value) Then
    
        Cells(r.Row, 4).Value = DataSheet.Cells(HitRow, 32).Value '仕入先コード
        Cells(r.Row, 5).Value = DataSheet.Cells(HitRow, 4).Value '仕入先名
        Cells(r.Row, 10).Value = DataSheet.Cells(HitRow, 13).Value '原価

    End If

Continue:

Next

End Sub

Sub FetchExcellJanInventory()
'棚なし在庫表データの確認


End Sub

Sub CalcPurchaseQuantity()

Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim i As Long
    i = r.Row
    
    Dim Rot As Double, Qty As Long, RequestQty As Double
    If IsEmpty(Cells(i, 11).Value) Then
        Rot = 1
    Else
        Rot = Cells(i, 11).Value
    End If
    
    RequestQty = Cells(i, 9).Value
    
    Qty = WorksheetFunction.Ceiling(Rot, RequestQty)

    Cells(i, 1).Value = Qty
Next

End Sub

Private Function GetKubun(ByVal KubunCode As Integer) As String
Dim tmp As String

Select Case KubunCode
    Case 3
        tmp = "商魂:販売中止"
    Case 7
        tmp = "商魂:在庫廃番"
    Case 8
        tmp = "商魂:在庫処分"
    Case 9
        tmp = "商魂:メーカー廃番"
    Case Else
        tmp = ""
End Select

GetKubun = tmp

End Function

Private Function OpenPurDataBook() As Workbook

'発注用商品情報のファイルを開きます。
'実行中のエクセルで発注用商品情報のファイルを開いていれば、そのワークブックを返します。

Const PUR_DATA_EXCELL_PATH As String = "\\Server02\商品部\ネット販売関連\発注関連\発注用商品情報.xlsm"

Dim WorkBookName As String
WorkBookName = Dir(PUR_DATA_EXCELL_PATH)

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = WorkBookName Then
        wb.Activate
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(PUR_DATA_EXCELL_PATH)

ret:

Set OpenPurDataBook = wb

End Function
