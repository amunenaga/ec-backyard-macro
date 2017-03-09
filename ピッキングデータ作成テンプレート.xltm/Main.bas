Attribute VB_Name = "Main"
Option Explicit
Sub ピッキング_振分生成()

OrderSheet.Activate

If Not Range("A2").Value = "" Then
    MsgBox "データ取得済です。"
    End
End If

'プログレスバーの準備
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 9

    .Show vbModeless
End With

'マクロ起動ボタンを消去
'OrderSheet.Shapes(1).Delete

ShowProgress.ProgressBar.Value = 2
ShowProgress.StepMessageLabel = "CSV読込中"

Call LoadCsv

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "ロケーションデータ取得中"
Application.Wait Now + TimeValue("00:00:01")
'1秒待機してプログレスバーを更新

Call ConnectDB.Make_List

'無効なロケーションをカット
Call DataVaridate.ModifyOrderSheet

'モール毎の電算室提出データ保存、振分けシート作成
Dim Mall As Variant, Malls As Variant, ProgressStep As Long

Malls = Array("Amazon", "楽天", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
    
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "データ処理中"
    
    'ピッキングシート作成・保存
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '振分け用シート作成
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Next

'アラートダイアログを抑止
Application.DisplayAlerts = False

'テンプレートシートを削除
'Worksheets("ピッキングシート提出用テンプレート").Delete
'Worksheets("振分用テンプレート").Delete

'このファイルを保存


ShowProgress.ProgressBar.Value = 7
ShowProgress.StepMessageLabel = Mall & "保存処理中"

Dim DeskTop As String, PutFileName As String, SavePath As String
Const SAVE_PATH = "\\server02\商品部\ネット販売関連\ピッキング\クロスモール\過去データ"

PutFileName = "ピッキング・振分" & Format(Date, "MMdd") & ".xlsx"

'当日8時取込分のタイムスタンプ無しファイルがないか確認
If Dir(SAVE_PATH & PutFileName) <> "" Then
    PutFileName = Format(Time, "hh:mm") & PutFileName
End If
    
'擬似的なTry-Catchで保存
On Error Resume Next
    
    ThisWorkbook.SaveAs Filename:="\\server02\商品部\ネット販売関連\ピッキング\クロスモール\過去データ", FileFormat:=xlWorkbookDefault
    
    'catch
    If Err Then
        Err.Clear
        MsgBox "ネット販売関連に繋がりませんでした、デスクトップへ保存します。"
        Dim DeskTop As String, SavePath As String
        DeskTop = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop")
    
        If Dir(DeskTop & "\" & PutFileName) <> "" Then
            PutFileName = Replace(PutFileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "AM/PMhhmm"))
        End If
    
    End If

'On Error Goto 0 宣言でErrは解除される
On Error GoTo 0

ShowProgress.ProgressBar.Value = 8
ShowProgress.StepMessageLabel = Mall & "振分シート プリント"

'実行PCデフォルトのプリンタでプリントアウト
'プリンタの指定してなければ、Windowsのデフォルトプリンタを使う。
Dim i As Long
For i = 2 To Worksheets.Count

    'Worksheets(i).PrintOut

Next

ShowProgress.ProgressBar.Value = 9

OrderSheet.Activate

'プログレスバーを消して終了メッセージ
ShowProgress.Hide
MsgBox Prompt:="処理完了", Buttons:=vbInformation

End Sub
