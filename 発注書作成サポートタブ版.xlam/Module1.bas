Attribute VB_Name = "Module1"
Option Explicit

'リボンUIのタブを追加するアドイン作成方法
'@url http://labs.septeni.co.jp/entry/20140709/1404839264
'@url https://qiita.com/fmaeyama/items/93d10a1a5cd6cd6e9dd8
'@url http://www.ka-net.org/ribbon/ri05.html

Sub CloseAllPo(control As IRibbonControl)

Dim Flg As Integer

Flg = MsgBox(Prompt:="開いている発注書ファイルを全て閉じます。" & vbLf & "よろしいですか？", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "発注書*" Then
        Book.Close SaveChanges:=False
    End If

Next

End Sub

Sub PrintAllPo(control As IRibbonControl)

Dim Flg As Integer

Flg = MsgBox(Prompt:="開いている発注書ファイルを全て印刷します。" & vbLf & "よろしいですか？" & vbLf & vbLf & "使用プリンタ：" & Application.ActivePrinter, Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "発注書*" Then
        Book.PrintOut
    End If

Next

End Sub
