Attribute VB_Name = "Module1"
Option Explicit

'���{��UI�̃^�u��ǉ�����A�h�C���쐬���@
'@url http://labs.septeni.co.jp/entry/20140709/1404839264
'@url https://qiita.com/fmaeyama/items/93d10a1a5cd6cd6e9dd8
'@url http://www.ka-net.org/ribbon/ri05.html

Sub CloseAllPo(control As IRibbonControl)

Dim Flg As Integer

Flg = MsgBox(Prompt:="�J���Ă��锭�����t�@�C����S�ĕ��܂��B" & vbLf & "��낵���ł����H", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "������*" Then
        Book.Close SaveChanges:=False
    End If

Next

End Sub

Sub PrintAllPo(control As IRibbonControl)

Dim Flg As Integer

Flg = MsgBox(Prompt:="�J���Ă��锭�����t�@�C����S�Ĉ�����܂��B" & vbLf & "��낵���ł����H" & vbLf & vbLf & "�g�p�v�����^�F" & Application.ActivePrinter, Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "������*" Then
        Book.PrintOut
    End If

Next

End Sub
