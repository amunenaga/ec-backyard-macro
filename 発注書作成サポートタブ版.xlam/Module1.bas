Attribute VB_Name = "Module1"
Option Explicit

'���{��UI�̃^�u�̃J�X�^�}�C�Y���@
'@url
'@url

Sub CloseAllPo(control As IRibbonControl)

Dim Flg As Integer

Flg = MsgBox(Prompt:="�J���Ă���t�@�C���S�ĕ��܂��B" & vbLf & "��낵���ł����H", Buttons:=vbOKCancel)

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

Flg = MsgBox(Prompt:="�J���Ă���t�@�C���S�Ĉ�����܂��B" & vbLf & "��낵���ł����H", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim Book As Variant

For Each Book In Workbooks

    If Book.Name Like "������*" Then
        Book.PrintOut
    End If

Next

End Sub
