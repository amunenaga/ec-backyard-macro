Attribute VB_Name = "Module1"
Option Explicit

Sub PrintAllPo()

Dim Flg As Integer

Flg = MsgBox(Prompt:="�J���Ă���t�@�C���S�Ă�������܂��B" & vbLf & "��낵���ł����H", Buttons:=vbOKCancel)

If Flg <> 1 Then Exit Sub

Dim i As Long
For i = 1 To Workbooks.Count

    Workbooks(i).PrintOut

Next

End Sub
