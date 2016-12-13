Attribute VB_Name = "Test"
Option Explicit

Sub test_ParseCode()

Dim c As Range, v As Variant

For Each v In Range("B2:B32")

    Dim Code As String
    Code = v.Value
    
    '77777�n�܂�Z�b�g�R�[�h�Ȃ�
    If Code Like "77777*" Then
        
        Debug.Print v.Address & Chr(9) & v.Value & Chr(9) & "7777�Z�b�g"
        Call ParseTiedItem(Cells(v.Row, 2))
        
    '-02 -04 -120 �n�C�t��-���� �Z�b�g�Ȃ� �n�C�t�����܂݃A���t�@�x�b�g�n�܂�łȂ�
    ElseIf InStr(Code, "-") > 1 And Not Code Like "[a-zA-Z]*" Then
        
        Debug.Print v.Address & Chr(9) & v.Value & Chr(9) & "�n�C�t���Z�b�g"
        Call ParseMultipleSet(Cells(v.Row, 2))
    
    End If

Next

End Sub

Sub test_ParseTiedItem()

    Call ParseTiedItem(Cells(25, 2))

End Sub
