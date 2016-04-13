Attribute VB_Name = "TestModule"
Option Explicit

Sub test_LastSecondInbentryGetQuantity()

    If IsNumeric(LastSecondInventry.getQuantity("4953571171524")) Then
        MsgBox "OK"
    Else
        MsgBox "NG!"
    End If

End Sub

Sub test_appendtime()

    Call Log.ApendProcessingTime(21)

End Sub

Sub test_GetSecondInventryQuantity()

    Debug.Print SecondInventry.getQuantity("4953571172101")
    
End Sub

Sub test_SortSecondInventry()
    
    Worksheets("�l�b�g�p�݌�").Range("A1").AutoFilter
    
    With Worksheets("�l�b�g�p�݌�").AutoFilter.Sort
    
        .SortFields.Clear '�\�[�g�t�B�[���h����U�N���A�[
        
        '�\�[�g�t�B�[���h���w��
        .SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        
        '�\�[�g�����w��
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        
        '�\�[�g�K�p
        .Apply

    End With


End Sub
