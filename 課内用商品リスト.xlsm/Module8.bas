Attribute VB_Name = "Module8"
Sub �����ꍞ��()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    '''''''''''''''''''''''
    Dim top As Worksheet
    Set top = ThisWorkbook.Worksheets("�g�b�v")
    Dim data As Worksheet
    Set data = ThisWorkbook.Worksheets("���i���")
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("�ݒ�")
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("m")
    Dim chosa As Worksheet
    Set chosa = ThisWorkbook.Worksheets("�P���������X�g")
    Dim entry As Worksheet
    Set entry = ThisWorkbook.Worksheets("���ǉ�")
    ''''''''''''''''''''''''''''''''''''''''''
    Dim datarow As String
    Dim finalrow As String
    finalrow = entry.Range("B500000").End(xlUp).Row
    With data
        For i = 2 To finalrow
            If Not IsError(Application.Match(entry.Cells(i, 2).Value, data.Range("B:B"), 0)) Then
            '�o�^������ꍇ
                datarow = Application.Match(entry.Cells(i, 2).Value, data.Range("B:B"), 0)
                For k = 1 To 30
                    If k <> 2 Then
                        If data.Cells(datarow, k).Value <> entry.Cells(i, k).Value And entry.Cells(i, k).Value <> "" Then
                            .Cells(datarow, k).Value = entry.Cells(i, k).Value
                        End If
                    End If
                Next k
            Else
            '�o�^���Ȃ��ꍇ
                datarow = data.Range("B500000").End(xlUp).Row + 1
                For k = 1 To 30
                    If data.Cells(datarow, k).Value <> entry.Cells(i, k).Value And entry.Cells(i, k).Value <> "" Then
                        .Cells(datarow, k).Value = entry.Cells(i, k).Value
                    End If
                Next k
            End If

        Next i
    End With

    MsgBox ("�������ꍞ�݊���")
    ''''''''''''''''''''''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub
Sub �����ꍞ�݃N���A()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    '''''''''''''''''''''''
    Dim top As Worksheet
    Set top = ThisWorkbook.Worksheets("�g�b�v")
    Dim data As Worksheet
    Set data = ThisWorkbook.Worksheets("���i���")
    Dim config As Worksheet
    Set config = ThisWorkbook.Worksheets("�ݒ�")
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("m")
    Dim chosa As Worksheet
    Set chosa = ThisWorkbook.Worksheets("�P���������X�g")
    Dim entry As Worksheet
    Set entry = ThisWorkbook.Worksheets("���ǉ�")
    ''''''''''''''''''''''''''''''''''''''''''
    Dim datarow As String
    Dim finalrow As String
    finalrow = entry.Range("B500000").End(xlUp).Row
    entry.Range("A2:AZ" & finalrow).Cells.Clear
    ''''''''''''''''''''''''''''''''''''''''''
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
End Sub



