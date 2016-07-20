Attribute VB_Name = "Module6"
Sub �o�b�N�A�b�v�R�s�[()

    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "�y�V"
    
    Dim yahoo As String
    yahoo = "���t�["
    
    Dim data As String
    data = "���i���"
    
    Dim pasted As String
    pasted = "�\�t"
    
    Dim config As String
    config = "�ݒ�"
    '------------------------

    Dim fileName As String
    fileName = ThisWorkbook.Name
    
    Dim before As String
    before = ThisWorkbook.Path & "\" & fileName
    
    Dim after As String
    after = ThisWorkbook.Worksheets(config).Range("C22").Value & "\BU" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".xls"

    Workbooks.Add
    
    Dim finalrow As String
    finalrow = ThisWorkbook.Worksheets(data).Range("B500000").End(xlUp).Row
    
    Dim finalRange As String
    finalRange = "A1:Z" & finalrow
    
    ActiveWorkbook.Worksheets(1).Range(finalRange).Value = ThisWorkbook.Worksheets(data).Range(finalRange).Value
    
    ActiveWorkbook.SaveAs fileName:=after
    ActiveWorkbook.Close True

    


End Sub

Sub �A�}�]�����o�^�����o()

    Call �o�b�N�A�b�v�R�s�[
    Call �A�}�]�����o�^�����oA
    Call �A�}�]�����o�^�����oB
    Call �A�}�]�����o�^�����oC
    
    MsgBox ("�捞����!!" & vbNewLine & vbNewLine & "����ȍ~�̔�����Ƃɉe�����o�邽�߁A����̔����܂łɕK��SKU�Ə��i���ȊO�̏���ǉ����Ă��������B")

End Sub

Sub �A�}�]�����o�^�����oA()

    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "�y�V"
    
    Dim yahoo As String
    yahoo = "���t�["
    
    Dim data As String
    data = "���i���"
    
    Dim pasted As String
    pasted = "�\�t"
    
    Dim config As String
    config = "�ݒ�"
    
    '------------------------------
    ThisWorkbook.Worksheets(pasted).Cells.Clear


    Dim filepath As String
    filepath = ThisWorkbook.Worksheets(config).Range("C3").Value & "/" & ThisWorkbook.Worksheets(config).Range("B3").Value & ".txt"

    Dim fileOpen As Workbook
    Set fileOpen = Workbooks.Open(filepath)
    
    Dim finalrow As String
    finalrow = fileOpen.Worksheets(1).Range("A300000").End(xlUp).Row
    
    Dim copyRange As String
    copyRange = "A1:AZ" & finalrow
    
    '�\�t
    With ThisWorkbook.Worksheets(pasted)
    
        
        .Range("C:C").NumberFormat = "@"
        .Range(copyRange).Value = fileOpen.Worksheets(1).Range(copyRange).Value
        .Range("D:AZ").Clear
    
    End With

    fileOpen.Close False



End Sub

Sub �A�}�]�����o�^�����oB()

    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "�y�V"
    
    Dim yahoo As String
    yahoo = "���t�["
    
    Dim data As String
    data = "���i���"
    
    Dim pasted As String
    pasted = "�\�t"
    
    Dim config As String
    config = "�ݒ�"
    
    '------------------------------



    'SKU��ҏW
    Dim pastedFinalRow As String
    For i = 2 To ThisWorkbook.Worksheets(pasted).Range("A300000").End(xlUp).Row
    
        If Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 5 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 6 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 10 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "000" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 11 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "00" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 12 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "0" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        ElseIf Len(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) = 13 And IsNumeric(ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) Then
        
            With ThisWorkbook.Worksheets(pasted)
            
                .Cells(i, 3).NumberFormat = "#########################"
                .Cells(i, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(i, 3).NumberFormat = "@"
                .Cells(i, 3).Value = "" & ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
            
            End With
        
        End If
    
    Next i


End Sub
Sub �A�}�]�����o�^�����oC()

    Dim amazon As String
    amazon = "amazon"
    
    Dim rakuten As String
    rakuten = "�y�V"
    
    Dim yahoo As String
    yahoo = "���t�["
    
    Dim data As String
    data = "���i���"
    
    Dim pasted As String
    pasted = "�\�t"
    
    Dim config As String
    config = "�ݒ�"
    
    '------------------------------

    Dim finalrow As String
    For i = 2 To ThisWorkbook.Worksheets(pasted).Range("B300000").End(xlUp).Row
    
        If Application.CountIf(ThisWorkbook.Worksheets(data).Range("B:B"), ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value) < 1 Then
        
            finalrow = ThisWorkbook.Worksheets(data).Range("B300000").End(xlUp).Row
            With ThisWorkbook.Worksheets(data)
            
                .Cells(finalrow + 1, 2).NumberFormat = "@"
                .Cells(finalrow + 1, 2).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 3).Value
                .Cells(finalrow + 1, 3).NumberFormat = "@"
                .Cells(finalrow + 1, 3).Value = ThisWorkbook.Worksheets(pasted).Cells(i, 1).Value

            
            End With
        
        End If
    
    Next i

ThisWorkbook.Worksheets(pasted).Cells.Clear


End Sub

