Attribute VB_Name = "Module4"
Sub SKU��������()

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

    For i = 2 To ThisWorkbook.Worksheets(data).Range("B200000").End(xlUp).Row
    
        With ThisWorkbook.Worksheets(data)
            '�d���於
            .Cells(i, 2).NumberFormat = "@"
            .Cells(i, 2).Value = ThisWorkbook.Worksheets(data).Cells(i, 2).Value
            
            'JAN
            .Cells(i, 5).NumberFormat = "@"
            .Cells(i, 5).Value = ThisWorkbook.Worksheets(data).Cells(i, 2).Value
            
        End With
        
        
    Next i


End Sub
