Attribute VB_Name = "Module5"
Sub ASIN�ǉ�()



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
    
    Dim lt As String
    lt = "LT"
    
    '------------------------------
    Dim nextRow As String
    For i = ThisWorkbook.Worksheets(lt).Range("A200000").End(xlUp).Row To 1 Step -1
    
        If ThisWorkbook.Worksheets(lt).Cells(i, 1).Value <> "" And ThisWorkbook.Worksheets(lt).Cells(i, 2).Value <> "" Then
            
            nextRow = Application.Match(ThisWorkbook.Worksheets(lt).Cells(i, 2).Value, ThisWorkbook.Worksheets(data).Range("B:B"), 0)
            ThisWorkbook.Worksheets(data).Cells(nextRow, 20).Value = ThisWorkbook.Worksheets(lt).Cells(i, 1).Value
        
        End If
        
    Next i
    


End Sub
Sub ASIN�ǉ����ڍ폜()



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
    
    Dim lt As String
    lt = "LT"
    
    '------------------------------
    ThisWorkbook.Worksheets(lt).Range("A:B").ClearContents


End Sub
