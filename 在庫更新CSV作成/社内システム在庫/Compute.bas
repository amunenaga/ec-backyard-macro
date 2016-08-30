Attribute VB_Name = "Compute"
Sub UploadQuantity()

yahoo6digit.Activate

Application.ScreenUpdating = False


Dim colAbstract As Integer, colQuantity As Integer, colAllow As Integer, colStatus As Integer

colAbstract = yahoo6digit.Rows(1).Find("abstract").Column

On Error Resume Next
    colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
    colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
    colStatus = yahoo6digit.Rows(1).Find("status").Column

    If Err Then
    
        '�w�b�_�[���Ȃ���ΒǋL
        yahoo6digit.Range("A1").End(xlToRight).Offset(0, 1).Resize(1, 3) = Array("quantity", "allow-overdraft", "status")
        Err.Clear
    
    End If
    
On Error GoTo 0

Dim r As Range

With yahoo6digit 'With�\�����ł̓I�u�W�F�N�g�Q�Ƃ��J��Ԃ���Ȃ����߁A�������������ɂȂ�炵��

    For Each r In .Range("YahooCodeRange")
        
        '�݌ɐݒ菜�O�V�[�g�ɂ���΁A�ȉ��̏����͍s��Ȃ��AContinue�֔��
        If WorksheetFunction.CountIf(ExceptQty.Range("ExceptCodeRange"), r.Value) > 0 Then GoTo Continue
       
        Dim i As Long
        i = r.Row

        'Item�I�u�W�F�N�g�𐶐�
        
        Dim YahooItem As Object
        Set YahooItem = New Item
        
        Call YahooItem.Constractor(r.Value, .Cells(i, colAbstract).Value)
        
        '�V�[�g�֏�������
       
        .Cells(i, colQuantity).Value = YahooItem.GetQuantity
        
        If YahooItem.GetAvailablePurchase Then  'AvailablePurchase��Bool�l�Ȃ̂�1/0�ɒu�������ďo��
            .Cells(i, colAllow).Value = 1
        Else
            .Cells(i, colAllow).Value = 0
        End If
               
        .Cells(i, colStatus).Value = YahooItem.status
       
Continue:
    
       Set Item = Nothing
       
    Next r

End With

End Sub
