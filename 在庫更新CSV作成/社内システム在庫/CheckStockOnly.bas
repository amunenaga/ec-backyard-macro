Attribute VB_Name = "CheckStockOnly"
Sub CheckEolInStockOnly()
'�����s�A�݌ɂ̂݃V�[�g�̏��i�R�[�h�ɂ��āA
'�敪���`�F�b�N�A�敪�����p�ԁA�̔����~->�p�ԁE�I���ɃR�s�[�A���łɃ��X�g�A�b�v�ςȂ�s�폜
'�敪�������E�ݔp�ō݌ɐ���0->�p�ԁE�I���ɃR�s�[�A���łɃ��X�g�A�b�v�ςȂ�s�폜

StockOnly.Activate

Application.ScreenUpdating = False

Call SetRangeName

Dim i As Long
i = 2

Do Until IsEmpty(Cells(i, 3))
      
   Call CheckEol(i)  '�s�ԍ����Q�Ɠn���ŊY���s��6�P�^�ɂ��ă`�F�b�N�A�Q�Ɠn���Ȃ̂ŏ������i��1�s�i�߂邱�Ƃ��ł���B
    
Loop

Application.ScreenUpdating = True

End Sub

Private Sub CheckEol(i As Long)

Dim Code As String
Code = Cells(i, 3)

'�p�ԁE�I�����X�g�ɓ]�L�ς݂��`�F�b�N

If WorksheetFunction.CountIf(Range("EolCodeRange"), Code) > 0 Then
        
        Rows(i).Delete
        Exit Sub

End If

Dim sy As Syokon
sy = SyokonMaster.GetSyokonQtyKubun(Code)


'�p�ԁE�̔��I�����X�g�ւ̓]�L�����Ɉ�v���邩�H

If InStr(sy.Status, "�p��") > 0 Or InStr(sy.Status, "�̔����~") > 0 Then

        Call PostEol(i)
        Exit Sub

ElseIf InStr(sy.Status, "�ݔp") > 0 Or InStr(sy.Status, "�����i") > 0 Then
    
    If sy.Quantity <= 0 Then
        
        Call PostEol(i)
        Exit Sub
    
    End If

End If

i = i + 1 '�s���폜���Ȃ������ꍇ�̂݁Ai��1�i�߂�

End Sub

Private Sub PostEol(i As Long)

Dim Code As String
Code = Cells(i, 3).Value

Call addCode(Code, "EolCodeRange")
Rows(i).Delete

End Sub
