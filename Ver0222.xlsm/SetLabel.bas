Attribute VB_Name = "SetLabel"
Private Sub Labeling()
'�g���Ă܂���  �����ԍ��́u���O�v���V�[�g�ɐU���ƌ������֗����Ȃ�
Dim rng As Range

r = 317

Do While Cells(r, 2).Value <> ""

    Id = Cells(r, 2).Value

    Set rng = Range(Cells(r, 1), Cells(r, 7))
    
    i = 0
    
    Do '���T�C�Y���ׂ��s�������肷��A�I�t�Z�b�g���Ē����ԍ����`�F�b�N
        
        i = i + 1
    
    Loop While rng.Cells(1, 2).Offset(i, 0).Value = Id
    
    If i > 1 Then
        
        Set rng = rng.Resize(i, 7)
    
    End If
         
    Dim Label As String
   
    Label = "NO" & Id
    
    Call SetLabel(Label, rng)
    
    r = r + i

Loop

End Sub

Private Sub SetLabel(Name As String, rng As Range)

End Sub
