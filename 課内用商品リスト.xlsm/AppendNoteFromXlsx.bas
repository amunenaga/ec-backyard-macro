Attribute VB_Name = "AppendNoteFromXlsx"
Sub InsertCaution()

'���̃G�N�Z���u�b�N���A���V�[�g�͈͖̔͂���w��̂���
'�C�~�f�B�G�C�g�ŁAWorkbooks(1).name�Ń��[�N�u�b�N�����m�F�ł���B
Set Rng = Workbooks("�����X�g�b�v��.xlsx").Sheets(1).Range("D2:D76")

'�ǋL��������������w��
Dim AdditionalNote As String
AdditionalNote = "�����X�g�b�v"

For Each r In Rng

    Dim Code As String
    Code = r.Value
    
    Dim c As Range
    
    'B����������āA�Y���R�[�h������΁A�d�����ɒǋL����
    With Workbooks("���i���X�g.xlsm").Worksheets("���i���").Columns(2)
    
        Set c = .Find(what:=Code, LookIn:=xlValues, LookAt:=xlWhole)

        If Not c Is Nothing Then
           '�ŏ��̃Z���̃A�h���X���L�^
           FirstAddress = c.Address
           
           '�J�Ԃ��������A�����𖞂������ׂẴZ������������
           Do
              
               c.Offset(0, 2) = c.Offset(0, 2) & " " & AdditionalNote
               
               Set c = .FindNext(c)
               If c Is Nothing Then Exit Do
           
           Loop Until c.Address = FirstAddress
         
         End If

    End With

Next

End Sub
