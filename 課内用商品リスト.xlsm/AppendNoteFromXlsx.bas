Attribute VB_Name = "AppendNoteFromXlsx"
Sub InsertCaution()

'���̃G�N�Z���u�b�N���A���V�[�g�͈͖̔͂���w��̂���
'�C�~�f�B�G�C�g�ŁAWorkbooks(1).name�Ń��[�N�u�b�N�����m�F�ł���B
Set Rng = Workbooks("����d�H�����X�g�b�v��.xlsx").Sheets(1).Range("D2:D76")

'�ǋL��������������w��
Dim AdditionalNote As String
AdditionalNote = "�����݌ɗL 16�N4��"

For Each r In Rng

    Dim Code As String
    Code = r.Value
    
    Dim c As Range
    
    'B����������āA�Y���Z�P�^������΁A�d����ɒǋL����
    With Workbooks("�����p���i���.xlsm").Worksheets("���i���").Columns(2)
    
        Set c = .Find(what:=Code, LookIn:=xlValues, LookAt:=xlWhole)

        If Not c Is Nothing Then
           '�ŏ��̃Z���̃A�h���X���o����
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
