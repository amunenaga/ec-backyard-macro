Attribute VB_Name = "SheetControler"
Option Explicit


Sub �Ǎ��t�H�[����\��()

OpPanel.Show

End Sub

Sub �v�]����\��()
    
    OrderSheet.Outline.ShowLevels ColumnLevels:=1
    
End Sub

Sub �������̂ݕ\��()
Attribute �������̂ݕ\��.VB_Description = "������ �� �� �u�o�גʒm���O�v��\��"
Attribute �������̂ݕ\��.VB_ProcData.VB_Invoke_Func = "u\n14"

OrderSheet.Activate

Application.ScreenUpdating = False

'�I�[�g�t�B���^�[���Z�b�g����Ă��Ȃ���΁A15��ڂ́u�����v�󗓂Ɓu�o�גʒm���O�v�̂ݕ\���Őݒ�
If Not OrderSheet.AutoFilterMode Then
    
    Range("A1").AutoFilter Field:=15, Criteria1:="="

Else
    '�t�B���^�[���Z�b�g����Ă���΁A�Z�b�g�������ݒ�
    Dim i As Integer
    For i = 1 To 17

        If i = 15 Then
           Range("A1").AutoFilter Field:=i, Criteria1:="=�o�גʒm���O", Operator:=xlOr, Criteria2:="="

        Else
        
           Range("A1").AutoFilter i  '���̓t�B���^�[�����ACriteria�w����ȗ��Łu�S�āv�\��
        
        End If

    Next

End If

OrderSheet.setProtect

End Sub

Sub ������̋󗓂̂ݕ\��()
Attribute ������̋󗓂̂ݕ\��.VB_ProcData.VB_Invoke_Func = " \n14"

'fillterShipping

OrderSheet.Activate

Application.ScreenUpdating = False

'�I�[�g�t�B���^�[���Z�b�g����Ă��Ȃ���΁A15��ڂ́u�����v�󗓂̂ݕ\��
If Not OrderSheet.AutoFilterMode Then
    
    Range("A1").AutoFilter Field:=15, Criteria1:="="

Else
    '�t�B���^�[���Z�b�g����Ă���΁A�Z�b�g�������ݒ�
    Dim i As Integer
    For i = 1 To 17
        
        If i = 15 Then
           Range("A1").AutoFilter Field:=i, Criteria1:="="
        
        Else
           Range("A1").AutoFilter i  '���̓t�B���^�[�����ACriteria�w����ȗ��Łu�S�āv�\��
        
        End If
    
    Next

End If

End Sub

Sub �S�Ă̔����󋵂�\��()

OrderSheet.Activate

Application.ScreenUpdating = False
    
    Range("A1").AutoFilter Field:=15

End Sub
