Attribute VB_Name = "sheetViewControler"
Option Explicit


Sub �Ǎ��t�H�[����\��()
    
''���j���̈��ڂ̋N���́A�t�B���^�[�����A�󒍃f�[�^�������`�F�b�N�𑣂��A���[�g
''���j����FirstOpen��True�Ȃ�A���[�g
'
'If Weekday(Date) = 2 Then
'
'    If LogSheet.Range("B5").Value = True Then
'
'        MsgBox prompt:="���j���̏���N���ł��B" & vbLf & "���c�ꗗ�̑O�T�����m�F���Ă��������B", _
'                Buttons:=vbExclamation
'
'        OrderSheet.AutoFilterMode = False
'
'        LogSheet.Range("B5").Value = False
'
'        End
'
'    End If
'
'Else
'
''���j�ȊO�Ȃ�True���Z�b�g���Ă���
'
'    LogSheet.Range("B5").Value = True
'
'End If

OpPanel.Show

End Sub


Sub hideWishCol()
    
    OrderSheet.Outline.ShowLevels ColumnLevels:=1
    
End Sub


Sub �������̂ݕ\��()
Attribute �������̂ݕ\��.VB_Description = "������ �� �� �u�o�גʒm���O�v��\��"
Attribute �������̂ݕ\��.VB_ProcData.VB_Invoke_Func = "u\n14"

OrderSheet.Activate

Application.ScreenUpdating = False

'�I�[�g�t�B���^�[���Z�b�g����Ă��Ȃ���΁A15��ڂ́u�����v�󗓂Ɓu�o�גʒm���O�v�̂ݕ\���Őݒ�
If Not OrderSheet.AutoFilterMode Then
    
    Range("A1").AutoFilter Field:=15, Criteria1:="=�o�גʒm���O", Operator:=xlOr, Criteria2:="="

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

End Sub

Sub ������̋󗓂̂ݕ\��()
Attribute ������̋󗓂̂ݕ\��.VB_ProcData.VB_Invoke_Func = " \n14"

'fillterShippingNull

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
