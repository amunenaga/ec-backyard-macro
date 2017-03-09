Attribute VB_Name = "Main"
Option Explicit
Sub �s�b�L���O_�U������()

OrderSheet.Activate

If Not Range("A2").Value = "" Then
    MsgBox "�f�[�^�擾�ςł��B"
    End
End If

'�v���O���X�o�[�̏���
With ShowProgress
    .ProgressBar.Min = 1
    .ProgressBar.Max = 9

    .Show vbModeless
End With

'�}�N���N���{�^��������
'OrderSheet.Shapes(1).Delete

ShowProgress.ProgressBar.Value = 2
ShowProgress.StepMessageLabel = "CSV�Ǎ���"

Call LoadCsv

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "���P�[�V�����f�[�^�擾��"
Application.Wait Now + TimeValue("00:00:01")
'1�b�ҋ@���ăv���O���X�o�[���X�V

Call ConnectDB.Make_List

'�����ȃ��P�[�V�������J�b�g
Call DataVaridate.ModifyOrderSheet

'���[�����̓d�Z����o�f�[�^�ۑ��A�U�����V�[�g�쐬
Dim Mall As Variant, Malls As Variant, ProgressStep As Long

Malls = Array("Amazon", "�y�V", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
    
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "�f�[�^������"
    
    '�s�b�L���O�V�[�g�쐬�E�ۑ�
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '�U�����p�V�[�g�쐬
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Next

'�A���[�g�_�C�A���O��}�~
Application.DisplayAlerts = False

'�e���v���[�g�V�[�g���폜
'Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Delete
'Worksheets("�U���p�e���v���[�g").Delete

'���̃t�@�C����ۑ�


ShowProgress.ProgressBar.Value = 7
ShowProgress.StepMessageLabel = Mall & "�ۑ�������"

Dim DeskTop As String, PutFileName As String, SavePath As String
Const SAVE_PATH = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[��\�ߋ��f�[�^"

PutFileName = "�s�b�L���O�E�U��" & Format(Date, "MMdd") & ".xlsx"

'����8���捞���̃^�C���X�^���v�����t�@�C�����Ȃ����m�F
If Dir(SAVE_PATH & PutFileName) <> "" Then
    PutFileName = Format(Time, "hh:mm") & PutFileName
End If
    
'�[���I��Try-Catch�ŕۑ�
On Error Resume Next
    
    ThisWorkbook.SaveAs Filename:="\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[��\�ߋ��f�[�^", FileFormat:=xlWorkbookDefault
    
    'catch
    If Err Then
        Err.Clear
        MsgBox "�l�b�g�̔��֘A�Ɍq����܂���ł����A�f�X�N�g�b�v�֕ۑ����܂��B"
        Dim DeskTop As String, SavePath As String
        DeskTop = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop")
    
        If Dir(DeskTop & "\" & PutFileName) <> "" Then
            PutFileName = Replace(PutFileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "AM/PMhhmm"))
        End If
    
    End If

'On Error Goto 0 �錾��Err�͉��������
On Error GoTo 0

ShowProgress.ProgressBar.Value = 8
ShowProgress.StepMessageLabel = Mall & "�U���V�[�g �v�����g"

'���sPC�f�t�H���g�̃v�����^�Ńv�����g�A�E�g
'�v�����^�̎w�肵�ĂȂ���΁AWindows�̃f�t�H���g�v�����^���g���B
Dim i As Long
For i = 2 To Worksheets.Count

    'Worksheets(i).PrintOut

Next

ShowProgress.ProgressBar.Value = 9

OrderSheet.Activate

'�v���O���X�o�[�������ďI�����b�Z�[�W
ShowProgress.Hide
MsgBox Prompt:="��������", Buttons:=vbInformation

End Sub
