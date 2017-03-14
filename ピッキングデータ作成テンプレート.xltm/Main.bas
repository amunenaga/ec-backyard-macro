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
    .ProgressBar.Max = 8

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

'�󒍃f�[�^�V�[�g�ł̏����I���A�V�[�g�ی��������
OrderSheet.Protect

'���[�����̓d�Z����o�f�[�^�ۑ��A�U�����V�[�g�쐬
Dim Mall As Variant, Malls As Variant, ProgressStep As Long

Malls = Array("Amazon", "�y�V", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
    
    '���[�����̎󒍌������[�����Ȃ�t�@�C���������Ȃ��B
    If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), CStr(Mall) & "*") = 0 Then GoTo Continue
    
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "�f�[�^������"
    
    '�s�b�L���O�V�[�g�쐬�E�ۑ�
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '�U�����p�V�[�g�쐬
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Continue:

Next

'�A���[�g�_�C�A���O��}�~
Application.DisplayAlerts = False

'�e���v���[�g�V�[�g���폜
Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Delete
Worksheets("�U���p�e���v���[�g").Delete

ShowProgress.ProgressBar.Value = 7
ShowProgress.StepMessageLabel = Mall & "�ۑ�������"
'���̃t�@�C����ۑ�

Dim DeskTop As String, SaveFileName As String, SavePath As String
Const SAVE_FOLDER = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[��\�ߋ��f�[�^\"

SaveFileName = "�s�b�L���O�E�U��" & Format(Date, "MMdd") & ".xlsx"


If Dir(SAVE_FOLDER, vbDirectory) <> "" Then
    '���ɖ{���t�@�C��������΁A�����t���ĕۑ�
    If Dir(SAVE_FOLDER & SaveFileName & ".xlsx") = "" Then
        SavePath = SAVE_FOLDER & SaveFileName
    Else
        SavePath = SAVE_FOLDER & Format(Time, "hhmm") & SaveFileName
    End If
    
        ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Else
    
    Dim DeskTopPath As String
    If Dir(DeskTopPath & SaveFileName & ".xlsx") = "" Then
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName
    Else
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & Format(Time, "hhmm") & SaveFileName
    End If
    
    MsgBox "�l�b�g�̔��֘A�Ɍq����Ȃ����߁A" & SaveFileName & "���f�X�N�g�b�v�ɕۑ����܂��B"
        
    ActiveWorkbook.SaveAs Filename:=DeskTopPath, FileFormat:=xlWorkbookDefault

End If

ShowProgress.ProgressBar.Value = 8
ShowProgress.StepMessageLabel = Mall & "�U���V�[�g �v�����g"

'���sPC�f�t�H���g�̃v�����^�Ńv�����g�A�E�g
'�v�����^�̎w�肵�ĂȂ���΁AWindows�̃f�t�H���g�v�����^���g���B
Dim i As Long
For i = 2 To Worksheets.Count

    Worksheets(i).Protect
    Worksheets(i).PrintOut

Next

OrderSheet.Activate

'�v���O���X�o�[�������ďI�����b�Z�[�W
ShowProgress.Hide
MsgBox Prompt:="��������", Buttons:=vbInformation, Title:="�����I��"

End Sub
