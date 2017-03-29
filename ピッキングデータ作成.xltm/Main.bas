Attribute VB_Name = "Main"
Option Explicit

'�����y�EAmazon�v���C�����̍쐬�����L�^����t���O BuildSheets.PreparePickingBook�Ŏg�p
Public IsSecondPicking As Boolean
Public IsTimeStampMode As Boolean

Sub �s�b�L���O_�U���쐬()

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

ShowProgress.ProgressBar.Value = 2
ShowProgress.StepMessageLabel = "CSV�Ǎ���"

Call LoadCsv

ShowProgress.ProgressBar.Value = 3
ShowProgress.StepMessageLabel = "���P�[�V�����f�[�^�擾��"
Application.Wait Now + TimeValue("00:00:02")
'1�b�ҋ@���ăv���O���X�o�[���X�V

'DB�ڑ��A���P�[�V�����擾�A�󒍃f�[�^�̏C��
Call ConnectDB.Make_List
Call DataValidate.FilterLocation

'�󒍃f�[�^�V�[�g�ł̏����I���A�V�[�g�ی��������
OrderSheet.Protect

'���[�����̓d�Z����o�f�[�^�ۑ��A�U�����V�[�g�쐬
Dim Mall As Variant, Malls As Variant, ProgressStep As Long

Malls = Array("Amazon", "�y�V", "Yahoo")
ProgressStep = 3

For Each Mall In Malls
        
    ProgressStep = ProgressStep + 1
    ShowProgress.ProgressBar.Value = ProgressStep
    ShowProgress.StepMessageLabel = Mall & "�f�[�^������"
    
    '���[�����̎󒍌������[�����Ȃ�t�@�C���������Ȃ��B
    If WorksheetFunction.CountIf(OrderSheet.Range("F:F"), Mall & "*") = 0 Then GoTo Continue
    
    '�s�b�L���O�V�[�g�쐬�E�ۑ�
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '�U�����p�V�[�g�쐬
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Continue:

Next

'�V�[�g�폜�A�ۑ����̃A���[�g�_�C�A���O��}�~
Application.DisplayAlerts = False

'�e���v���[�g�V�[�g���폜
Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Delete
Worksheets("�U���p�e���v���[�g").Delete

ShowProgress.ProgressBar.Value = 7
ShowProgress.StepMessageLabel = Mall & "�ۑ�������"

Dim DeskTop As String, SaveFileName As String, SavePath As String
Const SAVE_FOLDER = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[��\�ߋ��f�[�^\"

SaveFileName = "�s�b�L���O�E�U��" & Format(Date, "MMdd")

OrderSheet.Activate
If Dir(SAVE_FOLDER, vbDirectory) <> "" Then
    '���ɖ{���t�@�C��������΁A�����t���ĕۑ�
    If Dir(SAVE_FOLDER & SaveFileName & ".xlsx") = "" Then
        SavePath = SAVE_FOLDER & SaveFileName
    Else
        SavePath = SAVE_FOLDER & SaveFileName & "-" & Format(Time, "hhmm")
    End If
    
        ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Else
    
    Dim DeskTopPath As String
    If Dir(DeskTopPath & SaveFileName & ".xlsx") = "" Then
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName
    Else
        DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & SaveFileName & "-" & Format(Time, "hhmm")
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

'�v���O���X�o�[�������ďI�����b�Z�[�W
ShowProgress.Hide
MsgBox prompt:="��������", Buttons:=vbInformation, Title:="�����I��"

End Sub
