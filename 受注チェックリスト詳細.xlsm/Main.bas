Attribute VB_Name = "Main"
Option Explicit

Sub �󒍃`�F�b�N���X�g����()

'CSV�Ǎ��A��ƃV�[�g�փR�s�[
Importer.CSV�Ǎ�
Transfer.��ƃV�[�g�փf�[�^���o

'��ƃV�[�g�ł̃f�[�^�C������
Worksheets("��ƃV�[�g").Activate

SetParser.�Z�b�g����
Transfer.�Z������
Transfer.JAN�]�L
Transfer.�y�V���i���C��


Transfer.��o�p�V�[�g�֓]�L

'��o�t�@�C���ۑ�
Sheets("��o�V�[�g").Copy

Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:="��o" & Format(Date, "MMdd") & ".xlsx"
    ActiveWorkbook.Close
Application.DisplayAlerts = True

Dim w As Workbook

For Each w In Workbooks
    If w.Name = "��ď��iؽ�.xls" Then w.Close False
Next

MsgBox "�t�@�C���쐬 ����"

ThisWorkbook.Close False

End Sub
