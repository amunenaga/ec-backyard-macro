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


Transfer.�A�b�v���[�h�p�V�[�g�֓]�L

'�Z�b�g���i���X�g�u�b�N�����
Dim w As Workbook
For Each w In Workbooks
    If w.Name = "��ď��iؽ�.xls" Then w.Close False
Next

Worksheets("�A�b�v���[�h�V�[�g").Range("A1").Select

'�捞�A��ƃV�[�g���݂̃G�N�Z���t�@�C����ۑ�
Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:=ActiveWorkbook.Path & "\�󒍃`�F�b�N���X�g_" & Format(Date, "yyyymmdd") & ".xlsm", FileFormat:=52
    'ActiveWorkbook.SaveAs FileName:=ActiveWorkbook.Path & "\�󒍃`�F�b�N���X�g.xlsx", FileFormat:=xlOpenXMLWorkbook
    
Application.DisplayAlerts = True

'�f�[�^�x�[�X�ւ̓o�^�������s
Call InsertDB.CodeI2JAN_E

End Sub


