Attribute VB_Name = "Main"
Option Explicit

Sub ���t�[�݌ɍX�V�t�@�C������()

'�����̋敪�A�݌Ɍ���V�[�g�A�p�ԃV�[�g���`�F�b�N���āA
'���t�[�ɃA�b�v���[����݌ɐ��AAllow-overdraft���Z�b�g���܂��B

'���Ԍv�������܂�
Dim startTime As Long
startTime = Timer

'���t�[CSV���C���|�[�g
Prepare.ImportYahooCSV

Prepare.ImportSyokonAddinData

'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Prepare.SetRangeName

'---��������---

'�ݒ�݌ɐ��Z�o�A���񂹉ێZ�o

Compute.UploadQuantity

'�ꎞ��~���㏑��
HaltSheet.setHalt

'�ݔp�A������0�͔p�ԁE�I���ֈړ�
CheckEolInStockOnly

'�敪�̕\�����t�B���^�[
SetStatusFilter

'���t�[�f�[�^�V�[�g����CSV��ۑ�
Call Output.QtyCsv

ThisWorkbook.Save

'�I������������
Dim endTime As Long
endTime = Timer

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub
