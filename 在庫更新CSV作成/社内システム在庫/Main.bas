Attribute VB_Name = "Main"
Option Explicit

Sub ���t�[�݌ɍX�V�t�@�C������()

'�����̋敪�A���t�[�f�[�^��Abstract�A�݌Ɍ���V�[�g�A�p�ԃV�[�g���`�F�b�N���āA
'���t�[�ɃA�b�v���[����݌ɐ��AAllow-overdraft���Z�b�g���܂��B

'���Ԍv�������܂�
Dim startTime As Long
startTime = Timer

'SLIMS�f�[�^���C���|�[�g
Slims.ImportSlimsCSV

'���t�[CSV���C���|�[�g
Prepare.FetchYahooCSV

'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Prepare.SetRangeName

'---��������---

'�ݒ�݌ɐ��Z�o�A���񂹉ێZ�o

Compute.UploadQuantity


'�ꎞ��~���㏑��
Call halt.setHalt

'�ݔp�A������0�͔p�ԁE�I���ֈړ�
Call CheckEolInStockOnly

'���t�[�f�[�^�V�[�g����CSV��ۑ�
Output.QtyCsv


'�I������������
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub
