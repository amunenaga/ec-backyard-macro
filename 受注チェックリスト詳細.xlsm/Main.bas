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


'��o�p�t�@�C���쐬����
Transfer.��o�p�V�[�g�֓]�L

MsgBox "�V�[�g�쐬 ����"

End Sub
