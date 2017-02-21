Attribute VB_Name = "Main"
Option Explicit

Sub �󒍃t�@�C���Ǎ�()

OrderSheet.Activate

If Not Range("B2").Value = "" Then
    MsgBox "�f�[�^�擾�ςł��B"
    End
End If

Dim LineBuf As Variant

'�t�@�C������I�u�W�F�N�g����
Dim FSO As New FileSystemObject

' Meisai.csv��tyumon_H.csv��CSV�t�@�C���̃p�X���Z�b�g
'���ׂƒ����w�b�_�[�̂���t�H���_���w��A�Ō�K��\�}�[�N
Const CSV_PATH As String = "C:\Users\mos10\Desktop\���t�[\"
Const ALTER_CSV_PATH As String = "\\MOS10\���t�[\"

Dim MeisaiPath As String, TyumonhPath As String

If FSO.FileExists(CSV_PATH & "Meisai.csv") Then

    MeisaiPath = CSV_PATH & "Meisai.csv"
    TyumonhPath = CSV_PATH & "tyumon_H.csv"

ElseIf FSO.FileExists(ALTER_CSV_PATH & "Meisai.csv") Then
   
    MeisaiPath = ALTER_CSV_PATH & "Meisai.csv"
    TyumonhPath = ALTER_CSV_PATH & "tyumon_H.csv"

Else
    
    'TODO:�t�@�C���w��œǂݍ��܂���
    
    MsgBox "meisai.csv �t�@�C���Ȃ�"
    End

End If

Call ReadMeisai(MeisaiPath)

Call ReadTyumonH(TyumonhPath)

'�}�N���N���{�^��������
OrderSheet.Shapes(1).Delete

'�A�h�C���p�̍s�E�� �\��
Dim LastRow As Long
LastRow = Range("D1").SpecialCells(xlCellTypeLastCell).Row

Range("I1").Value = "�A�h�C���w�� �䒠�F9998"
Range("I2:L2") = Array(2, 4, LastRow, 9)

MsgBox "�A�h�C�������s���ĉ������B"

'�A�h�C���Ń��P�[�V�����擾�O�̏����I��

End Sub

'���̈ʒu�ɁA�A�h�C���ł̃��P�[�V�����擾���K�v�B
'DB�ڑ����ăf�[�^�Ƃ��Ă�����Main������1�N���b�N�ɂȂ�B

Sub �d�Z��o_�U�����V�[�g�쐬()

'�A�h�C����̏���
OrderSheet.Activate

'�A�h�C�������s�̍ۂ́A�_�C�A���O�Ōx�����o���ďI��
If InStr(Range("L1").Value, "�A�h�C���w��") > 0 Then
    MsgBox "�A�h�C�������s���ĉ������B"
    End
End If

'�����ȃ��P�[�V�������J�b�g
DataVaridate.ModifyOrderSheet

'�󒍈ꗗ�V�[�g�̏C���I���A�V�[�g��ی�A�f�[�^���b�N��������B
OrderSheet.Protect

'���[�����̓d�Z����o�f�[�^�ۑ��A�U�����V�[�g�쐬
Dim Mall As Variant, Malls As Variant

Malls = Array("���t�[")

For Each Mall In Malls
    '�s�b�L���O�V�[�g�쐬�E�ۑ�
    Call BuildSheets.OutputPickingData(CStr(Mall))
    
    '�U�����p�V�[�g�쐬
    Call BuildSheets.CreateSorterSheet(CStr(Mall))

Next

'�A���[�g�_�C�A���O��}�~
Application.DisplayAlerts = False

'�e���v���[�g�V�[�g���폜
Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Delete
Worksheets("�U���p�e���v���[�g").Delete

'���̃t�@�C����ۑ�
Dim PutFileName As String
PutFileName = "�s�b�L���O�E�U��" & Format(Date, "MMdd") & ".xlsx"

'�[���I��Try-Catch�ŕۑ������s
On Error Resume Next
    
    'Try
    '�T�[�o�[02�̏���̃t�H���_�֕ۑ��c�e�X�g�x�b�h�̃��t�[�p��MOS10\�f�X�N�g�b�v�̏���t�H���_�B
    ThisWorkbook.SaveAs FileName:="C:" & Environ("HOMEPATH") & "\Desktop\���t�[\�s�b�L���O�����p�ߋ��t�@�C��\" & PutFileName
    
    'catch
    If Err Then
    '�T�[�o�[02�֌q����Ȃ��Ƃ��́A���sPC�̃f�X�N�g�b�v�֕ۑ�
        Err.Clear
        ThisWorkbook.SaveAs FileName:="C:" & Environ("HOMEPATH") & "\Desktop\" & PutFileName

    End If
    
    'catch2
    If Err Then
        Err.Clear
        MsgBox "�t�@�C����ۑ��ł��܂���ł����B�蓮�Ŗ��O��t���ĕۑ����Ă��������B"
    End If

'On Error Goto 0 �錾��Err�͉��������
On Error GoTo 0

'���sPC�f�t�H���g�̃v�����^�Ńv�����g�A�E�g
'�v�����^�̎w��́AWindowsA
Dim i As Long
For i = 2 To Worksheets.Count

    Worksheets(i).PrintOut

Next

'���̌�AThisWorkBook�̃R�[�h�֏�����߂��Ȃ�
End

End Sub
