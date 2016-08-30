Attribute VB_Name = "BuildSheets"
Option Explicit

Sub �d�Z��o_�U�����V�[�g�쐬()

Const OUTPUT_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\"

OrderSheet.Activate

If InStr(Range("A1").Value, "�A�h�C���w��") > 0 Then
    MsgBox "�A�h�C�������s���ĉ������B"
End If

SyokonData.TransferOrderSheet

BuildSheets.TransferPickingData
BuildSheets.TransferSorterSheet

Dim i As Long

'�r������
For i = 2 To 5

    With Worksheets(i).Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
    End With

Next

'��o�p�쐬 100�� �I�L��
Sheets("100��").Copy

ActiveWorkbook.SaveAs filename:=OUTPUT_FOLDER & "���t�[P�V�[�g" & Format(Date, "MMdd") & "-2-3.xlsx"

ActiveWorkbook.Close

'��o�p�쐬 �I����

Sheets("�I����").Copy

ActiveWorkbook.SaveAs filename:=OUTPUT_FOLDER & "���t�[P�V�[�g" & Format(Date, "MMdd") & "-a.xlsx"

ActiveWorkbook.Close

'���̃t�@�C����ۑ�
Application.DisplayAlerts = False
ThisWorkbook.SaveAs filename:="\\MOS10\Users\mos10\Desktop\���t�[\�s�b�L���O�����p�ߋ��t�@�C��\" & "���t�[��o�E�U�����p" & Format(Date, "MMdd") & ".xlsx"

End Sub

Private Sub TransferSorterSheet()

Worksheets("�U�����p�ꗗ�V�[�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O"
Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O �Z�b�g"

Dim i As Long, k As Long, j As Long, Order(8) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�I�����V�[�g�s�J�E���^
j = 2

'100�ԃV�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = CStr(Range("A" & i).Value) '�����ԍ�
    Order(1) = Range("B" & i).Value '���͂��於
    Order(2) = Range("D" & i).Value '6�P�^
    Order(3) = Range("E" & i).Value '���i��
    Order(4) = Range("F" & i).Value '����
    Order(5) = Range("L" & i).Value 'JAN
    Order(6) = Range("I" & i).Value '���݌�
    Order(7) = Range("K" & i).Value '���l
    Order(8) = Range("J" & i).Value '���P�[�V����
    
    '�]�L�攻��
    '7777�n�܂�Z�b�g�ƃZ�b�g���e�i
    If Order(2) Like "7777*" Or Range("C" & i).Value = "Set" Then
       
        Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("A" & j & ":I" & j).NumberFormatLocal = "@"
        Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("A" & j & ":I" & j) = Order

        j = j + 1
          
    '����ȊO
    Else
    
        Worksheets("�U�����p�ꗗ�V�[�g").Range("A" & k & ":I" & k).NumberFormatLocal = "@"
        Worksheets("�U�����p�ꗗ�V�[�g").Range("A" & k & ":I" & k) = Order
       
       '�I�ԂȂ��́A�s�ɐF��t����B
        If Order(8) = "" Then
                
            With Worksheets("�U�����p�ꗗ�V�[�g").Range("A" & k & ":I" & k).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        
        End If
    
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub

Private Sub TransferPickingData()

Dim i As Long, k As Long, j As Long, Order(5) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�I�����V�[�g�s�J�E���^
j = 2

'100�ԃV�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = Range("A" & i).Value '�����ԍ�
    Order(1) = Range("D" & i).Value '6�P�^
    Order(2) = Range("E" & i).Value '���i��
    Order(3) = Range("F" & i).Value '����
    Order(4) = Range("G" & i).Value '���t�[�̔����i
    Order(5) = Range("J" & i).Value '�I��
    
    '�]�L�攻��
    '���P�[�V�����Ȃ�
    If Order(5) = "" Then
        
        If Not Order(0) Like "7777*" Then
           
           Worksheets("�I����").Range("B" & j).NumberFormatLocal = "@"
           Worksheets("�I����").Range("B" & j & ":G" & j) = Order
        
           j = j + 1
        
        End If
        
    '���P�[�V�����L��
    Else
    
       Worksheets("100��").Range("B" & k).NumberFormatLocal = "@"
       Worksheets("100��").Range("B" & k & ":G" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub
