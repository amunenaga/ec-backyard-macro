Attribute VB_Name = "BuildSheets"
Option Explicit

Sub �d�Z��o_�U�����V�[�g�쐬()

Const OUTPUT_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\"

OrderSheet.Activate

If InStr(Range("L1").Value, "�A�h�C���w��") > 0 Then
    MsgBox "�A�h�C�������s���ĉ������B"
End If

SyokonData.TransferOrderSheet

'�U�����p�V�[�g�̗񕝌Œ�̂��߂̕ی������
BuildSheets.TransferPickingData
BuildSheets.TransferSorterSheet

Dim i As Long

'�r������
For i = 2 To 5

    With Worksheets(i).Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
    End With

Next

'�U�����p�V�[�g��\�����āA�\�[�g
Worksheets("�U�����p�ꗗ�V�[�g").Activate
Sort.�U���p�V�[�g_�\�[�g

'�d�Z��o�p�ۑ� 100�� �I�L��
Sheets("100��").Copy
ActiveWorkbook.SaveAs Filename:=OUTPUT_FOLDER & "���t�[P�V�[�g" & Format(Date, "MMdd") & "-2-3.xlsx"
ActiveWorkbook.Close

'�d�Z��o�p�ۑ� �I����
Sheets("�I����").Copy
ActiveWorkbook.SaveAs Filename:=OUTPUT_FOLDER & "���t�[P�V�[�g" & Format(Date, "MMdd") & "-a.xlsx"
ActiveWorkbook.Close

'�U�����p�V�[�g�̗񕝒���
Call AdjustWidth

'���̃t�@�C����ۑ�
Application.DisplayAlerts = False

Const DEFAULT_XLSX_SAVE_PATH As String = "\\MOS10\Users\mos10\Desktop\���t�[\�s�b�L���O�����p�ߋ��t�@�C��\"

If Dir(DEFAULT_XLSX_SAVE_PATH, vbDirectory) <> "" Then

    ThisWorkbook.SaveAs Filename:=DEFAULT_XLSX_SAVE_PATH & "���t�[��o�E�U�����p" & Format(Date, "MMdd") & ".xlsx"

Else
    Dim SavePath As String
    SavePath = "C:" & Environ("HOMEPATH") & "\Desktop\���t�[��o�E�U�����p" & Format(Date, "MMdd") & ".xlsx"

End If

'�U�����p���i���X�g�̃V�[�g�ی���ăZ�b�g
ForSorterSheet.Protect
ForSorterSetItemSheet.Protect

'���̌�AThisWorkBook�̃R�[�h�֏�����߂��Ȃ�
End

End Sub

Private Sub TransferSorterSheet()

Worksheets("�U�����p�ꗗ�V�[�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O"
Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O �Z�b�g"

Dim i As Long, k As Long, j As Long, Order(8) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�U�����p�V�[�g�s�J�E���^
j = 2

'�U�����p�Z�b�g�V�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = CStr(Range("A" & i).Value) '�����ԍ�
    Order(1) = Range("B" & i).Value '���͂��於
    Order(2) = CStr(Range("D" & i).Value) '6�P�^
    Order(3) = Range("E" & i).Value '���i��
    Order(4) = Range("F" & i).Value '����
    Order(5) = CStr(Range("L" & i).Value) 'JAN
    Order(6) = Range("I" & i).Value '���݌�
    Order(7) = Range("K" & i).Value '���l
    Order(8) = Range("J" & i).Value '���P�[�V����
    
    '�]�L�攻��
    '7777�n�܂�Z�b�g�ƃZ�b�g���e�i
    If Range("C" & i) Like "7777*" Then
       
        With Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g")
            
            .Range("A" & j & ":I" & j).NumberFormatLocal = "@"
            .Range("A" & j & ":I" & j) = Order
            
            '���ʁA���݌ɂ͉E��
            .Range("E" & j).HorizontalAlignment = xlRight
            .Range("G" & j).HorizontalAlignment = xlRight
        
        End With
        
        j = j + 1
          
    '����ȊO
    Else
        With Worksheets("�U�����p�ꗗ�V�[�g")
        
            .Range("A" & k & ":I" & k).NumberFormatLocal = "@"
            .Range("A" & k & ":I" & k) = Order
       
           '���ʁA���݌ɂ͉E��
            .Range("E" & k).HorizontalAlignment = xlRight
            .Range("G" & k).HorizontalAlignment = xlRight
        
       
        '�I�ԂȂ��́A�s�ɐF��t����B
            If Order(8) = "" Then
                     
                With .Range("A" & k & ":I" & k).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
             
            End If
        
        End With
        
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

Worksheets("�U�����p�ꗗ�V�[�g").Range("A1").CurrentRegion.Font.Size = 9
Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("A1").CurrentRegion.Font.Size = 9

End Sub

Private Sub TransferPickingData()

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�I�����V�[�g�s�J�E���^
j = 2

'100�ԃV�[�g�s�J�E���^
k = 2

Do
    
    '�Z�b�g�R�[�h��7777�͓d�Z��o�f�[�^�Ɋ܂߂Ȃ��B
    '�����ς�6�P�^��������JAN�Œ�o����
    If Range("C" & i).Value Like "7777*" Then GoTo Continue

    '�z��ɍs���i�[
    Order(0) = Range("A" & i).Value '�����ԍ�
    Order(1) = CStr(Range("D" & i).Value) '6�P�^
    Order(2) = Range("E" & i).Value '���i��
    Order(3) = Range("F" & i).Value '����
    Order(4) = Range("G" & i).Value '���t�[�̔����i
    Order(5) = Range("I" & i).Value '���݌�
    Order(6) = Range("J" & i).Value '�I��
    
    '�]�L�攻��
    '���P�[�V�����Ȃ�
    If Order(6) = "" Then
        
       '�R�[�h�������͕�����Ƃ��āA�擪�[�����J�b�g����Ȃ��悤��
       Worksheets("�I����").Range("C" & j).NumberFormatLocal = "@"
       Worksheets("�I����").Range("B" & j & ":G" & j) = Order
    
       j = j + 1
        
    '���P�[�V�����L��
    Else
    
       '�R�[�h�������͕�����Ƃ��āA�擪�[�����J�b�g����Ȃ��悤��
       Worksheets("100��").Range("C" & k).NumberFormatLocal = "@"
       Worksheets("100��").Range("B" & k & ":H" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub

Private Sub AdjustWidth()
'�� �������ɃA���[�g���o��̂�}�~
Application.DisplayAlerts = False


Dim WidthArr As Variant
WidthArr = Array(8.13, 15.25, 11.75, 53.13, 4.75, 12.5, 5.5, 14.88, 13.5)

Dim SheetNameArr As Variant
SheetNameArr = Array("�U�����p�ꗗ�V�[�g", "�U�����p�ꗗ�V�[�g-�Z�b�g")

Dim j As Long, k As Long

For j = 0 To 1
    With Worksheets(SheetNameArr(j))
        For k = 0 To 8
            .Columns(k + 1).ColumnWidth = WidthArr(k)
        Next
    End With
Next

Application.DisplayAlerts = True

End Sub
