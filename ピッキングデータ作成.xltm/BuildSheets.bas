Attribute VB_Name = "BuildSheets"
Option Explicit

Sub CreateSorterSheet(MallName As String)

'�P�̏��i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_�U���p"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName
End With
Dim ForSorterSheet As Worksheet
Set ForSorterSheet = ActiveSheet

'�Z�b�g���i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_�U���p-�Z�b�g"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName & "-�Z�b�g���i"
End With
Dim ForSorterSetItemSheet As Worksheet
Set ForSorterSetItemSheet = ActiveSheet

'�A�N�e�B�u�ȃV�[�g�̓R�s�[�����V�[�g����󒍃V�[�g�ɕς��Ă���
OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�U�����p�V�[�g�s�J�E���^
j = 2

'�U�����p�Z�b�g�V�[�g�s�J�E���^
k = 2

Do

    '�����œn���ꂽ���[���ȊO�͔�΂�
    If Not Range("F" & i).Value Like (MallName & "*") Then GoTo Continue

    '�z��ɍs���i�[
    Order(0) = CStr(Range("B" & i).Value) '�󒍎��R�[�h
    Order(1) = ValidateName(Range("C" & i).Value) '���i���A���ȂǍ폜������œ]�L
    Order(2) = Range("D" & i).Value '�󒍐���
    Order(3) = CStr(Range("L" & i).Value) 'JAN
    Order(4) = Range("K" & i).Value '�L�����P�[�V����
    Order(5) = Range("N" & i).Value '���݌�
    
    
    '���݌ɂ��擾�ł��ĂȂ��Ƃ��́A������C�A�E�g�̊֌W�̂��ߋ�1��������Ă���
    If Order(5) = "" Then Order(5) = " "
    
    '����JAN���󗓂��A�󒍎��R�[�h��JAN�Ȃ�JAN���ڂɓ����
    If Order(3) = "" Then
        Dim RawCode As String
        RawCode = Range("B" & i).Value
        If RawCode Like String(13, "#") _
            And Not RawCode Like "77777*" Then
                Order(3) = RawCode
        End If
    End If
    
    '�]�L�攻��
    '7777�n�܂�Z�b�g�ƃZ�b�g�\�����i�A�󒍎��R�[�h7777***
    If Range("B" & i) Like "7777*" Then
       
       Order(0) = Range("I" & i).Value
       
        With ForSorterSetItemSheet
            
            .Range("A" & j & ":F" & j).NumberFormatLocal = "@"
            .Range("A" & j & ":F" & j) = Order
            
            '���ʁA���݌ɂ͉E��
            .Range("C" & j).HorizontalAlignment = xlRight
            .Range("F" & j).HorizontalAlignment = xlRight
        
        End With
        
        j = j + 1
          
    Else
    
        With ForSorterSheet
        
            .Range("A" & k & ":F" & k).NumberFormatLocal = "@"
            .Range("A" & k & ":F" & k) = Order
       
           '���ʁA���݌ɂ͉E�񂹁AJAN�͒���
            .Range("C" & k).HorizontalAlignment = xlRight
            .Range("D" & k).HorizontalAlignment = xlCenter
            .Range("F" & k).HorizontalAlignment = xlRight
       
            '�I�ԂȂ��́A�s�ɐF��t����B
            If OrderSheet.Range("K" & i).Value = "" Then
                     
                With .Range("A" & k & ":F" & k).Interior
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


Call SortHasLocation(ForSorterSheet)

ForSorterSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
ForSorterSetItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous


'�I������ Sheet���e�m��

'�O�̂��ߕ����Ďw��
Call AdjustWidth(ForSorterSheet)
Call AdjustWidth(ForSorterSetItemSheet)

ForSorterSheet.Protect
ForSorterSetItemSheet.Protect

End Sub

Sub OutputPickingData(ByVal MallName As String)

'�����̖��O�ŐV�K�u�b�N���쐬����
'�t�@�C������Amazon-�s�b�L���O�V�[�g�AYahoo=���t�[P�V�[�g�A�d�Z�����̏����̊֌W�ŌŒ�
Dim BookName As String
If MallName = "Amazon" Then
    BookName = "�s�b�L���O�V�[�g"
ElseIf MallName = "Yahoo" Then
    BookName = "���t�[P�V�[�g"
Else
    BookName = MallName & "P�V�[�g"
End If

'��o�p�t�@�C����p��
'100��/200�ԒI�L�� -2-3�A�d�Z����o
Dim ForSlimsBook As Workbook, ForSlimsSheet As Worksheet
Set ForSlimsBook = PreparePickingBook(BookName & Format(Date, "MMdd") & "-2-3")
Set ForSlimsSheet = ForSlimsBook.Worksheets(1)

'�o�^�����A�I���� -a
Dim NoEntryItemBook As Workbook, NoEntryItemSheet As Worksheet
Set NoEntryItemBook = PreparePickingBook(BookName & Format(Date, "MMdd") & "-a")
Set NoEntryItemSheet = NoEntryItemBook.Worksheets(1)

OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�I�����V�[�g�s�J�E���^
j = 3

'100�ԃV�[�g�s�J�E���^
k = 3

'�p�ӂ����u�b�N��1�s���R�s�[
Do

    '�����œn���ꂽ���[���ȊO�͔�΂�
    If Not Range("F" & i).Value Like (MallName & "*") Then GoTo Continue
    
    '�󒍎��R�[�h��7777�͓d�Z��o�f�[�^�Ɋ܂߂Ȃ��B
    If Range("I" & i).Value Like "7777*" Then GoTo Continue

    '��o����R�[�h�̐U��
    'SLIMS�ɓ�������̂̓��P�[�V�����L���6�P�^�̂�
    Dim OrderedCode As String, ForAddinCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("B" & i).Value)
    ForAddinCode = CStr(Range("I" & i).Value)
    AddinResultCode = CStr(Range("M" & i).Value)
    
    If AddinResultCode <> "" Then
        Code = AddinResultCode
    ElseIf ForAddinCode <> "" Then
        Code = ForAddinCode
    Else
        Code = OrderedCode
    End If
    
    '�z��ɒ�o�t�@�C��1�s���̃f�[�^���i�[
    '�A�}�]���̂݁A�d�Z�������ŃA�}�]�������ԍ��𔻒肵�Ă���A�A�ԕs��
    If MallName = "Amazon" Then
        Order(0) = CStr(Range("H" & i).Value) '���[�����̔Ԃ̒����ԍ�
    Else
        Order(0) = CStr(Range("A" & i).Value) '�N���X���[���̔Ԃ̘A��
    End If
    
    Order(1) = CStr(Code) '���i�R�[�h
    Order(2) = ValidateName(Range("C" & i).Value)  '���i�����ȂǍ폜���ē]�L
    Order(3) = Range("J" & i).Value '����
    Order(4) = Range("E" & i).Value '�̔����i
    Order(5) = Range("N" & i).Value '���݌�
    Order(6) = Range("K" & i).Value '�L�����P�[�V����
    
    '�]�L�攻��  �R�[�h�������͏����F������Ƃ��āA�擪�[�����J�b�g����Ȃ��悤��
    
    '���P�[�V�����Ȃ�
    If Order(6) = "" Then
        
        NoEntryItemSheet.Range("B" & j & ":C" & j).NumberFormatLocal = "@"
        NoEntryItemSheet.Range("B" & j & ":H" & j) = Order
    
        j = j + 1
    
    Else

        ForSlimsSheet.Range("B" & k & ":C" & k).NumberFormatLocal = "@"
        ForSlimsSheet.Range("B" & k & ":H" & k) = Order
       
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

'P�V�[�g�̃u�b�N�ۑ�����
'Amazon�̂ݑ����񂪕K�v�A������ 0�~ �Ŗ��߂�
ForSlimsBook.Activate
With ForSlimsSheet

    If MallName = "Amazon" Then
        .Columns("G").Insert
        .Range("G2").Value = "����"
        .Range(Cells(3, 7), Cells(ForSlimsSheet.UsedRange.Rows.Count, 7)).Value = 0
    End If

    '�r�������ĕۑ�
    .Range("A2:I2").Resize(Range("B2").CurrentRegion.Rows.Count - 1, 9).Borders.LineStyle = xlContinuous
    
End With
ForSlimsBook.Close SaveChanges:=True

NoEntryItemSheet.Activate
With NoEntryItemSheet
    .Activate
    
    If MallName = "Amazon" Then
        .Columns("G").Insert
        .Range("G1").Value = "����"
        .Range(Cells(2, 7), Cells(.UsedRange.Rows.Count, 7)).Value = 0
    End If
    
    '�r�������ĕۑ�
    .Range("A2:I2").Resize(Range("B2").CurrentRegion.Rows.Count - 1, 9).Borders.LineStyle = xlContinuous
    
End With
NoEntryItemBook.Close SaveChanges:=True

End Sub

Private Function PreparePickingBook(ByVal BookName As String) As Workbook
'�u�b�N����ς��邽�߂ɁA����̏ꏊ�֐�Ƀf�[�^�Ȃ��ŕۑ�����

Const PICKING_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\" '�Ō�A�K��\�}�[�N

ThisWorkbook.Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Copy
ActiveSheet.Name = BookName

'��U�V�K�쐬�u�b�N��ۑ����邱�ƂŃu�b�N����ύX����
'�V�K�쐬�t�@�C���̕ۑ����̓t�@�C���t�H�[�}�b�g�𖾎�����K�v�Ȗ͗l
Dim SavePath As String, SaveFolder As String

'�ۑ���ƕۑ��t�@�C�����̌���

'�l�b�g�̔��̃t�H���_�Ɍq���邩����
If Dir(PICKING_FOLDER, vbDirectory) <> "" Then
    SaveFolder = PICKING_FOLDER
Else
    SaveFolder = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\"
    MsgBox "�l�b�g�̔��֘A�Ɍq����Ȃ����߁A" & BookName & "���f�X�N�g�b�v�ɕۑ����܂��B"
End If

'�����y�E�v���C�����̃s�b�L���O���H
If Main.IsSecondPicking = True Then
    BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "AR"))
End If

'�����ۑ��̃t���O�����邩
If Main.IsTimeStampMode = True Then
    BookName = Replace(BookName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

'�t�@�C���������Ĥ�����y�v���C�����̃s�b�L���O�łȂ����̂ݤ�I���_�C�A���O��\��
If Dir(PICKING_FOLDER & BookName & ".xlsx") <> "" And Main.IsSecondPicking = False Then
    
    Dim IsAR As Integer
    IsAR = MsgBox(prompt:="�{�����̃t�@�C�������ɑ��݂��܂��B" & vbLf & "�����y�E�v���C�����Ƃ��ĕۑ����܂����H", _
            Buttons:=vbExclamation + vbYesNo)
    
    '�����y�v���C�����[�h�̃t���O�𗧂Ă�
    If IsAR = vbYes Then
        BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "AR"))
        Main.IsSecondPicking = True
    
    '�����y�v���C�����łȂ��Ƃ��A�������܂߂��t�@�C�����ۑ��t���O�𗧂Ă�
    Else
        BookName = Replace(BookName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
        Main.IsTimeStampMode = True
    
    End If

End If
    
'��L�����ɑS�ăq�b�g���Ȃ��ꍇ�́A����1��ڂ̐����ƂȂ�ABookName�͕ύX����Ă��Ȃ��B

If Dir(PICKING_FOLDER & BookName & ".xlsx") <> "" Then
    '�ۑ����悤�Ƃ���t�@�C�����ŁA���Ƀt�@�C��������ꍇ�t�@�C��������t-�����Ƃ���
    BookName = Replace(BookName, Format(Date, "MMdd"), (Format(Date, "MMdd") & "-" & Format(Time, "hhmm")))
End If

SavePath = SaveFolder & BookName

ActiveWorkbook.Sheets(1).Name = BookName
ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function

Private Sub AdjustWidth(TargetSheet As Worksheet)
'�� �������ɃA���[�g���o��̂�}�~
Application.DisplayAlerts = False

Dim WidthArr As Variant
WidthArr = Array(14.75, 84.13, 4.25, 15.88, 14.88, 6.63)

TargetSheet.Activate

Dim k As Long
For k = 0 To 5
    TargetSheet.Columns(k + 1).ColumnWidth = WidthArr(k)
Next

Application.DisplayAlerts = True

End Sub

Private Sub SortHasLocation(Sheet As Worksheet)

Dim SortRange As Range
Set SortRange = Sheet.Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Sheet.Range("A2:A" & SortRange.Rows.Count)

'�\�[�g�������Z�b�g
With Sheet.Sort
    
    '��U�\�[�g���N���A
    .SortFields.Clear
    
    '�\�[�g�L�[���Z�b�g ���L�[ ���i�R�[�h�F�F�A���L�[ ���i�R�[�h�F����
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    '�\�[�g�Ώۂ̃f�[�^�������Ă�͈͂��w�肵��
    .SetRange SortRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    '�Z�b�g����������K�p
    .Apply

End With

'�J�����g���[�W�������Z���N�g����Ă���̂ŁA�I���Z�����Z��A1�ɃZ�b�g������
Sheet.Activate
Range("A1").Select

End Sub
