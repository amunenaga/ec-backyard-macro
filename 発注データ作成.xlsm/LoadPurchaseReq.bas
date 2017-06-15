Attribute VB_Name = "LoadPurchaseReq"
Option Explicit

Const PICKING_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\"

Sub LoadAllPicking()
'��z�˗��`�F�b�N�ς̃s�b�L���O�t�@�C�����ꊇ���ēǍ�
'��z�˗��Ƃ��Ĕw�i�F���ς��Ă���s���R�s�[���܂��B

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'�Z���[���s�b�L���O�t�@�C���ǂݍ���
Dim PickingFiles As Variant, File As Variant

PickingFiles = Array( _
    "�s�b�L���O�V�[�g", _
    "�y�VP�V�[�g", _
    "���t�[P�V�[�g" _
    )

For Each File In PickingFiles
    Call LoadSellerPicking(CStr(File) & Format(Date, "MMdd") & "-a.xlsx")
Next

'���� �t�@�C���ǂݍ���
PickingFiles = Array( _
    "�A�}�]���I�Ȃ�" & Format(Date, "MMdd") & ".xlsx", _
    "�A�}�]���I�Ȃ�" & Format(Date, "MMdd") & "-outdoor.xlsx" _
    )
    
For Each File In PickingFiles
    Call LoadPoFile(CStr(File))
Next
End Sub

Sub LoadSellerPicking(ByVal FileName As String)
'�Z���[���̃s�b�L���O�t�@�C���ǂݍ���

Dim Mall As String, PickingFileName As String

'�s�b�L���O�V�[�g�����烂�[���L�����Z�b�g
Select Case True
    Case FileName Like "�s�b�L���O*"
        Mall = "A"
    Case FileName Like "�y�V*"
        Mall = "R"
    Case FileName Like "���t�[*"
        Mall = "Y"
    Case Else
        Mall = "SP"
End Select

'�s�b�L���O�V�[�g�u�b�N���J���A�A�N�e�B�u�Ȃ܂܎g��
On Error Resume Next
    
    Workbooks.Open FileName:=PICKING_FOLDER & FileName
    If Err Then Exit Sub

On Error GoTo 0


'�J���Ă���s�b�L���O�V�[�g����A��z�˗��Ǎ��V�[�g�փf�[�^�R�s�[
With ThisWorkbook.Worksheets("�Z���[��")
    Dim WriteRow As Long, i As Long
    WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
    
    For i = 3 To ActiveSheet.UsedRange.Rows.Count
        
        If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
            
            '�w�i���łȂ��s����U�R�s�[
            Range(Cells(i, 2), Cells(i, 5)).Copy
            '�l�œ\��t��
            .Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
            
            .Cells(WriteRow, 1).Value = Mall
            
            WriteRow = WriteRow + 1
        End If
    Next
End With

ActiveWorkbook.Close Savechanges:=False


End Sub
Sub LoadPoFile(ByVal FileName As String)
'Amazon���̃s�b�L���O�t�@�C���ǂݍ���

'�s�b�L���O�V�[�g�u�b�N���J���A�A�N�e�B�u�Ȃ܂܎g��
On Error Resume Next
    Workbooks.Open FileName:=PICKING_FOLDER & FileName
    If Err Then Exit Sub

On Error GoTo 0


'�J���Ă���s�b�L���O�V�[�g����A��z�˗��Ǎ��V�[�g�փf�[�^�R�s�[
With ThisWorkbook.Worksheets("����")
    Dim WriteRow As Long, i As Long
    WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
    
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        
        If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
            
            'PO��JAN���R�s�[�E�\��t��
            Range(Cells(i, 1), Cells(i, 2)).Copy
            .Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
            
            '���i��
            Cells(i, 5).Copy
            .Cells(WriteRow, 4).PasteSpecial Paste:=xlPasteValues
            
            '����
            Cells(i, 9).Copy
            .Cells(WriteRow, 5).PasteSpecial Paste:=xlPasteValues
            
            .Cells(WriteRow, 1).Value = "V"
            
            WriteRow = WriteRow + 1
        End If
    Next
End With

ActiveWorkbook.Close Savechanges:=False

End Sub
