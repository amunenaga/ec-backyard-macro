Attribute VB_Name = "LoadPurchaseReq"
Option Explicit

Sub LoadAllPicking(Optional ByRef TargetFolder As String)
'��z�˗��`�F�b�N�ς̃s�b�L���O�t�@�C�����ꊇ���ēǍ�
'��z�˗��Ƃ��Ĕw�i�F���ς��Ă���s���R�s�[���܂��B

Dim PickingFilePaths() As String
PickingFilePaths = SearchPickingFiles(TargetFolder)

'�Z���[���s�b�L���O�t�@�C���ǂݍ���
Dim Path As Variant
For Each Path In Array(PickingFilePaths(0), PickingFilePaths(1), PickingFilePaths(2))
    Call LoadSellerPicking(Path)
Next

'�����s�b�L���O�t�@�C���ǂݍ���
For Each Path In Array(PickingFilePaths(3), PickingFilePaths(4))
    Call LoadPoFile(Path)
Next

Call ApendSpToPurchseReq
Call VerifySyokonRegist

End Sub

Private Sub LoadSellerPicking(ByVal PickingFilePath As String)
'�Z���[���̃s�b�L���O�t�@�C���ǂݍ���

Dim Mall As String

'�s�b�L���O�V�[�g�����烂�[���L�����Z�b�g
Select Case True
    Case PickingFilePath Like "*�s�b�L���O�V�[�g*"
        Mall = "A"
    Case PickingFilePath Like "*�y�V*"
        Mall = "R"
    Case PickingFilePath Like "*���t�[*"
        Mall = "Y"
    Case Else
        Mall = "SP"
End Select

'�s�b�L���O�V�[�g�u�b�N���J��
On Error Resume Next
    
    Workbooks.Open FileName:=PickingFilePath
    If Err Then Exit Sub

On Error GoTo 0

Dim NoLocationSheet As Worksheet
Set NoLocationSheet = ActiveSheet


'�J���Ă���s�b�L���O�V�[�g����A�w�i�F�𔻒肵��1�s���f�[�^�R�s�[
Dim WriteRow As Long, i As Long
WriteRow = IIf(PurchaseReqSeller.Range("A2").Value = "", 2, PurchaseReqSeller.Range("A1").End(xlDown).Row + 1)

For i = 3 To NoLocationSheet.Range("A1").SpecialCells(xlLastCell).Row
    
    If NoLocationSheet.Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
        
        '�s�b�L���O-a�̔w�i���łȂ��s����U�R�s�[
        NoLocationSheet.Range(Cells(i, 2), Cells(i, 5)).Copy
        '�l�œ\��t��
        PurchaseReqSeller.Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
        
        PurchaseReqSeller.Cells(WriteRow, 1).Value = Mall
        
        WriteRow = WriteRow + 1
        
    End If

Next

ActiveWorkbook.Close SaveChanges:=False

End Sub
Private Sub LoadPoFile(ByVal PickingFilePath As String)
'Amazon���̃s�b�L���O�t�@�C���ǂݍ���

'�s�b�L���O�V�[�g�u�b�N���J��
On Error Resume Next

    Workbooks.Open FileName:=PickingFilePath
    If Err Then Exit Sub

On Error GoTo 0

Dim NoLocationSheet As Worksheet
Set NoLocationSheet = ActiveSheet

'�J���Ă���s�b�L���O�V�[�g����A��z�˗��Ǎ��V�[�g�փf�[�^�R�s�[

Dim WriteRow As Long, i As Long
WriteRow = IIf(PurchaseReqWholesall.Range("A2").Value = "", 2, PurchaseReqWholesall.Range("A1").End(xlDown).Row + 1)

For i = 2 To ActiveSheet.Range("A1").SpecialCells(xlLastCell).Row
    
    If Cells(i, 2).Interior.Color <> RGB(255, 255, 255) Then
        
        'PO��JAN���R�s�[�E�\��t��
        NoLocationSheet.Range(Cells(i, 1), Cells(i, 2)).Copy
        PurchaseReqWholesall.Cells(WriteRow, 2).PasteSpecial Paste:=xlPasteValues
        
        '���i��
        NoLocationSheet.Cells(i, 5).Copy
        PurchaseReqWholesall.Cells(WriteRow, 4).PasteSpecial Paste:=xlPasteValues
        
        '����
        NoLocationSheet.Cells(i, 9).Copy
        PurchaseReqWholesall.Cells(WriteRow, 5).PasteSpecial Paste:=xlPasteValues
        
        PurchaseReqWholesall.Cells(WriteRow, 1).Value = "V"
        
        WriteRow = WriteRow + 1
    End If
Next

ActiveWorkbook.Close SaveChanges:=False

End Sub

Private Sub ApendSpToPurchseReq()
'����͕��������V�[�g�A�Z���[���V�[�g�֐U�蕪���ăR�s�[
'180�ԁA187�Ԃŏ��i�ʂɐ��ʂ����Z���邽��

With Worksheets("����͕�")

    If IsEmpty(.Range("B2").Value) Then
        Exit Sub
    Else
        Dim CodeRange As Range
        Set CodeRange = .Range(.Cells(2, 2), .Cells(1, 2).End(xlDown))
    End If
    
End With

Dim r As Range, MallTicker As String

For Each r In CodeRange
    MallTicker = r.Offset(0, -1).Value
    
    If MallTicker Like "*[V|v]*" Then
    
        With Worksheets("����")
            .Range("A1").End(xlDown).Offset(1, 0).Value = "V"
            .Range("C1").End(xlDown).Offset(1, 0).NumberFormatLocal = "@"
            .Range("C1").End(xlDown).Offset(1, 0).Resize(1, 3).Value = r.Resize(1, 3).Value
        End With
        
    Else
    
        With Worksheets("�Z���[��")
            .Range("A1").End(xlDown).Offset(1, 0).Value = "SP"
            .Range("C1").End(xlDown).Offset(1, 0).NumberFormatLocal = "@"
            .Range("C1").End(xlDown).Offset(1, 0).Resize(1, 3).Value = r.Resize(1, 3).Value
        End With
    
    End If

Next

End Sub

Private Sub VerifySyokonRegist()

'�ڑ��̂��߂̃I�u�W�F�N�g���`�ADB�ڑ��ݒ���Z�b�g
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'�Z���[���͎���͂���JAN�����Ǔo�^�L�肪�Ȃ����𒲂ׂ�΂悢�̂ŁA�ŏI�s�����֒��ׂĂ���
With PurchaseReqSeller
    Dim EndRow As Long
    EndRow = PurchaseReqSeller.Range("A1").End(xlDown).Row
    
    Dim i As Long
    For i = EndRow To 2 Step -1
        
        If .Cells(i, 1).Value <> "SP" Then Exit For
        
        Dim Code As String
        Code = .Cells(i, 3).Value
        
        If Not .Cells(i, 3).Value Like String(13, "#") Then GoTo Continue
    
        '�N�G����6�P�^�擾
        Dim Sql As String
            
        Sql = "SELECT ���i�R�[�h FROM ���i�}�X�^ WHERE JAN�R�[�h = '" & Code & "'"
        
        Set DbRs = DbCnn.Execute(Sql)
    
        If Not DbRs.EOF Then
            .Cells(i, 3).NumberFormatLocal = "@"
            .Cells(i, 3).Value = IIf(Len(DbRs("���i�R�[�h")) = 5, "0" & CStr(DbRs("���i�R�[�h")), CStr(DbRs("���i�R�[�h")))
        End If
    
Continue:
    Next

End With

'�����́A�S�����ׂ�
With PurchaseReqWholesall
    EndRow = Range("A1").End(xlDown).Row
    
    For i = 2 To EndRow
        
        If .Cells(i, 3).Value = "" Then Exit For
        
        Code = .Cells(i, 3).Value
        
        If Not .Cells(i, 3).Value Like String(13, "#") Then GoTo Continue2

        Sql = "SELECT ���i�R�[�h FROM ���i�}�X�^ WHERE JAN�R�[�h = '" & Code & "'"
        
        Set DbRs = DbCnn.Execute(Sql)
    
        If Not DbRs.EOF Then
            .Cells(i, 3).NumberFormatLocal = "@"
            .Cells(i, 3).Value = IIf(Len(DbRs("���i�R�[�h")) = 5, "0" & CStr(DbRs("���i�R�[�h")), CStr(DbRs("���i�R�[�h")))
        End If
    
Continue2:
    Next

End With

End Sub

Private Function SearchPickingFiles(Optional ByRef FolderPath As String) As String()
'�t�H���_�w������ɁA�s�b�L���O�t�@�C���̃p�X��z��ŕԂ��܂��B

'PickingFiles(0) : Amazon�Z���[
'PickingFiles(1) : �y�V
'PickingFiles(2) : ���t�[
'PickingFiles(3) : Amazon��
'PickingFiles(4) : Amazon���A�E�g�h�A�J�e�S��

Dim Fso As New FileSystemObject, PickingFolder As Folder, File As File

Set PickingFolder = Fso.GetFolder(FolderPath)
Dim PickingFiles(4) As String

For Each File In PickingFolder.Files

    Select Case True
        Case File.Name Like "�s�b�L���O�V�[�g*-a*"
            PickingFiles(0) = FolderPath & "\" & File.Name
            
        Case File.Name Like "�y�V*-a*"
            PickingFiles(1) = FolderPath & "\" & File.Name
            
        Case File.Name Like "���t�[*-a*"
            PickingFiles(2) = FolderPath & "\" & File.Name
        
        Case File.Name Like "�A�}�]���I�Ȃ�####.xlsx"
            PickingFiles(3) = FolderPath & "\" & File.Name
            
        Case File.Name Like "�A�}�]���I�Ȃ�*-outdoor*"
            PickingFiles(4) = FolderPath & "\" & File.Name
    
    End Select
        
Next

SearchPickingFiles = PickingFiles

End Function
Private Sub TestGetPickingFiles()

Dim Files As Variant, File As Variant

Files = GetPickingFiles(PICKING_FOLDER)

For Each File In Files

    Debug.Print File

    If File Like "*�s�b�L���O�V�[�g*" Then
        Debug.Print "Amazon OK"
    ElseIf File Like "*�y�V*" Then
        Debug.Print "�y�V OK"
    ElseIf File Like "*���t�[*" Then
        Debug.Print "���t�[ OK"
    End If
Next

End Sub
