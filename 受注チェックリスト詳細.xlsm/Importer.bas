Attribute VB_Name = "Importer"
Option Explicit
Sub �󒍃`�F�b�NCSV�Ǎ�()

GetOrderCheckListPath


End Sub
Private Function GetOrderCheckListPath(FileName As String) As String
'�s�b�L���O�V�[�g��-a���I���̃Z�b�g����O�t�@�C����T���ăt���p�X���Z�b�g

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\" '����\�}�[�N�K�{

'�y�V�̏ꍇ�A�y�VP�V�[�g0627-a.xls

'���s���o�C���f�B���O ScriptingRuntime��Dictionary�z��g���̂ɕK�v�ŎQ��ON������A���O�o�C���f�B���O�ł��������B
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, Newest As Object
      
'���O�o�C���f�B���O
'Dim FSO As FileSystemObject
'Set FSO = New FileSystemObject

'Dim f As File, Newest As File


'�w��t�H���_�[����FileName���܂ރt�@�C�����𒲂ׂāA�ŐV�̃t�@�C����1�擾����B
'LINQ�������A1�\���ōςނ̗~����

For Each f In FSO.GetFolder(SANTYOKU_DUMP_FOLDER).Files

    If f.Name Like FileName & ".csv" Then
    
        Set Newest = f
    
        Exit For
    End If

Next


RetrievePickingFilePath = PICKING_FILE_FOLDER & Newest.Name

End Function
