Attribute VB_Name = "ConnectDB"
Sub Make_List(Optional ByVal arg As Boolean)
'�쐬�F���i��

'SQL�n�ϐ�
Dim DB_Cnn As New ADODB.Connection
Dim DB_Cmd  As New ADODB.Command
Dim DB_Rs As New ADODB.Recordset
Dim SQL_W1 As String

'���Ԋm�F�ϐ�
Dim Target_RowEnd As Integer
Dim Loop_Count As Integer
Dim A As Integer
Dim Target_Code As String
Dim Loc_Text As String

'�萔�Z�b�g
Const Target_RowBase = 2
Const Target_ColBase = 9
Const Output_RowBase = 2
Const OutPut_ColBase = 12

'�V�[�g�^�C�g���p�ϐ�
Dim S_HEAD(3)

'�V�[�g�^�C�g���Z�b�g
S_HEAD(0) = "JAN�R�[�h"
S_HEAD(1) = "���i�R�[�h"
S_HEAD(2) = "���݌ɐ�"
S_HEAD(3) = "���P�[�V����"


'Workbook���J���Ă��邩�m�F
A = 0
For Each wn In Workbooks
A = A + 1
Next
If A = 0 Then
MsgBox ("�V�[�g���J���Ă��������B")
End
End If

'SQL Server�ڑ�
DB_Cnn.ConnectionTimeout = 0
DB_Cnn.Open "PROVIDER=SQLOLEDB;Server=;Database=;UID=;PWD=;"
DB_Cmd.CommandTimeout = 180
Set DB_Cmd.ActiveConnection = DB_Cnn


'---�����J�n---
'�w�b�_�[�Z�b�g
For Loop_Count = 0 To 3
    Cells(1, 12 + Loop_Count).Select
    Cells(1, 12 + Loop_Count).Value = S_HEAD(Loop_Count)
Next Loop_Count

'�ŏIRow�擾
Cells(2, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Target_RowEnd = ActiveSheet.Cells.SpecialCells(xlLastCell).Row

'���ʍX�V�A�Čv�Z�}�~
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'���C�����[�v
Cells(Target_RowBase, Target_ColBase).Select
For Loop_Count = Target_RowBase To Target_RowEnd
    Target_Code = Cells(Loop_Count, Target_ColBase).Value
    
    '�R�[�h���ʁi�C���X�g�A�EJAN�j
    If Len(Target_Code) <= 6 Then
        SQL_W1 = "SELECT ���i�}�X�^.���i�R�[�h, ���i�}�X�^.JAN�R�[�h, Sum(�݌Ƀ}�X�^.���݌ɐ�) AS ���݌ɐ��v "
        SQL_W1 = SQL_W1 & "FROM ���i�}�X�^ INNER JOIN �݌Ƀ}�X�^ ON ���i�}�X�^.���i�R�[�h = �݌Ƀ}�X�^.���i�R�[�h "
        SQL_W1 = SQL_W1 & "GROUP BY ���i�}�X�^.���i�R�[�h, ���i�}�X�^.JAN�R�[�h "
        SQL_W1 = SQL_W1 & "HAVING (((���i�}�X�^.���i�R�[�h)=" & Target_Code & "));"
    Else
        SQL_W1 = "SELECT ���i�}�X�^.���i�R�[�h, ���i�}�X�^.JAN�R�[�h, Sum(�݌Ƀ}�X�^.���݌ɐ�) AS ���݌ɐ��v "
        SQL_W1 = SQL_W1 & "FROM ���i�}�X�^ INNER JOIN �݌Ƀ}�X�^ ON ���i�}�X�^.���i�R�[�h = �݌Ƀ}�X�^.���i�R�[�h "
        SQL_W1 = SQL_W1 & "GROUP BY ���i�}�X�^.���i�R�[�h, ���i�}�X�^.JAN�R�[�h "
        SQL_W1 = SQL_W1 & "HAVING (((���i�}�X�^.JAN�R�[�h)='" & Target_Code & "'));"
    End If
    
    Set DB_Rs = DB_Cnn.Execute(SQL_W1)

    If Not DB_Rs.EOF Then
        Cells(Loop_Count, OutPut_ColBase).Value = DB_Rs("JAN�R�[�h")
        Cells(Loop_Count, OutPut_ColBase + 1).NumberFormatLocal = "@"
        Cells(Loop_Count, OutPut_ColBase + 1).Value = Format(DB_Rs("���i�R�[�h"), "000000")
        Cells(Loop_Count, OutPut_ColBase + 2).Value = DB_Rs("���݌ɐ��v")
        
        '���P�[�V���������̎擾
        SQL_W1 = "SELECT �݌Ƀ}�X�^.���i�R�[�h,"
        SQL_W1 = SQL_W1 & "[�݌Ƀ}�X�^].[�K]+'-'+[�݌Ƀ}�X�^].[�ʘH]+'-'+[�݌Ƀ}�X�^].[�I��]+'-'+[�݌Ƀ}�X�^].[�i]+'-'+[�݌Ƀ}�X�^].[��] AS ���P�[�V���� "
        SQL_W1 = SQL_W1 & "FROM �݌Ƀ}�X�^ WHERE (�݌Ƀ}�X�^.���i�R�[�h=" & DB_Rs("���i�R�[�h") & ");"
        
        Set DB_Rs = DB_Cnn.Execute(SQL_W1)
        
        Loc_Text = ""
        Do While Not DB_Rs.EOF
            Loc_Text = Loc_Text & "[" & DB_Rs("���P�[�V����") & "]"
            DB_Rs.MoveNext
        Loop
        
        Cells(Loop_Count, OutPut_ColBase + 3).Value = Loc_Text
        
    Else
        '�݌Ƀ}�X�^�[�ɓo�^���Ȃ��ꍇ�A���i�}�X�^���珤�i�R�[�h��JAN�̂ݎ擾����
        
        '�R�[�h���ʁi�C���X�g�A�EJAN�j-> WHERE���Z�b�g DB�ŃR�[�h�͐��l�^�AJAN�̓e�L�X�g�^
        Dim Clause_WHERE As String
        Clause_WHERE = IIf(Len(Target_Code) <= 6, "���i�}�X�^.���i�R�[�h = " & Target_Code, "���i�}�X�^.JAN�R�[�h = '" & Target_Code & "'")
    
        SQL_W1 = "SELECT ���i�}�X�^.���i�R�[�h, ���i�}�X�^.JAN�R�[�h "
        SQL_W1 = SQL_W1 & "FROM ���i�}�X�^ "
        SQL_W1 = SQL_W1 & "WHERE " & Clause_WHERE
        
        'SQL���s���ďo��
        Set DB_Rs = DB_Cnn.Execute(SQL_W1)
        
        If Not DB_Rs.EOF Then
            Cells(Loop_Count, OutPut_ColBase).Value = DB_Rs("JAN�R�[�h")
            
            Cells(Loop_Count, OutPut_ColBase + 1).NumberFormatLocal = "@"
            Cells(Loop_Count, OutPut_ColBase + 1).Value = Format(DB_Rs("���i�R�[�h"), "000000")
        End If
        
    End If
Next Loop_Count
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'Cells.Select
    'Cells.EntireColumn.AutoFit
    Range("A1").Select

DB_Cnn.Close

Set DB_Rs = Nothing
Set DB_Cnn = Nothing
Set DB_Cmd = Nothing


End Sub


