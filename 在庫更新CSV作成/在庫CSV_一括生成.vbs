'�t�@�C�������i�[����z��
Dim QtyBooks(1)
QtyBooks(0) = "�݌ɕ\1.xlsm"
QtyBooks(1) = "�݌ɕ\2.xlsm"

'�t�@�C������ADir�擾FSO�I�u�W�F�N�g
Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")

dim CurrentPath
CurrentPath = Fso.GetFolder(".")

dim Mon,Today
Mon = Mid(date,6,2)

Today = Right(date,2)

'CSV�ƃw�b�_�[��p��
Dim QtyCsv
Set QtyCsv = Fso.CreateTextFile(CurrentPath & "\" & "���t�[�݌ɍX�V" & Mon & Today & ".csv")
QtyCsv.WriteLine("code,quantity,allow-overdraft")

QtyCsv.Close

'�eExcel��CSV�ǋL�}�N�����Ăяo��
Dim exApp
Set exApp = Wscript.CreateObject("Excel.Application")
exApp.Visible = True

For i = 0 to 1

	Dim wb
	set wb = exApp.Workbooks.Open (CurrentPath & "\" & QtyBooks(i))
	exApp.Run ("CSV����")

	wb.Close True

Next

exApp.Quit
Set exApp = Nothing
