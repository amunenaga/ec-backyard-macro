Dim QtyBooks(2)
QtyBooks(0) = "�݌Ƀt�@�C��01.xlsm"
QtyBooks(1) = "�݌Ƀt�@�C��02.xlsm"
QtyBooks(2) = "�݌Ƀt�@�C��03.xlsm"

Dim objPath
Set objPath = CreateObject("Scripting.FileSystemObject").GetFolder(".")

Dim exApp
Set exApp = Wscript.CreateObject("Excel.Application")
exApp.Visible = True

For i =0 to 2

	Dim wb
	set wb = exApp.Workbooks.Open (objPath & "\" & QtyBooks(i))
	exApp.Run ("CSV����")

	wb.Close True

Next

exApp.Quit
Set exApp = Nothing
