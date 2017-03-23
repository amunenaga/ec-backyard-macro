Attribute VB_Name = "DataValidate"
Option Explicit

Sub FilterLocation(Optional ByVal arg As Boolean)
'アドインで取得したデータを修正する
'エクセルのマクロ一覧に出さないようにするため引数付きとしている。

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    'ロケーション修正、商品名バリデーション
    Cells(i, 11).Value = CutOffUnlocation(Cells(i, 15).Value)
    
Next

End Sub

Function CutOffUnlocation(Location As String) As String
' 正規表現でロケーション[0-0-0-0][0- -0- - ][1-0-0-0-0]などを削除して返します。

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "\[[0-9|\s]\-[0,1,2|\s]\-[0|\s]\-[0|\s]\-[0|\s]\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

Function ValidateName(Name As String) As String

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = ",|\!|\.|&"

Name = Reg.Replace(Name, "")


Reg.Pattern = "^((≪|【).*?(】|≫))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function
