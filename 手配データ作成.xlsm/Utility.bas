Attribute VB_Name = "Utility"
Option Explicit

Sub PrepareSheet(ByRef TargetSheet As Variant)

If Not IsEmpty(TargetSheet.Range("A2").Value) Then
    TargetSheet.Range("A1").CurrentRegion.Offset(1, 0).Delete Shift:=xlShiftUp
End If

End Sub

Sub ClosePurDataBook()

End Sub
