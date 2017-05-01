Attribute VB_Name = "TestRegExp"
Option Explicit

Sub TestCutoffUnlocation()

Dim UnlocaitonStrings As Variant, LocString As Variant, TestResult As Boolean

TestResult = True

UnlocaitonStrings = Array( _
                            "[4-0- -0-0]", _
                            "[5-0- -0-0]", _
                            "[2-0- -0-0]", _
                            "[3-0- -0-5]", _
                            "[0-0-0-0-0]", _
                            "[ -1- -0-0]", _
                            "[ -0- -0-0]", _
                            "[ -0-2-0-0]", _
                            "[ - - - - ]", _
                            "[1-0- -0-15]", _
                            "[3-15-1-2-6]" _
                        )

For Each LocString In UnlocaitonStrings

    If CutOffUnlocation(CStr(LocString)) <> "" Then
        TestResult = False
        Debug.Print "Miss! ;" & LocString
    End If

Next

Dim ValidLocationStrings As Variant
ValidLocationStrings = Array( _
                            "[3-14-I-4-6]", _
                            "[1-6-R-4-3]", _
                            "[9-55-A-2-3-9]" _
                            )

For Each LocString In ValidLocationStrings

    If CutOffUnlocation(CStr(LocString)) = "" Then
        TestResult = False
        Debug.Print "Don't Cut! ;" & LocString
    End If

Next

If TestResult = True Then
    Debug.Print "Test Passed!"
Else
    Debug.Print "Test Missed"
End If

End Sub

Function CutOffUnlocation(Location As String) As String
'���K�\���Ń��P�[�V����[0-0-0-0][0- -0- - ][1-0-0-0-0]�Ȃǂ��폜���ĕԂ��܂��B

Dim Reg As New RegExp

Reg.Global = True

'���P�[�V�����̕��� �K-�ʘH-�I��-�i-��  �I�Ԃ�A�`Q�A���t�@�x�b�g
Reg.Pattern = "\[[0-9|\s]\-[0-2|\s]\-[0-9|\s]\-[0|\s]\-(([0-9]{1,})|\s)\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

