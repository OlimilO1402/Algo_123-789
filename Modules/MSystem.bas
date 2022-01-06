Attribute VB_Name = "MSystem"
Option Explicit
Private Type TLngHiLo
    Lo As Long
    Hi As Long
End Type
Private Type TCur
    CurVal As Currency
End Type

Public Function Int64CurToHex(ByVal CurVal As Currency) As String
    Dim c As TCur:   c.CurVal = CurVal / 10000
    Dim l As TLngHiLo: LSet l = c
    With l
        If (.Hi) > 0 Then Int64CurToHex = Hex$(.Hi)
        Int64CurToHex = Int64CurToHex & Hex$(.Lo)
    End With
End Function

Public Function MaxL(ByVal Val1 As Long, ByVal Val2 As Long) As Long
    If Val1 > Val2 Then MaxL = Val1 Else MaxL = Val2
End Function
Public Function MinL(ByVal Val1 As Long, ByVal Val2 As Long) As Long
    If Val1 < Val2 Then MinL = Val1 Else MinL = Val2
End Function

'nur ein Zeichen:
'Public Function IsHexNumeric(astrval As String) As Boolean
'    Select Case Asc(strarr(i))
'    Case 48 To 57, 65 To 70, 97 To 102
'    Case Else
'        IsHexNumeric = False
'    End Select
'End Function
'oder der ganze String:
'Public Function IsHexNumeric(astrval As String) As Boolean
'    Dim l As Long
'Try1E: On Error GoTo Try2
'    l = CLng(astrval)
'    IsHexNumeric = True: On Error GoTo 0: Exit Function
'Try2: On Error GoTo 0
'Try2E: On Error GoTo CatchE
'    l = CLng("&H" & astrval)
'    IsHexNumeric = True
'CatchE: On Error GoTo 0
'End Function

'Public Function StrArrIsHexNumeric(ByRef strarr) As Boolean
'    Dim i As Long
'    StrArrIsHexNumeric = True
'    For i = 0 To UBound(strarr)
'        If Not IsHexNumeric(CStr(strarr(i))) Then
'            StrArrIsHexNumeric = False
'            Exit Function
'        End If
'    Next
'End Function

