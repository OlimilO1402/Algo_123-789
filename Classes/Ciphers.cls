VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ciphers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'speichert nur die Werte der Ziffern
Private mCiphers() As Byte
'es dürfen auch hexadezimale Ziffern sein
Public Enum HexDez
    dezimal
    hexadezimal
    alphanumerisch
End Enum
Private mCiphersKind As HexDez

Public Function Parse(astrval As String, Optional delim As String = " ") As Boolean
    Dim i As Long, n As Long, l As Long
    Dim sH As String, s As String
    Dim VA: VA = Split(astrval, delim)
    ReDim mCiphers(0 To UBound(VA))
    Parse = True
    For i = 0 To UBound(VA)
        s = VA(i)
        If Len(s) > 1 Then
            Parse = False
        ElseIf Len(s) = 1 Then
            mCiphersKind = MaxL(mCiphersKind, CipherKind(AscW(s)))
            If mCiphersKind = alphanumerisch Then
                mCiphers(n) = AscW(s)
                n = n + 1
                Parse = False
            Else
                l = CLng("&H" & VA(i))
                mCiphers(n) = CLng("&H" & VA(i))
                n = n + 1
            End If
        End If
    Next
    If n < UBound(mCiphers) + 1 Then
        ReDim Preserve mCiphers(0 To n - 1)
    End If
End Function

Private Property Get CipherKind(ByVal Char As Integer) As HexDez
    Select Case Char
    Case 48 To 57
        CipherKind = dezimal
    Case 65 To 70, 97 To 102
        CipherKind = hexadezimal
    Case Else
        CipherKind = alphanumerisch
    End Select
End Property

Public Property Get CiphersKind() As HexDez
    CiphersKind = mCiphersKind
End Property

Private Function HexDezToString(enmval As HexDez) As String
    Select Case enmval
    Case dezimal:        HexDezToString = "dezimal"
    Case hexadezimal:    HexDezToString = "hexadezimal"
    Case alphanumerisch: HexDezToString = "alphanumerisch"
    End Select
End Function

Public Sub Sort() 'a simple bubbler
    Dim i As Long, p As Long
    Dim l As Long: l = LBound(mCiphers)
    Dim u As Long: u = UBound(mCiphers)
    While u > l
        p = l
        For i = l To u - 1
            If CompareByt(mCiphers(i), mCiphers(i + 1)) > 0 Then
                Call Swap(mCiphers(), i, i + 1)
                p = i
            End If
        Next
        u = p
    Wend
End Sub

Private Function CompareByt(ByVal b1 As Byte, ByVal b2 As Byte) As Long
    If b1 > b2 Then
        CompareByt = 1
    ElseIf b1 < b2 Then
        CompareByt = -1
    End If
End Function

Private Sub Swap(aBytArr() As Byte, ByVal i1 As Long, ByVal i2 As Long)
    Dim temp As Byte: temp = aBytArr(i1)
    aBytArr(i1) = aBytArr(i2): aBytArr(i2) = temp
End Sub

Public Function ToString() As String
    Dim s As String: s = "{"
    Dim delim As String: delim = ", "
    Dim i As Long, n As Long: n = UBound(mCiphers)
    Dim c As Integer
    For i = LBound(mCiphers) To n
        s = s & CipherToString(mCiphers(i))
        If i < n Then
            s = s & delim
        End If
    Next
    s = s & "}: " & HexDezToString(mCiphersKind) & vbCrLf
    's = s &
    ToString = s
End Function

Private Function CipherToString(ByVal c As Integer) As String
    If c > &HF Then
        CipherToString = ChrW(c)
    Else
        CipherToString = Hex$(c)
    End If
End Function

Public Sub ToListBox(aLB As ListBox)
    Dim i As Long
    aLB.Clear
    For i = 0 To UBound(mCiphers)
        Call aLB.AddItem(CipherToString(mCiphers(i)))
    Next
End Sub


