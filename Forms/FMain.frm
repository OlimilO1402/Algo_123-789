VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FMain"
   ClientHeight    =   2175
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text3 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      ToolTipText     =   "Doppelklick befördert die Zahlen wieder zurück in die Liste"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "Doppelklick befördert die Zahlen wieder zurück in die Liste"
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   5040
      TabIndex        =   1
      ToolTipText     =   "Doppelklick befördert die Zahlen in die TextBox"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStrVal As String
Private mCiphers As Ciphers
Private Type TValPair
    LngVal1 As Long
    LngVal2 As Long
End Type
Private mValPairs() As TValPair

Private Sub Form_Load()
    Me.Caption = "Algo 123-789 v." & App.Major & "." & App.Minor & "." & App.Revision
    Set mCiphers = New Ciphers
    mStrVal = "1 7 8 9 2 3"
    Text1.Text = mStrVal
    Label1.Caption = "1. Bilde aus den obigen Ziffern zwei Zahlen, so daß jede Ziffer nur einmal vorkommt." & vbCrLf & _
                     "Bsp: {1,2,3,7,8,9}: Zahl1: 723; Zahl2: 189"
    Label2.Caption = "2. Bilde die beiden Zahlen so, daß sich beim Subtrahieren der Zahlen eine möglichst kleine positive Differenz ergibt."
    Option1.Value = True
End Sub

Private Sub BtnRead_Click()
    mStrVal = Text1.Text
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    Label4.Caption = vbNullString
    Set mCiphers = New Ciphers
    With mCiphers
        If Not .Parse(mStrVal) Then
            MsgBox "Achtung parse war nicht erfolgreich"
        End If
        Call .Sort
        Call .ToListBox(List1)
        Label3.Caption = .ToString
    End With
End Sub

Private Sub List1_DblClick()
    Dim s As String: s = List1.List(List1.ListIndex)
    If Not Option1.Value And Not Option2.Value Then
        Option1.Value = True
    End If
    If Option1.Value Then
        Call TextBoxAppend(Text2, s)
        Call List1.RemoveItem(List1.ListIndex)
    ElseIf Option2.Value Then
        Call TextBoxAppend(Text3, s)
        Call List1.RemoveItem(List1.ListIndex)
    End If
End Sub
Private Sub TextBoxAppend(aTB As TextBox, s As String)
    aTB.Text = aTB.Text & s
End Sub
Private Function TBDiffToString(aTB1 As TextBox, aTB2 As TextBox) As String
    Dim d As Currency
    d = TBToCur(aTB1) - TBToCur(aTB2)
    If mCiphers.CiphersKind = dezimal Then
        TBDiffToString = CStr(d)
    ElseIf mCiphers.CiphersKind = hexadezimal Then
        TBDiffToString = Int64CurToHex(d)
    End If
End Function
Private Function TBToCur(aTB As TextBox) As Currency 'Long
    On Error Resume Next
    If Len(aTB.Text) Then
        If mCiphers.CiphersKind = dezimal Then
            TBToCur = CCur(aTB.Text)
        ElseIf mCiphers.CiphersKind = hexadezimal Then
            TBToCur = CCur("&H" & aTB.Text)
        End If
    End If
End Function
Private Sub Text2_Change()
    Label4.Caption = TBDiffToString(Text2, Text3)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 65 To 70, 97 To 102, vbKeyBack
    Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 65 To 70, 97 To 102
    Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub Text3_Change()
    Label4.Caption = TBDiffToString(Text2, Text3)
End Sub

Private Sub Text2_GotFocus()
    Option1.Value = True
End Sub
Private Sub Text3_GotFocus()
    Option2.Value = True
End Sub
Private Sub Text2_DblClick()
    Call MoveLastToListBox(List1, Text2)
End Sub
Private Sub Text3_DblClick()
    Call MoveLastToListBox(List1, Text3)
End Sub
Private Sub MoveLastToListBox(aLB As ListBox, aTB As TextBox)
    Dim t As String: t = aTB.Text
    If Len(t) > 0 Then
        Dim s As String: s = Right$(t, 1)
        Call aLB.AddItem(s)
        aTB.Text = Left$(t, Len(t) - 1)
    End If
End Sub

