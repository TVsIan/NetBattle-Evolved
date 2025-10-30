VERSION 5.00
Begin VB.Form DatabaseMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Per-Server Database Changes"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7080
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtDB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      HideSelection   =   0   'False
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6915
   End
   Begin VB.Label Label5 
      Caption         =   "move,Solrock,Rapid Spin       trait,Gengar,Clear Body,2     illegal,Blissey,Heal Bell,Softboiled"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   4395
   End
   Begin VB.Label Label4 
      Caption         =   "Examples:"
      Height          =   195
      Left            =   2220
      TabIndex        =   5
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "move,[pokémon],[attack]    trait,[pokémon],[trait],[slot] illegal,[pokémon],[move1],[move2],[...]"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   2820
      Width           =   4755
   End
   Begin VB.Label Label2 
      Caption         =   "Syntax:"
      Height          =   195
      Left            =   2220
      TabIndex        =   3
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   $"DatabaseMod.frx":0000
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   2580
      Width           =   1935
   End
End
Attribute VB_Name = "DatabaseMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If CompressScript Then
        ApplyDBMod
        Unload Me
    End If
End Sub

Private Function CompressScript() As Boolean
    Dim Build As String
    Dim Lines() As String
    Dim Words() As String
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Pos As Long
    Dim L As Long
    CompressScript = False
    Build = vbNullString
    
    Lines = Split(txtDB.Text, vbNewLine)
    For X = 0 To UBound(Lines)
        Lines(X) = LCase$(Lines(X))
        L = Len(Lines(X))
        If Len(Trim$(Lines(X))) <> 0 Then
            Words = Split(Lines(X), ",")
            If UBound(Words) < 1 Then GoTo CrapOut
            For Y = 0 To UBound(Words)
                If Len(Trim$(Words(Y))) = 0 Then GoTo CrapOut
            Next Y
                
            Select Case Trim$(Words(0))
            Case "move"
                If UBound(Words) <> 2 Then GoTo CrapOut
                Build = Build & "01"
                Y = GetPokeNum(Trim$(Words(1)))
                If Y = 0 Then SetError Pos, Words, 1, "Invalid Pokémon": Exit Function
                Build = Build & Dec2Bin(Y, 9)
                Y = GetMoveNum(Trim$(Words(2)))
                If Y = 0 Then SetError Pos, Words, 2, "Invalid Move": Exit Function
                Build = Build & Dec2Bin(Y, 9)
            Case "trait"
                If UBound(Words) <> 3 Then GoTo CrapOut
                Build = Build & "10"
                Y = GetPokeNum(Trim$(Words(1)))
                If Y = 0 Then SetError Pos, Words, 1, "Invalid Pokémon": Exit Function
                Build = Build & Dec2Bin(Y, 9)
                Y = GetTraitNum(Trim$(Words(2)))
                If Y = 0 Then SetError Pos, Words, 2, "Invalid Trait": Exit Function
                Build = Build & Dec2Bin(Y, 7)
                Select Case Trim$(Words(3))
                Case "1": Build = Build & "0"
                Case "2": Build = Build & "1"
                Case Else: SetError Pos, Words, 3, "Invalid Trait Number (1 or 2)": Exit Function
                End Select
            Case "illegal"
                If UBound(Words) < 2 Or UBound(Words) > 5 Then GoTo CrapOut
                Build = Build & "11"
                Y = GetPokeNum(Trim$(Words(1)))
                If Y = 0 Then SetError Pos, Words, 1, "Invalid Pokémon": Exit Function
                Build = Build & Dec2Bin(Y, 9)
                Build = Build & Dec2Bin(UBound(Words) - 2, 2)
                For Z = 2 To UBound(Words)
                    Y = GetMoveNum(Trim$(Words(Z)))
                    If Y = 0 Then SetError Pos, Words, 2, "Invalid Move": Exit Function
                    Build = Build & Dec2Bin(Y, 9)
                Next Z
            Case Else: GoTo CrapOut
            End Select
        End If
        Pos = Pos + L + 2
    Next X
    'Debug.Print Build
    'MsgBox "kansei"
    SaveSetting "NetBattle", "Server", "dbmod", Build
    DBModStr = Bin2Chr(Build)
    CompressScript = True
    Exit Function
    
CrapOut:
    txtDB.SelStart = Pos
    txtDB.SelLength = L
    txtDB.SetFocus
    MsgBox "Syntax Error"
End Function
Private Sub ExpandScript()
    Dim Build As String
    Dim X As Long
    Dim Text As String
    'On Error GoTo ExpandScript_Error
    Build = Chr2Bin(DBModStr)
    Do While Len(Build) > 2
        Select Case ChopString(Build, 2)
        Case "00"
            Exit Do
        Case "01"
            Text = Text & "Move, "
            Text = Text & BasePKMN(Bin2Dec(ChopString(Build, 9))).Name & ", "
            Text = Text & Moves(Bin2Dec(ChopString(Build, 9))).Name & vbNewLine
        Case "10"
            Text = Text & "Trait, "
            Text = Text & BasePKMN(Bin2Dec(ChopString(Build, 9))).Name & ", "
            Text = Text & AttributeText(Bin2Dec(ChopString(Build, 7))) & ", "
            Text = Text & CStr(CLng(ChopString(Build, 1)) + 1) & vbNewLine
            
        Case "11"
            Text = Text & "Illegal, "
            Text = Text & BasePKMN(Bin2Dec(ChopString(Build, 9))).Name
            For X = 0 To Bin2Dec(ChopString(Build, 2))
                Text = Text & ", " & Moves(Bin2Dec(ChopString(Build, 9))).Name
            Next X
            Text = Text & vbNewLine
        End Select
    Loop
    txtDB.Text = Left$(Text, Len(Text) - 2)
    
    Exit Sub

ExpandScript_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExpandScript of Form DatabaseMod"
End Sub
Private Sub SetError(ByVal Pos As Long, Words() As String, ByVal iWord As Long, ByVal Message As String)
    Dim X As Long
    Dim Y As Long
    Pos = Pos + Len(Words(iWord)) - Len(LTrim$(Words(iWord)))
    For X = 0 To iWord - 1
        Pos = Pos + Len(Words(X)) + 1
    Next X
    txtDB.SelStart = Pos
    txtDB.SelLength = Len(Trim$(Words(iWord)))
    txtDB.SetFocus
    MsgBox Message, vbExclamation, "Error"
End Sub

Private Sub Form_Load()
    If Len(DBModStr) <> 0 Then ExpandScript
End Sub
