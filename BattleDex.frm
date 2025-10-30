VERSION 5.00
Begin VB.Form BattleDex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BattleDex"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   9240
   Icon            =   "BattleDex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dual &Type..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Not ready yet..."
      Top             =   5760
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   6720
      ScaleHeight     =   975
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do 2x damage"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do 1/2 damage"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack will do no damage"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   60
         Width           =   2055
      End
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   2
         Left            =   60
         Top             =   660
         Width           =   240
      End
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   1
         Left            =   60
         Top             =   360
         Width           =   240
      End
      Begin VB.Image DispImg 
         Height          =   240
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6555
      Left            =   120
      ScaleHeight     =   6495
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Defender"
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Attacker"
         Top             =   180
         Width           =   135
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   6480
         X2              =   6480
         Y1              =   0
         Y2              =   6540
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "White rows are normal attacks, red rows are special attacks."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print..."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close BattleDex"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Web Site..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About..."
         Index           =   2
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "BattleDex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer

    Dim CurrentIcon As Integer
    '>>> Call WriteDebugLog("Battle Dex loaded")
    CenterWindow Me
    '>>> Call WriteDebugLog("Centered")
    
    DispImg(0) = MainContainer.Conditions.ListImages(10).Picture
    DispImg(1) = MainContainer.Conditions.ListImages(9).Picture
    DispImg(2) = MainContainer.Conditions.ListImages(1).Picture
    Z = 360 '(Picture1.Height - 100) \ 18
    
    Picture1.Line (0, Z * 2)-(Picture1.Width, Z * 7), Picture1.FillColor, BF
    Picture1.Line (0, Z * 11)-(Picture1.Width, Z * 12), Picture1.FillColor, BF
    Picture1.Line (0, Z * 15)-(Picture1.Width, Z * 17), Picture1.FillColor, BF
    For X = 1 To 18
        Picture1.Line (X * Z, 0)-(X * Z, Picture1.Height), vbBlack
        Picture1.Line (0, X * Z)-(Picture1.Width, X * Z), vbBlack
    Next X
    Picture1.Line (0, 0)-(Z, Z), vbBlack
    For X = 1 To 17
        Picture1.PaintPicture MainContainer.Types.ListImages(X).Picture, X * Z + 60, 60
        Picture1.PaintPicture MainContainer.Types.ListImages(X).Picture, 60, X * Z + 60
        For Y = 1 To 17
            Select Case BattleMatrix(X, Y)
            Case 0
                Picture1.PaintPicture MainContainer.Conditions.ListImages(10).Picture, Y * Z + 60, X * Z + 60
            Case 0.5
                Picture1.PaintPicture MainContainer.Conditions.ListImages(9).Picture, Y * Z + 60, X * Z + 60
            Case 2
                Picture1.PaintPicture MainContainer.Conditions.ListImages(1).Picture, Y * Z + 60, X * Z + 60
            End Select
        Next Y
    Next
    '>>> Call WriteDebugLog("Images Set")
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0
            'Insert print code here
        Case 2
            Unload Me
        Case 3
            End
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            MainContainer.GoExplorer "http://www.tvsian.com/netbattle"
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim Z As Integer
    X1 = X \ 360
    Y1 = Y \ 360
    If X1 > 17 Or Y1 > 17 Then
        Picture1.ToolTipText = ""
        Exit Sub
    End If
    If X1 = 0 Or Y1 = 0 Then
        If X1 = 0 And Y1 = 0 Then
            Picture1.ToolTipText = ""
        ElseIf X1 = 0 Then
            Picture1.ToolTipText = Element(Y1)
        Else
            Picture1.ToolTipText = Element(X1)
        End If
    Else
        Select Case BattleMatrix(Y1, X1)
        Case 0
            Picture1.ToolTipText = Element(X1) & " types are immune to " & Element(Y1) & " attacks"
        Case 0.5
            Picture1.ToolTipText = Element(X1) & " types are resistant to " & Element(Y1) & " attacks"
        Case 1
            Picture1.ToolTipText = ""
        Case 2
            Picture1.ToolTipText = Element(X1) & " types are weak to " & Element(Y1) & " attacks"
        End Select
    End If
End Sub



