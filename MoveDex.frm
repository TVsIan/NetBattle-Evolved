VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form MoveDex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MoveDex"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4335
   Icon            =   "MoveDex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4335
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4095
      Begin VB.Label Label8 
         Caption         =   "Hits All"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         ToolTipText     =   "In 2v2, hits all other Pokemon"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Image HAStat 
         Height          =   240
         Left            =   2040
         Picture         =   "MoveDex.frx":1272
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Hits Both"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         ToolTipText     =   "In 2v2, hits both opponents"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Image HBStat 
         Height          =   240
         Left            =   240
         Picture         =   "MoveDex.frx":17FC
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Contact Move"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "Rough Skin, Static, etc. will affect this move"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image CMStat 
         Height          =   240
         Left            =   2040
         Picture         =   "MoveDex.frx":1D86
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Sound Move"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         ToolTipText     =   "Move is blocked by Soundproof"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image SndStat 
         Height          =   240
         Left            =   240
         Picture         =   "MoveDex.frx":2310
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Substitute Blocked"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         ToolTipText     =   "Substitute will prevent this move from working properly"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image SBStat 
         Height          =   240
         Left            =   2040
         Picture         =   "MoveDex.frx":289A
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Self-Affecting"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         ToolTipText     =   "This move does not affect the opponent"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image SMStat 
         Height          =   240
         Left            =   240
         Picture         =   "MoveDex.frx":2E24
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "King's Rock"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "King's Rock gives this move a flinching bonus"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image KRStat 
         Height          =   240
         Left            =   2040
         Picture         =   "MoveDex.frx":33AE
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Bright Powder"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         ToolTipText     =   "Bright Powder may cause this move to miss"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image BPStat 
         Height          =   240
         Left            =   240
         Picture         =   "MoveDex.frx":3938
         Top             =   240
         Width           =   240
      End
   End
   Begin CCRProgressBar6.ccrpProgressBar AttackBar 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveDex.frx":3EC2
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveDex.frx":3FD4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveDex.frx":40E6
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveDex.frx":41F8
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MoveDex.frx":430A
            Key             =   "Camera"
         EndProperty
      EndProperty
   End
   Begin CCRProgressBar6.ccrpProgressBar AccBar 
      Height          =   255
      Left            =   2280
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CCRProgressBar6.ccrpProgressBar PPBar 
      Height          =   255
      Left            =   120
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   1
   End
   Begin VB.Label ADVTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label GSCTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label RBYTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image GBall 
      Height          =   240
      Left            =   5160
      Picture         =   "MoveDex.frx":441C
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image Ball 
      Height          =   240
      Left            =   4920
      Picture         =   "MoveDex.frx":49A6
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image RuSaCompat 
      Height          =   480
      Left            =   3720
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image GSCompat 
      Height          =   480
      Left            =   3000
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Desc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Image RBCompat 
      Height          =   480
      Left            =   2280
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label PPLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "PP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label AccLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Accuracy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label AttackLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Power:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label MoveName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Image TypeImg 
      Height          =   240
      Left            =   120
      Top             =   600
      Width           =   240
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
         Caption         =   "&Close MoveDex"
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
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MoveDex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentMove As Integer
Dim MoveSort() As Integer
Dim SortSwap() As String

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    
    CurrentMove = 1
    ReDim MoveSort(UBound(Moves)) As Integer
    ReDim SortSwap(UBound(Moves)) As String
    For X = 1 To UBound(Moves)
        If Moves(X).Power > AttackBar.Max Then AttackBar.Max = Moves(X).Power
        If Moves(X).Accuracy > AccBar.Max Then AccBar.Max = Moves(X).Accuracy
        If Moves(X).PP > PPBar.Max Then PPBar.Max = Moves(X).PP
        SortSwap(X) = Moves(X).Name
    Next
    Call SortStringArray(SortSwap())
    For X = 1 To UBound(Moves)
        For Y = 1 To UBound(Moves)
            If Moves(Y).Name = SortSwap(X) Then
                MoveSort(X) = Y
                Exit For
            End If
        Next
    Next
    RBCompat.Picture = LoadResPicture("RBY", vbResIcon)
    GSCompat.Picture = LoadResPicture("GSC", vbResIcon)
    RuSaCompat.Picture = LoadResPicture("ADV", vbResIcon)
    Call RefreshDisplay
End Sub

Private Sub RefreshDisplay()
    With Moves(MoveSort(CurrentMove))
        MoveName.Caption = .Name
        TypeImg.Picture = MainContainer.Types.ListImages(.Type).Picture
        TypeImg.ToolTipText = Element(.Type)
        AttackBar.Value = .Power
        AccBar.Value = .Accuracy
        PPBar.Value = .PP
        AttackLbl.Caption = "Power: " & .Power
        If .Accuracy > 0 Then
            AccLbl.Caption = "Accuracy: " & .Accuracy & "%"
            AccBar.Value = .Accuracy
        Else
            AccLbl.Caption = "Accuracy: 100%"
            AccBar.Value = 100
        End If
        PPLbl.Caption = "PP: " & .PP
        RBYTM.Caption = .OldTM
        GSCTM.Caption = .NewTM
        ADVTM.Caption = .ADVTM
        If Not .RBYMove Then RBCompat.Visible = False Else RBCompat.Visible = True
        If Not .GSCMove Then GSCompat.Visible = False Else GSCompat.Visible = True
        If Not .ADVMove Then RuSaCompat.Visible = False Else RuSaCompat.Visible = True
        Desc.Caption = .Text
        If .BrightPowder Then BPStat.Picture = Ball.Picture Else BPStat.Picture = GBall.Picture
        If .KingsRock Then KRStat.Picture = Ball.Picture Else KRStat.Picture = GBall.Picture
        If .SelfMove Then SMStat.Picture = Ball.Picture Else SMStat.Picture = GBall.Picture
        If .SubstituteBlocks Then SBStat.Picture = Ball.Picture Else SBStat.Picture = GBall.Picture
        If .SoundMove Then SndStat.Picture = Ball.Picture Else SndStat.Picture = GBall.Picture
        If .PhysMove Then CMStat.Picture = Ball.Picture Else CMStat.Picture = GBall.Picture
        If .HitsAll Then HBStat.Picture = Ball.Picture Else HBStat.Picture = GBall.Picture
        If .HitsTeam Then HAStat.Picture = Ball.Picture Else HAStat.Picture = GBall.Picture
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload MoveSearch
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            CurrentMove = CurrentMove - 1
            If CurrentMove = 0 Then CurrentMove = UBound(Moves)
            Call RefreshDisplay
        Case 2
            MoveSearch.Show
        Case 3
            CurrentMove = CurrentMove + 1
            If CurrentMove > UBound(Moves) Then CurrentMove = 1
            Call RefreshDisplay
    End Select
End Sub

Public Sub ChangeMe(ToMove As String)
    Dim X As Integer
    
    For X = 1 To UBound(Moves)
        If Moves(MoveSort(X)).Name = ToMove Then
            CurrentMove = X
            Exit For
        End If
    Next
    Call RefreshDisplay
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

