VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Playback 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playback"
   ClientHeight    =   7440
   ClientLeft      =   3810
   ClientTop       =   1860
   ClientWidth     =   7830
   Icon            =   "Playback.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleMode       =   0  'User
   ScaleWidth      =   7830
   Begin VB.TextBox txtFocusHolder 
      Height          =   375
      Left            =   8640
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "&Switch Team"
      Height          =   375
      Left            =   1320
      TabIndex        =   36
      Top             =   4680
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   4095
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7223
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Playback.frx":1272
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
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton ExpandWindow 
      Caption         =   "&More >>"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame PlayerFrame 
      Caption         =   "'s Team"
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   7575
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   0
         Left            =   240
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   20
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   24
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   -15
               TabIndex        =   25
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   1240
            Width           =   1095
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   0
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   0
            Left            =   120
            Top             =   400
            Width           =   840
         End
      End
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   1
         Left            =   1440
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   18
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   26
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   -15
               TabIndex        =   27
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   1
            Left            =   120
            Top             =   400
            Width           =   840
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   1
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   1240
            Width           =   1095
         End
      End
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   2
         Left            =   2640
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   16
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   28
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   -15
               TabIndex        =   29
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   2
            Left            =   120
            Top             =   400
            Width           =   840
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   2
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   17
            Top             =   1240
            Width           =   1095
         End
      End
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   3
         Left            =   3840
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   14
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   30
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   -15
               TabIndex        =   31
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   3
            Left            =   120
            Top             =   400
            Width           =   840
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   3
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   15
            Top             =   1240
            Width           =   1095
         End
      End
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   4
         Left            =   5040
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   12
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   32
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   -15
               TabIndex        =   33
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   4
            Left            =   120
            Top             =   400
            Width           =   840
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   4
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   13
            Top             =   1240
            Width           =   1095
         End
      End
      Begin VB.PictureBox AdvancedData 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Index           =   5
         Left            =   6240
         ScaleHeight     =   1515
         ScaleWidth      =   1035
         TabIndex        =   10
         Top             =   360
         Width           =   1095
         Begin VB.PictureBox FakeProgBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   39
            TabIndex        =   34
            Top             =   120
            Width           =   615
            Begin VB.Label lblProg 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   -15
               TabIndex        =   35
               Top             =   15
               Width           =   615
            End
         End
         Begin VB.Image AdvPoke 
            Height          =   840
            Index           =   5
            Left            =   120
            Top             =   400
            Width           =   840
         End
         Begin VB.Image CondIcon 
            Height          =   240
            Index           =   5
            Left            =   740
            Top             =   130
            Width           =   240
         End
         Begin VB.Label Pokename 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   11
            Top             =   1240
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   3585
      TabIndex        =   5
      Top             =   1680
      Width           =   3615
      Begin VB.Image ReplayPKMN 
         Height          =   960
         Index           =   2
         Left            =   2160
         Top             =   120
         Width           =   960
      End
      Begin VB.Image ReplayPKMN 
         Height          =   960
         Index           =   1
         Left            =   360
         Top             =   720
         Width           =   960
      End
      Begin VB.Image Shadow 
         Height          =   120
         Left            =   2400
         Picture         =   "Playback.frx":12F4
         Top             =   1020
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playback.frx":1341
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playback.frx":18DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playback.frx":1E75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playback.frx":240F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer HPAnimTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   720
   End
   Begin NetBattle.ColorProgress HPBar 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   4080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
   End
   Begin NetBattle.ColorProgress HPBar 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      ToolTipText     =   "Move Interval"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   4
      Min             =   1
      SelStart        =   5
      Value           =   5
   End
   Begin VB.Timer PBTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   720
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   741
      ButtonWidth     =   1482
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pause"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label WaitList 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time:  5 seconds"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   11
      Left            =   600
      Picture         =   "Playback.frx":29A9
      Top             =   4270
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   10
      Left            =   360
      Picture         =   "Playback.frx":2A32
      Top             =   4270
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   9
      Left            =   120
      Picture         =   "Playback.frx":2ABB
      Top             =   4270
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   8
      Left            =   600
      Picture         =   "Playback.frx":2B44
      Top             =   4030
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   7
      Left            =   360
      Picture         =   "Playback.frx":2BCD
      Top             =   4030
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   6
      Left            =   120
      Picture         =   "Playback.frx":2C56
      Top             =   4030
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   5
      Left            =   3480
      Picture         =   "Playback.frx":2CDF
      Top             =   1390
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   4
      Left            =   3240
      Picture         =   "Playback.frx":2D68
      Top             =   1390
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   3
      Left            =   3000
      Picture         =   "Playback.frx":2DF1
      Top             =   1390
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   2
      Left            =   3480
      Picture         =   "Playback.frx":2E7A
      Top             =   1150
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   1
      Left            =   3240
      Picture         =   "Playback.frx":2F03
      Top             =   1150
      Width           =   240
   End
   Begin VB.Image TIcon 
      Height          =   480
      Index           =   2
      Left            =   3120
      Top             =   600
      Width           =   480
   End
   Begin VB.Image TIcon 
      Height          =   480
      Index           =   1
      Left            =   240
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Active 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Active 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2775
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "Playback.frx":2F8C
      Top             =   1150
      Width           =   240
   End
   Begin VB.Image StatIcon 
      Height          =   240
      Index           =   0
      Left            =   9240
      Picture         =   "Playback.frx":3015
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StatIcon 
      Height          =   240
      Index           =   1
      Left            =   9480
      Picture         =   "Playback.frx":359F
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StatIcon 
      Height          =   240
      Index           =   2
      Left            =   9720
      Picture         =   "Playback.frx":3684
      Top             =   720
      Width           =   240
   End
End
Attribute VB_Name = "Playback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentTeam As Byte
Dim Switched As Boolean
Dim TimerLoops As Byte
Dim pos As Integer
Dim CachedCommand As String
Dim CachedData As String
Dim Started As Boolean
Dim RePlayer(1 To 2) As Trainer
Dim ReplayCommand() As String
Dim Data As String
Dim Loaded As Boolean
Dim PNum As Byte
Dim ONum As Byte
Dim EndCause As Byte
Public RTB As RTBClass
'The class module
Private ThisBattle As BattleData
'Blanks for swapping
Private BlankPKMN As Pokemon
Private BlankCondition As BattleStuff
Private BlankTC As TeamCond
Private TurnNumber As Integer
Private DrawCount As Integer
Private UnrateCount As Integer
'For display purposes, since the actual code is in the class module
Private BattleCurrent(2) As Pokemon
Private BattlePKMN(2, 6) As Pokemon
Private BattleCondition(2) As BattleStuff
Private BattleTC(2) As TeamCond
Private BattleWeather As Integer
Private NumPoke As Integer
Private StadiumMode As Boolean
Private FaintCount As Integer
Private BattleMode  As BattleModes
Private Const vbDarkGreen As Long = 38400
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const EM_LINESCROLL = &HB6
Private FileNum As Integer
Public CmdFile As String

Private Sub cmdNext_Click()
    TimerLoops = Slider1.Value
    Call PBTimer_Timer
End Sub

Private Sub Command1_Click()
    Close
    Unload Me
End Sub

'Private Sub AdvPoke_Click(Index As Integer)
'    Dim X As Byte
'
'    With BattlePKMN(CurrentTeam, Index + 1)
'        SelData.Text = .Nickname & ":"
'        Call AddSelLine("Max HP: " & .HP)
'        Call AddSelLine("Attack: " & .Attack)
'        Call AddSelLine("Defense: " & .Defense)
'        Call AddSelLine("Speed: " & .Speed)
'        Call AddSelLine("Sp.Attack: " & .SpecialAttack)
'        Call AddSelLine("Sp.Defense: " & .SpecialDefense)
'        For X = 1 To 4
'            Call AddSelLine("Move #" & X & ": " & Moves(.Move(X)).Name)
'        Next
'    End With
'End Sub

Private Sub ExpandWindow_Click()
    If Playback.Height = 7780 Then
        Playback.Height = 5550
        ExpandWindow.Caption = "&More >>"
    Else
        Playback.Height = 7780
        ExpandWindow.Caption = "&Less <<"
    End If
    Me.Refresh
    Call RefreshScreen
End Sub

Private Sub Form_Load()
    Set RTB = New RTBClass
    RTB.SetRTBHook Messages, txtFocusHolder
    Playback.Height = 5550
    CenterWindow Me
    ExpandWindow.Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Set ThisBattle = New BattleData
    ThisBattle.DoReplay = True
    If UseBG Then
        Call MainContainer.DoPicture("bg0.gif")
        Picture3.Picture = MainContainer.SwapSpace.Picture
    End If
    Loaded = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RTB.UnsetRTBHook
    Loader.Show
End Sub

Private Sub PBTimer_Timer()
    Dim Temp As String
    Dim E As Boolean
    TimerLoops = TimerLoops + 1
    If TimerLoops >= Slider1.Value Then
        E = PBTimer.Enabled
        PBTimer.Enabled = False
        If pos > UBound(ReplayCommand) Then Exit Sub
        TimerLoops = 0
        Do
            Call ProcessCommand
            pos = pos + 1
            If pos > UBound(ReplayCommand) Then Exit Do
        Loop Until Left(ReplayCommand(pos), 5) = "BCMD:"
        If pos = UBound(ReplayCommand) + 1 Then
            Toolbar1.Buttons(3).Value = tbrUnpressed
            Toolbar1.Buttons(4).Value = tbrUnpressed
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(4).Enabled = False
            Toolbar1.Refresh
            Started = False
            cmdNext.Visible = True
            cmdNext.Enabled = False
            Slider1.Visible = False
            WaitList.Visible = False
            Call AddMessage(vbNewLine & "End of replay.", , , , True)
            If ThisBattle.BattleOver Then
                Select Case ThisBattle.Winner
                    Case 3
                        If EndCause = 0 Then
                            Call AddMessage("Ended in a tie.", , , , True)
                        Else
                            Call AddMessage("The players agreed to tie.", , , , True)
                        End If
                    Case Else
                        If EndCause = 0 Then
                            Call AddMessage(RePlayer(ThisBattle.Winner).Name & " won.", , , , True)
                        ElseIf EndCause = 1 Then
                            Call AddMessage("The battle was hacked; " & RePlayer(ThisBattle.Winner).Name & " won by default.", , , , True)
                        Else
                            Call AddMessage("The battle timed out; " & RePlayer(ThisBattle.Winner).Name & " won by default.", , , , True)
                        End If
                End Select
                EndCause = 0
            End If
            Exit Sub
        End If
        PBTimer.Enabled = E
    End If
End Sub

Private Sub Slider1_Change()
    Dim Temp As String
    Temp = CStr(Slider1.Value) & " second" & IIf(Slider1.Value > 1, "s", "")
    WaitList.Caption = "Wait Time:" & vbNewLine & Temp
End Sub

Private Sub Slider1_Scroll()
    Call Slider1_Change
End Sub

Private Sub cmdSwitch_Click()
    Dim A As Integer
    Dim B As Integer
    Dim X As Byte
    Dim Y As Byte
    Dim V1 As Boolean
    Dim V2 As Boolean

    Switched = Not Switched
    If Switched Then X = 1 Else X = 2
    Y = OtherTeam(X)
    
    A = HPBar(X).Left
    B = HPBar(X).Top
    HPBar(X).Move HPBar(Y).Left, HPBar(Y).Top
    HPBar(Y).Move A, B
    
    A = Active(X).Left
    B = Active(X).Top
    Active(X).Move Active(Y).Left, Active(Y).Top
    Active(Y).Move A, B
    
    A = TIcon(X).Left
    B = TIcon(X).Top
    TIcon(X).Move TIcon(Y).Left, TIcon(Y).Top
    TIcon(Y).Move A, B
    
    ReplayPKMN(1).Visible = False
    ReplayPKMN(2).Visible = False
    Call RefreshScreen
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim X As Byte
    Dim Temp As String
    If Button.Value = tbrPressed Then Exit Sub
    Select Case Button.Index
        'Open
        Case 1
            If Started Then Call Toolbar1_ButtonClick(Toolbar1.Buttons(4))
            Toolbar1.Buttons(4).Value = tbrUnpressed
            If CmdFile = "" Then
                With MainContainer.FileBox
                    .DialogTitle = "Open Replay"
                    .CancelError = True
                    .FileName = ""
                    .DefaultExt = ".btl"
                    Temp = GetSetting("NetBattle", "Options", "InitDir", "")
                    If Temp <> "" Then .InitDir = Temp
                    .Filter = "Replay Files (*.btl)|*.btl"
                    On Error GoTo Cancelled
                    .ShowOpen
                    CmdFile = .FileName
                    SaveSetting "NetBattle", "Options", "InitDir", Left$(CmdFile, InStrRev(CmdFile, "\"))
                End With
            End If
            If CmdFile <> "" Then
                Messages.Text = ""
                Loaded = False
                ThisBattle.ResetBattle
                ThisBattle.DoReplay = True
                Call ReadReplay(CmdFile)
                CmdFile = ""
                If Loaded Then
                    TimerLoops = Slider1.Value
                    Call PBTimer_Timer
                    If Switched Then Call cmdSwitch_Click
                    cmdNext.Enabled = True
                    ExpandWindow.Enabled = True
                    Toolbar1.Buttons(3).Enabled = True
                    Toolbar1.Buttons(4).Enabled = True
                    Toolbar1.Buttons(5).Enabled = True
                    Toolbar1.Buttons(3).Value = tbrUnpressed
                    Toolbar1.Buttons(4).Value = tbrUnpressed
                Else
                    cmdNext.Enabled = False
                    Toolbar1.Buttons(3).Enabled = False
                    Toolbar1.Buttons(4).Enabled = False
                    Toolbar1.Buttons(5).Enabled = False
                    If Me.Height <> 5550 Then Call ExpandWindow_Click
                    ExpandWindow.Enabled = False
                    ReplayPKMN(1).Picture = Nothing
                    ReplayPKMN(2).Picture = Nothing
                    TIcon(1).Picture = Nothing
                    TIcon(2).Picture = Nothing
                    HPBar(1).Value = 0
                    HPBar(1).Max = 0
                    HPBar(1).RefreshBar
                    HPBar(2).Value = 0
                    HPBar(2).Max = 0
                    HPBar(2).RefreshBar
                    Active(1).Caption = ""
                    Active(2).Caption = ""
                    For X = 0 To 11
                        OpponentStat(X).Picture = StatIcon(0)
                    Next X
                End If
            End If
Cancelled:
        'Play
        Case 3
            If pos > UBound(ReplayCommand) Then Exit Sub
            Started = True
            PBTimer.Enabled = True
            Toolbar1.Buttons(3).Value = tbrPressed
            Toolbar1.Buttons(4).Value = tbrUnpressed
            cmdNext.Visible = False
            Slider1.Visible = True
            WaitList.Visible = True
        'Pause
        Case 4
            If Not Started Then Exit Sub
            Started = False
            PBTimer.Enabled = False
            Toolbar1.Buttons(3).Value = tbrUnpressed
            Toolbar1.Buttons(4).Value = tbrPressed
            cmdNext.Visible = True
            Slider1.Visible = False
            WaitList.Visible = False
        'Stop
        Case 5
            PBTimer.Enabled = False
            Started = False
            Messages.Text = ""
            ThisBattle.ResetBattle
            ThisBattle.DoReplay = True
            pos = 1
            TimerLoops = Slider1.Value
            Call PBTimer_Timer
            cmdNext.Visible = True
            cmdNext.Enabled = True
            Slider1.Visible = False
            WaitList.Visible = False
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(4).Enabled = True
            Toolbar1.Buttons(3).Value = tbrUnpressed
            Toolbar1.Buttons(4).Value = tbrUnpressed
    End Select
    Toolbar1.Refresh
End Sub

Private Sub ProcessCommand()
    Dim X As Byte
    Dim Worked As Boolean
    Dim Command As String
    Dim Temp As String
    Data = ReplayCommand(pos)
    Command = ChopString(Data, 5)
    With ThisBattle
        Select Case Command
            Case "INFO:"
                .Rules = ChopString(Data, 8)
                PNum = Val(ChopString(Data, 1))
                ONum = OtherTeam(PNum)
                RePlayer(ONum).Picture = Dec(ChopString(Data, 2))
                TIcon(ONum).Picture = MainContainer.Trainers.ListImages(RePlayer(ONum).Picture).Picture
                RePlayer(ONum).Version = Dec(ChopString(Data, 1))
                RePlayer(ONum).Name = Trim(ChopString(Data, 20))
                Worked = .SetTeam(ONum, Data)
                Worked = .SetVer(ONum, RePlayer(ONum).Version)
                Worked = .SetName(ONum, RePlayer(ONum).Name)
            Case "MYTM:"
                RePlayer(PNum).Name = Trim(ChopString(Data, 20))
                RePlayer(PNum).Picture = Dec(ChopString(Data, 2))
                TIcon(PNum).Picture = MainContainer.Trainers.ListImages(RePlayer(PNum).Picture).Picture
                RePlayer(PNum).Version = Dec(ChopString(Data, 1))
                Worked = .SetTeam(PNum, Data)
                Worked = .SetVer(PNum, RePlayer(PNum).Version)
                Worked = .SetName(PNum, RePlayer(PNum).Name)
                If Not .StadiumMode Then
                    .StartBattle
                    RefreshScreen
                End If
            Case "SPKM:"
                X = Val(ChopString(Data, 1))
                Worked = .SetSPoke(X, Data)
                If Not .NeedsStadiumSelect(OtherTeam(X)) Then
                    Worked = .DoThreePKMN
                    .StartBattle
                    RefreshScreen
                End If
            Case "CHAT:"
                Temp = Left(Data, InStr(1, Data, ":") - 1)
                If Temp = RePlayer(BottomNum).Name Then
                    Call AddMessage(Data, , ":", vbRed, True)
                ElseIf Temp = RePlayer(TopNum).Name Then
                    Call AddMessage(Data, , ":", vbBlue, True)
                Else
                    Call AddMessage(Data, , ":", vbDarkGreen, True)
                End If
            Case "BCMD:"
                Worked = .ParseBattle(Data, Me)
                Call RefreshScreen
            Case "HACK:"
                Call ThisBattle.ForceLoss(Val(Data))
                BattleCurrent(Val(Data)).HP = 0
                BattleCurrent(Val(Data)).Condition = 8
                For X = 1 To 6
                    BattlePKMN(Val(Data), X).HP = 0
                    BattlePKMN(Val(Data), X).Condition = 8
                Next X
                Call RefreshScreen
                EndCause = 1
            Case "TIME:"
                Call ThisBattle.ForceLoss(Val(Data))
                BattleCurrent(Val(Data)).HP = 0
                BattleCurrent(Val(Data)).Condition = 8
                For X = 1 To 6
                    BattlePKMN(Val(Data), X).HP = 0
                    BattlePKMN(Val(Data), X).Condition = 8
                Next X
                EndCause = 2
            Case "TIACC"
                Call ThisBattle.ForceLoss(1)
                Call ThisBattle.ForceLoss(2)
                EndCause = 3
        End Select
    End With
End Sub

Private Sub RefreshBattle()
    Dim X As Byte
    Dim Y As Byte
    
    For X = 1 To 2
        BattleCurrent(X) = GetClassPKMN(ThisBattle, X, 0)
        BattleTC(X) = GetClassTC(ThisBattle, X)
        For Y = 1 To 6
            BattlePKMN(X, Y) = GetClassPKMN(ThisBattle, X, Y)
        Next
    Next X
    For X = 1 To ThisBattle.ActNum
        BattleCondition(X) = GetClassBC(ThisBattle, X)
    Next X
End Sub

Private Sub RefreshScreen()
    Dim W As Byte
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Byte
    Dim TempStat As Integer
    Dim Percent As Single
    Dim TempVar As String
    Dim TTString As String
    Dim BImage(1 To 2) As String
    Dim LBPika As Boolean
    Dim TCWak As Boolean
    Dim MPDitto As Boolean
    
    Call RefreshBattle
    
    If Not Loaded Then Exit Sub
    For X = 1 To 2
        With BattleCurrent(X)
'            If UseBG Then
'                BImage(X) = ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, 3, True)
'            Else
'                BImage(X) = ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, RePlayer(X).Version, True)
'            End If
        End With
    Next X
    
    If BattleCondition(ONum).Substitute <> 0 Then
        If Switched Then
            TempVar = "substb.gif"
        Else
            TempVar = "subst.gif"
        End If
    ElseIf Switched Then
        If UseBG Then
            With BattleCurrent(ONum)
'                TempVar = ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, 3, True)
            End With
        Else
            TempVar = BImage(ONum)
        End If
    Else
        If UseBG Then
            With BattleCurrent(ONum)
'                TempVar = ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, 3)
            End With
        Else
            TempVar = BattleCurrent(ONum).Image
        End If
    End If
    Call MainContainer.DoPicture(TempVar)
    ReplayPKMN(2).Picture = MainContainer.SwapSpace.Picture
    
    If BattleCondition(PNum).Substitute <> 0 Then
        If Switched Then
            TempVar = "subst.gif"
        Else
            TempVar = "substb.gif"
        End If
    ElseIf Switched Then
        TempVar = BattleCurrent(PNum).Image
    Else
        TempVar = BImage(PNum)
    End If
    Call MainContainer.DoPicture(TempVar)
    ReplayPKMN(1).Picture = MainContainer.SwapSpace.Picture
    
    If Switched Then X = 2 Else X = 1
    Y = OtherTeam(X)
    ReplayPKMN(X).Move 360 + ((960 - ReplayPKMN(X).Width) / 2), 720 + (960 - ReplayPKMN(X).Height)
    ReplayPKMN(Y).Move 2160 + ((960 - ReplayPKMN(Y).Width) / 2), 120 + (960 - ReplayPKMN(Y).Height)
    If BattleCurrent(ONum).HP > 0 And BattleCondition(ONum).SemiInvul = 0 Then ReplayPKMN(2).Visible = True Else ReplayPKMN(2).Visible = False
    If BattleCurrent(PNum).HP > 0 And BattleCondition(PNum).SemiInvul = 0 Then ReplayPKMN(1).Visible = True Else ReplayPKMN(1).Visible = False
    If InStr(1, BattleCurrent(Y).Image, "rs") >= 1 Then
        ReplayPKMN(Y).Top = ReplayPKMN(Y).Top - (Screen.TwipsPerPixelY * BasePKMN(BattleCurrent(Y).No).Offset)
        If ReplayPKMN(Y).Visible = True And BasePKMN(BattleCurrent(Y).No).Offset > 0 Then Shadow.Visible = True Else Shadow.Visible = False
    End If

    For X = 1 To 2
        If X = 1 Then W = TopNum Else W = BottomNum
        For Y = 1 To ThisBattle.NumPoke
            If X = 1 Then Z = Y - 1 Else Z = Y + 5
            Select Case BattlePKMN(W, Y).Condition
                Case 1
                    OpponentStat(Z).Picture = StatIcon(0).Picture
                Case 8
                    OpponentStat(Z).Picture = StatIcon(2).Picture
                Case Else
                    OpponentStat(Z).Picture = StatIcon(1).Picture
            End Select
        Next Y
    Next X
    For X = ThisBattle.NumPoke + 1 To 6
        OpponentStat(X - 1).Visible = False
        OpponentStat(X + 5).Visible = False
    Next X

    Active(1).Caption = RePlayer(PNum).Name & "'s " & BattleCurrent(PNum).Nickname & " (Lv." & BattleCurrent(PNum).Level & " " & BattleCurrent(PNum).Name & ") - " & Condition(BattleCurrent(PNum).Condition)
    If HPBar(1).Value > BattleCurrent(PNum).MaxHP Then HPBar(1).Value = 0
    HPBar(1).Max = BattleCurrent(PNum).MaxHP
    HPBar(1).Value = BattleCurrent(PNum).HP
    Active(2).Caption = RePlayer(ONum).Name & "'s " & BattleCurrent(ONum).Nickname & " (Lv." & BattleCurrent(ONum).Level & " " & BattleCurrent(ONum).Name & ") - " & Condition(BattleCurrent(ONum).Condition)
    If HPBar(2).Value > BattleCurrent(ONum).MaxHP Then HPBar(2).Value = 0
    HPBar(2).Max = BattleCurrent(ONum).MaxHP
    HPBar(2).Value = BattleCurrent(ONum).HP


    For X = 1 To 2
        TTString = ""
        TCWak = ((BattleCurrent(X).No = 104 Or BattleCurrent(X).No = 105) And BattleCurrent(X).Item = 40)
        LBPika = (BattleCurrent(X).No = 25 And BattleCurrent(X).Item = 22)
        MPDitto = (BattlePKMN(X, BattleCurrent(X).TeamNumber).No = 132 And BattleCurrent(X).Item = 26)
        
        TempStat = BattleCurrent(X).Attack * StatChange(BattleCondition(X).AttackChange)
        If TCWak Then TempStat = ThisBattle.Rollover(TempStat * 2)
        TTString = TTString & "ATK:" & Cap(TempStat)
        If BattleCondition(X).AttackChange <> 0 Then TTString = TTString & "*"
        TTString = TTString & " • "
        
        TempStat = BattleCurrent(X).Defense * StatChange(BattleCondition(X).DefenseChange)
        If MPDitto Then TempStat = ThisBattle.Rollover(TempStat * 1.5)
        TTString = TTString & "DEF:" & Cap(TempStat)
        If BattleCondition(X).DefenseChange <> 0 Then TTString = TTString & "*"
        TTString = TTString & " • "
        
        TempStat = BattleCurrent(X).Speed * StatChange(BattleCondition(X).SpeedChange)
        TTString = TTString & "SPD:" & Cap(TempStat)
        If BattleCondition(X).SpeedChange <> 0 Then TTString = TTString & "*"
        TTString = TTString & " • "
        
        TempStat = BattleCurrent(X).SpecialAttack * StatChange(BattleCondition(X).SAttackChange)
        If LBPika Then TempStat = ThisBattle.Rollover(TempStat * 2)
        TTString = TTString & "SATK:" & Cap(TempStat)
        If BattleCondition(X).SAttackChange <> 0 Then TTString = TTString & "*"
        TTString = TTString & " • "
        
        TempStat = BattleCurrent(X).SpecialDefense * StatChange(BattleCondition(X).SDefenseChange)
        If MPDitto Then TempStat = ThisBattle.Rollover(TempStat * 1.5)
        TTString = TTString & "SDEF:" & Cap(TempStat)
        If BattleCondition(X).SDefenseChange <> 0 Then TTString = TTString & "*"
        TTString = TTString & " • "
        
        TempStat = BattleCondition(X).EvadeChange
        TTString = TTString & "EVA:"
        If TempStat >= 0 Then TTString = TTString & "+"
        TTString = TTString & TempStat & " • "
        
        TempStat = BattleCondition(X).AccuracyChange
        TTString = TTString & "ACC:"
        If TempStat >= 0 Then TTString = TTString & "+"
        TTString = TTString & TempStat

        If X = PNum Then
            ReplayPKMN(1).ToolTipText = TTString
        Else
            ReplayPKMN(2).ToolTipText = TTString
        End If
    Next X

    PlayerFrame.Caption = RePlayer(BottomNum).Name & "'s Team"
    For X = 0 To 5
        With BattlePKMN(BottomNum, X + 1)
'            Call MainContainer.DoPicture(ChooseImage(.No, .DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk, 2))
            AdvPoke(X).Picture = MainContainer.SwapSpace.Picture
            CondIcon(X).Picture = MainContainer.Conditions.ListImages(.Condition).Picture
            
            'In case you're wondering, I used picture boxes instead of actual
            'ProgressBars because the caption text changed to bright green
            'when the fill went past it.
            lblProg(X).Caption = .HP
            If .HP = 0 Then
                FakeProgBar(X).Picture = Nothing
            Else
                Percent = .HP / .MaxHP
                Y = FakeProgBar(X).ScaleHeight
                Z = FakeProgBar(X).ScaleWidth
                W = Int(Percent * Z)
                If W <> Z Then
                    FakeProgBar(X).Line (W + 1, 0)-(Z, Y), vbWhite, BF
                End If
                If Percent > 0.5 Then
                    FakeProgBar(X).Line (0, 0)-(W, Y), vbGreen, BF
                ElseIf Percent > 0.25 Then
                    FakeProgBar(X).Line (0, 0)-(W, Y), RGB(245, 245, 0), BF 'Yellow, just not as bright
                Else
                    FakeProgBar(X).Line (0, 0)-(W, Y), vbRed, BF
                End If
            End If
            Pokename(X).Caption = .Nickname
            AdvPoke(X).ToolTipText = .Name
        End With
    Next
End Sub

Public Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    If DebugMessage And Not DebugMode Then Exit Sub
    Call RTB.AddMessage(Message, BreakChar, Color, Bold, Italic)
End Sub

Private Sub ReadReplay(FileName As String)
    Dim UCSize As Long
    Dim HSize As Long
    Dim ByteArray() As Byte
    Dim Final As String
    Dim X As Long
    Dim Y As Integer
    Dim C1 As Long
    Dim C2 As Long
    Dim C3 As Long
    Dim C4 As Long
    Dim B As Boolean
    Dim Worked As Boolean
    On Error GoTo LoadFailed
    ReDim ReplayCommand(0)
    FileNum = FreeFile
    If Not FileExists(FileName) Then Exit Sub
    Open FileName For Input As #FileNum
    Input #FileNum, UCSize
    HSize = Len(CStr(UCSize)) + Len(vbCrLf) + 1
    ReDim ByteArray(LOF(FileNum) - HSize) As Byte
    Close #FileNum
    Open FileName For Binary Access Read As #FileNum
    Get #FileNum, HSize, ByteArray()
    Close #FileNum
    Worked = MainContainer.Compressor.DecompressData(ByteArray(), UCSize)
    'Check the checksums
    For X = 0 To UBound(ByteArray) - 4
        C1 = C1 + ByteArray(X)
        If C1 > 10000 Then C1 = C1 Mod 89
        C2 = C2 + ByteArray(X) - 55
        If Abs(C2) > 10000 Then C2 = C2 Mod 75
        If X <> 0 Then C3 = C3 + (ByteArray(X - 1) Xor ByteArray(X)) - 10
        If Abs(C3) > 100000 Then C3 = C3 Mod 101
        C4 = C4 + ByteArray(X) - ByteArray(UBound(ByteArray) - X - 4) + 10
        If Abs(C4) > 1000 Then C3 = C3 Mod 39
    Next X
    B = True
    If ByteArray(X) <> Abs(C1) Mod 256 Then B = False
    If ByteArray(X + 1) <> Abs(C2) Mod 256 Then B = False
    If ByteArray(X + 2) <> Abs(C3) Mod 256 Then B = False
    If ByteArray(X + 3) <> Abs(C4) Mod 256 Then B = False
    If Not B Then
        MsgBox "This replay file is invalid.", vbCritical, "Invalid Replay"
        Exit Sub
    End If
    ReDim Preserve ByteArray(X - 1)
    Final = String$(UBound(ByteArray) + 1, " ")
    For X = 0 To UBound(ByteArray)
        Mid(Final, X + 1) = Chr$(ByteArray(X))
    Next X
    ReDim ReplayCommand(0)
    Y = 1
    Do Until Len(Final) = 0
        ReDim Preserve ReplayCommand(Y)
        X = Asc(ChopString(Final, 1)) + 1
        ReplayCommand(Y) = XORDecrypt(ChopString(Final, X))
        Select Case Left(ReplayCommand(Y), 5)
        Case "CHAT:", "BCMD:", "INFO:", "MYTM:", "SPKM:", "HACK:", "TIME:", "TIACC:"
        Case Else
            MsgBox "This replay file is invalid.", vbCritical, "Invalid Replay"
            Exit Sub
        End Select
        Y = Y + 1
    Loop
    pos = 1
    Loaded = True
    Exit Sub
LoadFailed:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error Loading Replay."
    If InVBMode Then
        Stop
        Resume
    End If
    On Error Resume Next
    Close #FileNum
End Sub

Public Function BottomNum() As Byte
    If Switched Then
        BottomNum = OtherTeam(PNum)
    Else
        BottomNum = PNum
    End If
End Function
Public Function TopNum() As Byte
    If Switched Then
        TopNum = OtherTeam(ONum)
    Else
        TopNum = ONum
    End If
End Function

