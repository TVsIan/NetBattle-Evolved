VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Battle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battle"
   ClientHeight    =   6690
   ClientLeft      =   3975
   ClientTop       =   1845
   ClientWidth     =   8130
   ClipControls    =   0   'False
   Icon            =   "Battle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   8130
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo Move"
      Height          =   375
      Left            =   360
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2250
   End
   Begin VB.Timer tmrDX 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   3960
   End
   Begin VB.Timer tmrDelay 
      Left            =   9360
      Top             =   3240
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "&1: Ally Deoxys"
      Height          =   325
      Index           =   0
      Left            =   5400
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "&1: Ally Deoxys"
      Height          =   325
      Index           =   1
      Left            =   5400
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   5055
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "&1: Ally Deoxys"
      Height          =   325
      Index           =   2
      Left            =   5400
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   5430
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtWatchIgnore 
      Enabled         =   0   'False
      Height          =   555
      Left            =   3760
      TabIndex        =   99
      Text            =   "Both players are ignoring spectator chat."
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Timer tmrBattleQueue 
      Interval        =   1
      Left            =   8880
      Top             =   3240
   End
   Begin VB.Timer PBTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   3240
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "&Leave"
      Height          =   375
      Left            =   3780
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1255
   End
   Begin VB.CommandButton SendMsg 
      Caption         =   "Sen&d"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1255
   End
   Begin VB.PictureBox picBuild 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   41
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPKMNMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picPKMNImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8760
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picShadowMask 
      Height          =   180
      Left            =   9840
      ScaleHeight     =   120
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   2520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picShadow 
      AutoRedraw      =   -1  'True
      Height          =   180
      Left            =   9840
      Picture         =   "Battle.frx":1272
      ScaleHeight     =   120
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTerrain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   3375
      Left            =   3765
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5953
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Battle.frx":12BF
   End
   Begin VB.Timer FlashTimer 
      Interval        =   125
      Left            =   8880
      Top             =   2760
   End
   Begin VB.PictureBox BattleArea 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1700
      Left            =   60
      ScaleHeight     =   1665
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   1260
      Width           =   3630
      Begin VB.Image PokeCenter 
         Height          =   1680
         Left            =   0
         Picture         =   "Battle.frx":1341
         Top             =   0
         Width           =   3600
      End
      Begin VB.Image Computer 
         Height          =   1680
         Left            =   0
         Picture         =   "Battle.frx":26B7
         Top             =   0
         Visible         =   0   'False
         Width           =   3600
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Weather"
            Object.ToolTipText     =   "Weather"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Terrain"
            Object.Tag             =   "Terrain"
            Object.ToolTipText     =   "Terrain"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9155
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Chatbox 
      Height          =   555
      Left            =   3760
      MaxLength       =   235
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Timer HPAnimTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8400
      Top             =   2760
   End
   Begin VB.Timer ReadyTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   9360
      Top             =   2760
   End
   Begin TabDlg.SSTab ControlTab 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Attack"
      TabPicture(0)   =   "Battle.frx":33CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MoveDesc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MoveSel(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MoveSel(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MoveSel(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MoveSel(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Attack"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Pokémon"
      TabPicture(1)   =   "Battle.frx":33E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PokeTiles"
      Tab(1).Control(1)=   "Switch"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox PokeTiles 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -74760
         ScaleHeight     =   1095
         ScaleWidth      =   7335
         TabIndex        =   68
         Top             =   480
         Width           =   7335
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   0
            Picture         =   "Battle.frx":3402
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   0
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   0
               Left            =   -240
               TabIndex        =   98
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   97
               Top             =   240
               Width           =   465
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   96
               Top             =   240
               Width           =   945
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   0
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   0
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   660
               TabIndex        =   95
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   0
               Left            =   600
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   1
            Left            =   2520
            Picture         =   "Battle.frx":3536
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   0
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   1
               Left            =   -120
               TabIndex        =   93
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   660
               TabIndex        =   92
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   1
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   1
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   91
               Top             =   240
               Width           =   945
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   90
               Top             =   240
               Width           =   465
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   1
               Left            =   600
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   2
            Left            =   5050
            Picture         =   "Battle.frx":36BF
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   0
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   2
               Left            =   -120
               TabIndex        =   88
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   87
               Top             =   240
               Width           =   465
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   840
               TabIndex        =   86
               Top             =   240
               Width           =   945
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   2
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   2
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   660
               TabIndex        =   85
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   2
               Left            =   600
               Top             =   240
               Width           =   240
            End
         End
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   3
            Left            =   0
            Picture         =   "Battle.frx":3848
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   600
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   3
               Left            =   -120
               TabIndex        =   83
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   660
               TabIndex        =   82
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   3
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   3
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   3
               Left            =   600
               Top             =   240
               Width           =   240
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   840
               TabIndex        =   81
               Top             =   240
               Width           =   945
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   1800
               TabIndex        =   80
               Top             =   240
               Width           =   465
            End
         End
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   4
            Left            =   2520
            Picture         =   "Battle.frx":39D1
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   600
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   4
               Left            =   -120
               TabIndex        =   78
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   660
               TabIndex        =   77
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   4
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   4
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   4
               Left            =   600
               Top             =   240
               Width           =   240
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   840
               TabIndex        =   76
               Top             =   240
               Width           =   945
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   1800
               TabIndex        =   75
               Top             =   240
               Width           =   465
            End
         End
         Begin VB.PictureBox Entry 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   5
            Left            =   5050
            Picture         =   "Battle.frx":3B5A
            ScaleHeight     =   495
            ScaleWidth      =   2250
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   600
            Width           =   2250
            Begin VB.Label SwitchTile 
               BackStyle       =   0  'Transparent
               Height          =   735
               Index           =   5
               Left            =   -120
               TabIndex        =   73
               Top             =   -120
               Width           =   2535
            End
            Begin VB.Label PKName 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   660
               TabIndex        =   72
               Top             =   30
               Width           =   1305
            End
            Begin VB.Image PokeIcon 
               Height          =   480
               Index           =   5
               Left            =   60
               Top             =   0
               Width           =   480
            End
            Begin VB.Image PKGender 
               Height          =   240
               Index           =   5
               Left            =   1920
               Top             =   0
               Width           =   240
            End
            Begin VB.Image StatIcon 
               Height          =   240
               Index           =   5
               Left            =   600
               Top             =   240
               Width           =   240
            End
            Begin VB.Label PKHP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   840
               TabIndex        =   71
               Top             =   240
               Width           =   945
            End
            Begin VB.Label PKLevel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   1800
               TabIndex        =   70
               Top             =   240
               Width           =   465
            End
         End
      End
      Begin VB.CommandButton Switch 
         Caption         =   "&Switch"
         Height          =   375
         Left            =   -69720
         TabIndex        =   19
         Top             =   1680
         Width           =   2250
      End
      Begin VB.CommandButton Attack 
         Caption         =   "&Attack"
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   1680
         Width           =   2250
      End
      Begin VB.PictureBox MoveSel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   240
         Picture         =   "Battle.frx":3CE3
         ScaleHeight     =   495
         ScaleWidth      =   2250
         TabIndex        =   15
         Top             =   480
         Width           =   2250
         Begin VB.Label AttackTile 
            BackStyle       =   0  'Transparent
            Height          =   735
            Index           =   0
            Left            =   -120
            TabIndex        =   22
            Top             =   -120
            Width           =   2535
         End
         Begin VB.Label MoveNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   30
            Width           =   2055
         End
         Begin VB.Image MoveType 
            Height          =   240
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   240
         End
         Begin VB.Label PPLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00/00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox MoveSel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   2760
         Picture         =   "Battle.frx":3E17
         ScaleHeight     =   495
         ScaleWidth      =   2250
         TabIndex        =   12
         Top             =   480
         Width           =   2250
         Begin VB.Label AttackTile 
            BackStyle       =   0  'Transparent
            Height          =   735
            Index           =   1
            Left            =   -120
            TabIndex        =   23
            Top             =   -120
            Width           =   2535
         End
         Begin VB.Label PPLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00/00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.Image MoveType 
            Height          =   240
            Index           =   1
            Left            =   120
            Top             =   240
            Width           =   240
         End
         Begin VB.Label MoveNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   30
            Width           =   2055
         End
      End
      Begin VB.PictureBox MoveSel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   240
         Picture         =   "Battle.frx":3FA0
         ScaleHeight     =   495
         ScaleWidth      =   2250
         TabIndex        =   9
         Top             =   1080
         Width           =   2250
         Begin VB.Label AttackTile 
            BackStyle       =   0  'Transparent
            Height          =   735
            Index           =   2
            Left            =   -120
            TabIndex        =   24
            Top             =   -120
            Width           =   2535
         End
         Begin VB.Label MoveNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   30
            Width           =   2055
         End
         Begin VB.Image MoveType 
            Height          =   240
            Index           =   2
            Left            =   120
            Top             =   240
            Width           =   240
         End
         Begin VB.Label PPLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00/00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox MoveSel 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   2760
         Picture         =   "Battle.frx":4129
         ScaleHeight     =   495
         ScaleWidth      =   2250
         TabIndex        =   6
         Top             =   1080
         Width           =   2250
         Begin VB.Label AttackTile 
            BackStyle       =   0  'Transparent
            Height          =   735
            Index           =   3
            Left            =   -120
            TabIndex        =   25
            Top             =   -120
            Width           =   2535
         End
         Begin VB.Label MoveNameLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   30
            Width           =   2055
         End
         Begin VB.Image MoveType 
            Height          =   240
            Index           =   3
            Left            =   120
            Top             =   240
            Width           =   240
         End
         Begin VB.Label PPLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "00/00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
      End
      Begin RichTextLib.RichTextBox MoveDesc 
         Height          =   1095
         Left            =   5280
         TabIndex        =   20
         Tag             =   "0"
         Top             =   480
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         DisableNoScroll =   -1  'True
         TextRTF         =   $"Battle.frx":42B2
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   2
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   2955
      TabIndex        =   26
      Top             =   3000
      Width           =   2960
      Begin NetBattle.ColorProgress HPBar 
         Height          =   270
         Index           =   2
         Left            =   195
         TabIndex        =   27
         Top             =   870
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
      End
      Begin NetBattle.ColorProgress HPBar 
         Height          =   270
         Index           =   4
         Left            =   75
         TabIndex        =   29
         Top             =   270
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
      End
      Begin VB.Image PokeCond 
         Height          =   240
         Index           =   2
         Left            =   2595
         Top             =   630
         Width           =   240
      End
      Begin VB.Image PokeCond 
         Height          =   240
         Index           =   4
         Left            =   2475
         Top             =   30
         Width           =   240
      End
      Begin VB.Shape SelPKMN 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         FillColor       =   &H00C0C0FF&
         Height          =   540
         Index           =   4
         Left            =   60
         Top             =   30
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Shape SelPKMN 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         FillColor       =   &H00C0C0FF&
         Height          =   540
         Index           =   2
         Left            =   180
         Top             =   630
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label PokeText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   30
         Top             =   60
         Width           =   2595
      End
      Begin VB.Label PokeText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   225
         TabIndex        =   28
         Top             =   660
         Width           =   2595
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   1
      Left            =   840
      ScaleHeight     =   1215
      ScaleWidth      =   2955
      TabIndex        =   31
      Top             =   30
      Width           =   2960
      Begin NetBattle.ColorProgress HPBar 
         Height          =   270
         Index           =   3
         Left            =   195
         TabIndex        =   32
         Top             =   870
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
      End
      Begin NetBattle.ColorProgress HPBar 
         Height          =   270
         Index           =   1
         Left            =   75
         TabIndex        =   33
         Top             =   270
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   476
      End
      Begin VB.Image PokeCond 
         Height          =   240
         Index           =   3
         Left            =   2595
         Top             =   630
         Width           =   240
      End
      Begin VB.Shape SelPKMN 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         FillColor       =   &H00C0C0FF&
         Height          =   540
         Index           =   3
         Left            =   180
         Top             =   630
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Image PokeCond 
         Height          =   240
         Index           =   1
         Left            =   2475
         Top             =   30
         Width           =   240
      End
      Begin VB.Shape SelPKMN 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         FillColor       =   &H00C0C0FF&
         Height          =   540
         Index           =   1
         Left            =   60
         Top             =   30
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label PokeText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   34
         Top             =   60
         Width           =   2595
      End
      Begin VB.Label PokeText 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   225
         TabIndex        =   35
         Top             =   660
         Width           =   2595
      End
   End
   Begin VB.PictureBox ReplayControls 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3780
      ScaleHeight     =   1095
      ScaleWidth      =   4215
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdStop 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         Picture         =   "Battle.frx":4334
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdPlay 
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         Picture         =   "Battle.frx":48BE
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdPause 
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         Picture         =   "Battle.frx":4E48
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   46
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSwitchTeam 
         Caption         =   "&Switch Team"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   375
         Left            =   2880
         TabIndex        =   44
         Top             =   600
         Width           =   1335
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   1440
         TabIndex        =   47
         ToolTipText     =   "Move Interval"
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   4
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label WaitList 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait Time:  5 seconds"
         Height          =   375
         Left            =   0
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame OldSwitchFrame 
      Caption         =   "Pokémon"
      Height          =   1695
      Left            =   3240
      TabIndex        =   52
      Top             =   4560
      Visible         =   0   'False
      Width           =   4695
      Begin MSComctlLib.ListView OldPokeList 
         Height          =   855
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1508
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "HP"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4455
         TabIndex        =   64
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton cmdOldCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   0
            TabIndex        =   67
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton OldSwitch 
            Caption         =   "&Switch"
            Height          =   375
            Left            =   1920
            TabIndex        =   65
            Top             =   960
            Visible         =   0   'False
            Width           =   2535
         End
      End
   End
   Begin VB.Frame OldAttackFrame 
      Caption         =   "Attacks"
      Height          =   1695
      Left            =   240
      TabIndex        =   53
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   2535
         TabIndex        =   54
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton OldAttack 
            Caption         =   "OldAttack"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton OldAttack 
            Caption         =   "OldAttack"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   58
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton OldAttack 
            Caption         =   "OldAttack"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   57
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton OldAttack 
            Caption         =   "OldAttack"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   56
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton OldAttackButton 
            Caption         =   "&Attack"
            Height          =   375
            Left            =   0
            TabIndex        =   55
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label OldPP 
            Alignment       =   2  'Center
            Caption         =   "00/00"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   63
            Top             =   0
            Width           =   615
         End
         Begin VB.Label OldPP 
            Alignment       =   2  'Center
            Caption         =   "00/00"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   62
            Top             =   240
            Width           =   615
         End
         Begin VB.Label OldPP 
            Alignment       =   2  'Center
            Caption         =   "00/00"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   61
            Top             =   480
            Width           =   615
         End
         Begin VB.Label OldPP 
            Alignment       =   2  'Center
            Caption         =   "00/00"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   60
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Image TIcon 
      Height          =   480
      Index           =   1
      Left            =   120
      Top             =   60
      Width           =   480
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   5
      Left            =   510
      Picture         =   "Battle.frx":53D2
      Top             =   900
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   4
      Left            =   270
      Picture         =   "Battle.frx":545B
      Top             =   900
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   3
      Left            =   30
      Picture         =   "Battle.frx":54E4
      Top             =   900
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   2
      Left            =   510
      Picture         =   "Battle.frx":556D
      Top             =   660
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   1
      Left            =   270
      Picture         =   "Battle.frx":55F6
      Top             =   660
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   0
      Left            =   30
      Picture         =   "Battle.frx":567F
      Top             =   660
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   6
      Left            =   3000
      Picture         =   "Battle.frx":5708
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   7
      Left            =   3240
      Picture         =   "Battle.frx":5791
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   8
      Left            =   3480
      Picture         =   "Battle.frx":581A
      Top             =   3060
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   9
      Left            =   3000
      Picture         =   "Battle.frx":58A3
      Top             =   3300
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   10
      Left            =   3240
      Picture         =   "Battle.frx":592C
      Top             =   3300
      Width           =   240
   End
   Begin VB.Image OpponentStat 
      Height          =   240
      Index           =   11
      Left            =   3480
      Picture         =   "Battle.frx":59B5
      Top             =   3300
      Width           =   240
   End
   Begin VB.Image TIcon 
      Height          =   480
      Index           =   0
      Left            =   3120
      Top             =   3660
      Width           =   480
   End
   Begin VB.Image GoodImage 
      Height          =   480
      Left            =   8400
      Picture         =   "Battle.frx":5A3E
      Top             =   1320
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image GenderImage 
      Height          =   240
      Index           =   2
      Left            =   9840
      Picture         =   "Battle.frx":5BC7
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image GenderImage 
      Height          =   240
      Index           =   1
      Left            =   9480
      Picture         =   "Battle.frx":5C28
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image SelImage 
      Height          =   480
      Left            =   8400
      Picture         =   "Battle.frx":5C8E
      Top             =   720
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image BadImage 
      Height          =   480
      Left            =   8400
      Picture         =   "Battle.frx":5DC2
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image PKBall 
      Height          =   240
      Index           =   2
      Left            =   9120
      Picture         =   "Battle.frx":5F4B
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PKBall 
      Height          =   240
      Index           =   1
      Left            =   8760
      Picture         =   "Battle.frx":604F
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PKBall 
      Height          =   240
      Index           =   0
      Left            =   8400
      Picture         =   "Battle.frx":6134
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save Log..."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &Dump..."
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&End Game"
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuReplayFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuReplayFileItem 
         Caption         =   "&Open..."
         Index           =   0
      End
      Begin VB.Menu mnuReplayFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuReplayFileItem 
         Caption         =   "&Close Replay Window"
         Index           =   2
      End
      Begin VB.Menu mnuReplayFileItem 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuBattle 
      Caption         =   "&Battle"
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Run Away (Forfeit)"
         Index           =   0
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Propose Draw"
         Index           =   1
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "Request &Unrated Battle"
         Index           =   2
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "&Save Replay"
         Index           =   4
      End
      Begin VB.Menu mnuBattleItem 
         Caption         =   "Save &Log"
         Index           =   5
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Drain PP"
         Index           =   0
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Copy PKMN"
         Index           =   1
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Refresh Listing"
         Index           =   2
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Add messages"
         Index           =   3
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Use AI Move"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "Do &Pokemon Test"
         Index           =   5
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Sync Test"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Transfer Test"
         Enabled         =   0   'False
         Index           =   7
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Sound"
         Index           =   0
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Music"
         Index           =   1
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Disable Delay"
         Index           =   3
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Ignore Watch Chat"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPokedex 
      Caption         =   "&DataDex"
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&PokéDex"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&MoveDex"
         Index           =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&BattleDex"
         Index           =   2
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPokedexItem 
         Caption         =   "&DamageCalc"
         Index           =   3
         Shortcut        =   ^D
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
      End
   End
End
Attribute VB_Name = "Battle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Battle.frm
'Used to be where all the good stuff happens.
'Now it's just a frontend for the server.
Option Explicit
Private Type ImageType
    OrigVis As Boolean
    Vis As Boolean
    X As Single
    Y As Single
End Type

'The class module
Private ThisBattle As BattleData
Public PNum As Long
Public ONum As Long
Public ImJustWatching As Boolean
Public WatchID As Integer
Public WatchP1 As Integer
Public WatchP2 As Integer
Public RTB As RTBClass
Public ActNum As Long
Private Opponent As Trainer
Private BattleID As Integer
'Blanks for swapping
Private BlankPKMN As Pokemon
Private BlankCondition As BattleStuff
Private BlankTC As TeamCond
Private TurnNumber As Integer
Private DrawCount As Integer
Private UnrateCount As Integer
'For display puropses, since the actual code is in the class module
Private BattleCurrent(4) As Pokemon
Private BattlePKMN(2, 6) As Pokemon
Private BattleCondition(4) As BattleStuff
Private BattleTC(4) As TeamCond
Private BattleWeather As Integer
Private SelectedMove(1 To 4) As Integer
Private SelectedPKMN As Integer
Private Const vbDarkGreen As Long = 38400
Private OriginalTop(1 To 4) As Integer
Private OriginalLeft(1 To 4) As Integer
'These must be used instead of You.Name and Opponent.Name due to Watch Battle.
Private YourName As String
Private OpponentName As String
'Rules
Public StadiumPKMN As Integer
Public OppStadPKMN As Integer
Private NumPoke As Integer
Private StadiumMode As Boolean
Private FaintCount As Integer
Public RBYMode As Boolean
'Network stuff, mostly.
Private FinalExit As Boolean
Private BufferedData As String
Private ReceivedBOVER As Boolean
Private StartedTransfer As Boolean
Private NetworkStuff(1024) As String
Private TFileName As String
Private TabSwitch As Boolean
Private StoredRandTeam As String
'For the ToolTips
Private StoredTT1 As String
Private StoredTT2 As String
Private Display2 As Boolean
Private Display1 As Boolean
Private TT As CTooltip
Private Display(5) As Boolean
Private StoredTT(5) As String
Private CurrentTT(1 To 4) As String
'Declarations required for the scrollable text window
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const EM_LINESCROLL = &HB6
Private ReplayNum As Integer
'These are for keyboard control
Private KeyMode As Byte
Private ValidTarget(1 To 4) As Boolean
Private SelectedTarget As Byte
'Extra variables for Double Battles
Private DoFlash(1 To 4) As Boolean
Private SelPoke As Long
Private FirstSwitch As Byte
Private SelMove(1 To 4) As Byte
Private SelTarg(1 To 4) As Byte
Private SelSwitch(1 To 4) As Byte
Private Shadow(1 To 4) As ImageType
Private PKMNImage(1 To 4) As ImageType
Public SkipDelay As Boolean
Public DelayLoaded As Boolean
Public Delay As Long
Private SyncString As String
Private BattleQueue() As String
'For replays
Private IsReplayWindow As Boolean
Private CmdFile As String
Private CurrentTeam As Byte
Private Switched As Boolean
Private TimerLoops As Byte
Private Pos As Integer
Private CachedCommand As String
Private CachedData As String
Private Started As Boolean
Private RePlayer(1 To 2) As Trainer
Private ReplayCommand() As String
Private Loaded As Boolean
Private EndCause As Byte
Private ReplayVersion As String
Private FileNum As Byte
Private RePlayerTemp(1 To 2) As Trainer
Private Cancellable As Boolean
Public Resuming As Boolean

'DirectX =D
'Public WithEvents DX As clsDX


Sub DoTie()
    Dim P1 As Byte
    Dim P2 As Byte
    Dim X As Integer
    If ThisBattle.BattleOver Then Exit Sub
    P1 = 0
    P2 = 0
    For X = 1 To NumPoke
        If BattlePKMN(1, X).HP <> 0 Then P1 = P1 + 1
        If BattlePKMN(2, X).HP <> 0 Then P2 = P2 + 1
    Next X
    Call AddMessage("")
    ThisBattle.BattleOver = True
    Call ThisBattle.AddMessage(31)
    Call ThisBattle.AddMessage(36, nbNumber, P1, nbNumber, P2)
    If Not ImJustWatching Then
        Call AddMessage(YourName & ": " & You.LoseMess, False, ":", vbRed, True, False)
        Call SendData("CMSG:" & You.LoseMess)
    End If
    Call DoEndBattle
End Sub

Private Sub Command2_Click()
    While True
        'DX.Blt
        DoEvents
    Wend
End Sub

Private Function CheckArea(X As Single, Y As Single) As Long
    Dim Z As Long
    If UseDX Then
'        For Z = 1 To 4
'            With DX.Surface(Z)
'                If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then CheckArea = Z: Exit Function
'            End With
'        Next Z
    Else
        For Z = 1 To 4
            With PKMNImage(Z)
                If X >= .X And X <= .X + picPKMNImage(Z).Width And Y >= .Y And Y <= .Y + picPKMNImage(Z).Height Then CheckArea = Z: Exit Function
            End With
        Next
    End If
    CheckArea = 0
End Function
Private Sub ClearFlash()
    Dim X As Long
    Dim A As Boolean
    If UseDX Then
'        For X = 1 To 4
'            DoFlash(X) = False
'            DX.Surface(X).Visible = DX.Surface(X).RealVis
'            SelPKMN(X).Visible = False
'        Next X
'        DX.Surface(5).Visible = DX.Surface(5).RealVis
'        DX.Surface(6).Visible = DX.Surface(6).RealVis
    Else
        For X = 1 To 4
            DoFlash(X) = False
            If PKMNImage(X).Vis <> PKMNImage(X).OrigVis Then PKMNImage(X).Vis = PKMNImage(X).OrigVis: A = True
            If Shadow(X).Vis <> Shadow(X).OrigVis Then Shadow(X).Vis = Shadow(X).OrigVis: A = True
            SelPKMN(X).Visible = False
        Next X
        If A Then Call RepaintBattleArea
    End If
End Sub

Private Sub BattleArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Long
    Dim B As Boolean
    If Button <> vbLeftButton Then Exit Sub
    Z = CheckArea(X, Y)
    If Z = 0 Then Exit Sub
    If KeyMode = 3 Then
        If SelectedTarget = Z Then Call Targeted(Z)
    Else
        'If UseDX Then B = DX.Surface(Z).Visible Else
        B = PKMNImage(Z).Vis
        If (Z = PNum Or Z = PNum + 2) And B And SelPoke <> 0 Then
            With ThisBattle
                If Not .Ready(Z) Or .GetMoved(Z) Or .GetSwitchTo(Z) > 0 Then
                    Call .UnloadMove(Z)
                    Call .UnloadSwitch(Z)
                    SelMove(Z) = 0
                    SelTarg(Z) = 0
                    SelSwitch(Z) = 0
                    Call SetSelPoke(Z)
                End If
            End With
        End If
    End If
End Sub


Private Sub Chatbox_GotFocus()
    If IsReplayWindow Then cmdQuit.SetFocus
    If txtWatchIgnore.Visible Then cmdLeave.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Call DeactivateTargetMode
    SelMove(SelPoke) = 0
End Sub

Private Sub cmdNext_Click()
    TimerLoops = Slider1.Value
    Call PBTimer_Timer
End Sub

Private Sub cmdPause_Click()
    If Not Started Then Exit Sub
    Started = False
    PBTimer.Enabled = False
    cmdPlay.Enabled = True
    cmdStop.Enabled = True
    cmdPause.Enabled = False
    cmdNext.Visible = True
    Slider1.Enabled = False
    WaitList.Visible = False
End Sub

Private Sub cmdPlay_Click()
    If Pos > UBound(ReplayCommand) Then Exit Sub
    Started = True
    PBTimer.Enabled = True
    cmdPlay.Enabled = False
    cmdPause.Enabled = True
    cmdStop.Enabled = True
    cmdNext.Visible = False
    Slider1.Enabled = True
    WaitList.Visible = True
End Sub

Private Sub cmdQuit_Click()
    Close
    Unload Me
End Sub

Private Sub cmdStop_Click()
    PBTimer.Enabled = False
    Started = False
    Messages.Text = ""
    ThisBattle.ResetBattle
    ThisBattle.DoReplay = True
    Pos = 1
    TimerLoops = Slider1.Value
    Call PBTimer_Timer
    cmdNext.Visible = True
    cmdNext.Enabled = True
    Slider1.Enabled = False
    WaitList.Visible = False
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    cmdPlay.Enabled = True
End Sub

Private Sub cmdSwitchTeam_Click()
    PNum = ONum
    ONum = OtherTeam(PNum)
    Call BattleSync
    Call ReplayPKMNTiles
    Call ReplayRefresh
End Sub

Private Sub cmdTarget_Click(Index As Integer)
    Call Targeted(cmdTarget(Index).Tag)
End Sub

Private Sub cmdTarget_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdTarget(Index).Tag <> SelectedTarget Then
        SelectedTarget = cmdTarget(Index).Tag
        Call ClearFlash
        DoFlash(cmdTarget(Index).Tag) = True
    End If
End Sub


Private Sub cmdUndo_Click()
    cmdUndo.Enabled = False
    If Cancellable Then
        SendData "CANCL"
        Cancellable = False
    End If
End Sub

Private Sub ControlTab_Click(PreviousTab As Integer)
    If ControlTab.Tab = 1 Then RefreshPokeList Else RefreshMoveList
End Sub

Private Sub AttackTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveSel_MouseDown(Index, Button, Shift, X, Y)
End Sub
Private Sub AttackTile_DblClick(Index As Integer)
    If SelPoke = 0 Then Exit Sub
    If Attack.Enabled And KeyMode <> 3 And SelectedMove(SelPoke) = Index + 1 And Not ThisBattle.StruggleOK(SelPoke) Then Call Attack_Click
End Sub

Private Sub ControlTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

'Private Sub DX_AnimationFinished()
'    Dim X As Long
'    For X = DX.NumSurfaces To 7 Step -1
'        DX.DeleteSurface X
'    Next X
'    tmrDelay.Enabled = False
'    'Call ThisBattle.ParseBattle("")
'End Sub

Private Sub Entry_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SwitchTile_MouseMove(Index, Button, 0, 0, 0)
End Sub

Private Sub FlashTimer_Timer()
    Static B As Boolean
    Dim A As Boolean
    Dim X As Long
    Dim Y As Long
    A = False
    B = Not B
    For X = 1 To 4
        If DoFlash(X) Then
            If UseDX Then
'                With DX.Surface(X)
'                    SelPKMN(X).Visible = .Visible
'                    .Visible = Not .Visible
'                    If X = ONum Then
'                        If .Visible Then
'                            DX.Surface(5).Visible = DX.Surface(5).RealVis
'                        Else
'                            DX.Surface(5).Visible = False
'                        End If
'                    ElseIf X = ONum + 2 Then
'                        If .Visible Then
'                            DX.Surface(6).Visible = DX.Surface(6).RealVis
'                        Else
'                            DX.Surface(6).Visible = False
'                        End If
'                    End If
'                End With
            Else
                PKMNImage(X).Vis = Not PKMNImage(X).Vis
                If PKMNImage(X).Vis Then
                    Shadow(X).Vis = Shadow(X).OrigVis
                Else
                    Shadow(X).Vis = False
                End If
                A = True
                SelPKMN(X).Visible = Not PKMNImage(X).Vis
            End If
        End If
    Next X
                    
                        
    If B Then
        If UseDX Then
'            For X = PNum To 4 Step 2
'                Y = (64 - DX.Surface(X).Height) + OriginalTop(X)
'                If DX.Surface(X).Top <> Y Then
'                    DX.Surface(X).Top = Y
'                ElseIf SelPoke = X Then
'                    If cmdLeave.Caption <> "&Leave" Then
'                        DX.Surface(X).Top = DX.Surface(X).Top + 1
'                    End If
'                End If
'            Next X
        Else
            For X = PNum To 4 Step 2
                Y = (960 - picPKMNImage(X).Height) + OriginalTop(X)
                If PKMNImage(X).Y <> Y Then
                    PKMNImage(X).Y = Y: A = True
                ElseIf SelPoke = X Then
                    If cmdLeave.Caption <> "&Leave" Then
                        PKMNImage(X).Y = Y + Screen.TwipsPerPixelY: A = True
                    End If
                End If
            Next X
        End If
    End If
    
    If A And Not UseDX Then Call RepaintBattleArea
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TempTarget As Byte
    Dim X As Byte
    Dim TempOrder(0 To 3) As Byte
    On Error Resume Next
    If KeyCode = vbKeyControl Then
        SkipDelay = True
    End If
    
    'If Me.ActiveControl.Name = "Chatbox" Then Exit Sub
    Select Case KeyCode
        Case vbKeyNumpad8, vbKeyNumpad2, vbKeyNumpad4, vbKeyNumpad6, vbKeyNumpad0, vbKeyNumpad5, vbKeySpace
            If (Shift And vbCtrlMask) = 0 Then Exit Sub
        Case Else
            Exit Sub
    End Select
    If KeyMode < 3 Then
        'If KeyCode = vbKeySpace Or KeyCode = vbKeyNumpad5 Then Exit Sub
        KeyMode = ControlTab.Tab + 1
    End If
    Select Case KeyMode
        Case 1
            X = SelectedMove(SelPoke)
            Select Case X
                Case 1
                    If KeyCode = vbKeyNumpad6 Then SelectedMove(SelPoke) = 2
                    If KeyCode = vbKeyNumpad2 Then SelectedMove(SelPoke) = 3
                Case 2
                    If KeyCode = vbKeyNumpad4 Then SelectedMove(SelPoke) = 1
                    If KeyCode = vbKeyNumpad2 Then SelectedMove(SelPoke) = 4
                Case 3
                    If KeyCode = vbKeyNumpad6 Then SelectedMove(SelPoke) = 4
                    If KeyCode = vbKeyNumpad8 Then SelectedMove(SelPoke) = 1
                Case 4
                    If KeyCode = vbKeyNumpad4 Then SelectedMove(SelPoke) = 3
                    If KeyCode = vbKeyNumpad8 Then SelectedMove(SelPoke) = 2
            End Select
            Attack.SetFocus
            If X <> SelectedMove(SelPoke) Then Call RefreshMoveList 'END EDIT
            If KeyCode = vbKeyNumpad0 Then ControlTab.Tab = 1: Switch.SetFocus
            If (KeyCode = vbKeySpace Or KeyCode = vbKeyNumpad5) And Attack.Enabled = True Then Call Attack_Click
        Case 2
            X = SelectedPKMN
            Select Case SelectedPKMN
                Case 1
                    If KeyCode = vbKeyNumpad6 Then SelectedPKMN = 2
                    If KeyCode = vbKeyNumpad2 Then SelectedPKMN = 4
                Case 2
                    If KeyCode = vbKeyNumpad4 Then SelectedPKMN = 1
                    If KeyCode = vbKeyNumpad6 Then SelectedPKMN = 3
                    If KeyCode = vbKeyNumpad2 Then SelectedPKMN = 5
                Case 3
                    If KeyCode = vbKeyNumpad4 Then SelectedPKMN = 2
                    If KeyCode = vbKeyNumpad2 Then SelectedPKMN = 6
                Case 4
                    If KeyCode = vbKeyNumpad6 Then SelectedPKMN = 5
                    If KeyCode = vbKeyNumpad8 Then SelectedPKMN = 1
                Case 5
                    If KeyCode = vbKeyNumpad4 Then SelectedPKMN = 4
                    If KeyCode = vbKeyNumpad6 Then SelectedPKMN = 6
                    If KeyCode = vbKeyNumpad8 Then SelectedPKMN = 2
                Case 6
                    If KeyCode = vbKeyNumpad4 Then SelectedPKMN = 5
                    If KeyCode = vbKeyNumpad8 Then SelectedPKMN = 3
            End Select
            Switch.SetFocus
            If X <> SelectedPKMN Then Call RefreshPokeList
            If KeyCode = vbKeyNumpad0 Then ControlTab.Tab = 0: Attack.SetFocus
            If (KeyCode = vbKeySpace Or KeyCode = vbKeyNumpad5) And Switch.Enabled = True Then Call Switch_Click
        Case 3
            If KeyCode = vbKeyNumpad0 Or SelectedTarget = 0 Then Exit Sub
            TempOrder(0) = PNum
            TempOrder(1) = PNum + 2
            TempOrder(2) = ONum
            TempOrder(3) = ONum + 2
            TempTarget = SelectedTarget
            For X = 0 To 3
                If TempOrder(X) = SelectedTarget Then Exit For
            Next X
            Do
                If KeyCode = vbKeyNumpad8 Or KeyCode = vbKeyNumpad4 Then
                    X = (X + 1) Mod 4
                ElseIf KeyCode = vbKeyNumpad2 Or KeyCode = vbKeyNumpad6 Then
                    X = (X + 3) Mod 4
                End If
                TempTarget = TempOrder(X)
            Loop Until ValidTarget(TempTarget)
            If TempTarget <> SelectedTarget Then
                Call ClearFlash
                SelectedTarget = TempTarget
                DoFlash(SelectedTarget) = True
                Call FlashTimer_Timer
            End If
            If KeyCode = vbKeySpace Or KeyCode = vbKeyNumpad5 Then
                Call Targeted(SelectedTarget)
            End If
    End Select
    KeyCode = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    SkipDelay = mnuOptionsItem(3).Checked
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Display1 Or Display2 Then
        TT.Destroy
        TT.Tag = ""
        Display1 = False
        Display2 = False
    End If
End Sub

'Private Sub Form_Resize()
'    If Me.WindowState = vbMinimized Then Exit Sub
'    If Me.Width < MinWidth Then Me.Width = MinWidth
'    If Me.Height < MinHeight Then Me.Height = MinHeight
'    If IsReplayWindow Then
'        ReplayControls.Top = Me.Height - 4380
'        ChatBox.Left = Me.Width + (Screen.TwipsPerPixelX * 32)
'    Else
'        ChatBox.Top = Me.Height - 3765
'        ChatBox.Width = Me.Width - 4005
'        SendMsg.Left = ChatBox.Left + (ChatBox.Width - SendMsg.Width)
'        SendMsg.Top = Me.Height - 3165
'        KillConn.Top = Me.Height - 3165
'        cmdLeave.Top = Me.Height - 3165
'    End If
'    Messages.Width = Me.Width - 4005
'    Messages.Height = Me.Height - 3870
'    ControlTab.Top = Me.Height - 3045
'    ControlTab.Width = Me.Width - 375
'    OldSwitchFrame.Top = Me.Height - 2695
'    OldAttackFrame.Top = Me.Height - 2695
'End Sub

Private Sub HPBar_Click(Index As Integer)
    If Not ImJustWatching And Index Mod 2 = PNum Mod 2 Then
        If HPBar(Index).Caption = nbExact Then
            HPBar(Index).Caption = nbPercent
        Else
            HPBar(Index).Caption = nbExact
        End If
        HPBar(Index).RefreshBar
    End If
End Sub

Private Sub mnuReplayFileItem_Click(Index As Integer)
    Dim X As Byte
    Dim Temp As String
    
    Select Case Index
        'Open
        Case 0
            If Started Then Call cmdStop_Click
            cmdStop.Enabled = False
            cmdPause.Enabled = False
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
                    'If Switched Then Call cmdSwitchTeam_Click
                    cmdNext.Enabled = True
                    cmdPlay.Enabled = True
                    cmdPause.Enabled = False
                    cmdStop.Enabled = False
                    cmdSwitchTeam.Enabled = True
                Else
                    cmdNext.Enabled = False
                    cmdPlay.Enabled = False
                    cmdPause.Enabled = False
                    cmdStop.Enabled = False
                End If
            End If
Cancelled:
        Case 2
            Loader.Show
            Unload Me
        Case 3
            Unload MainContainer
    End Select
End Sub

Private Sub MoveSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If MoveSel(Index).Picture = BadImage.Picture Then Exit Sub
    If SelectedMove(SelPoke) = Index + 1 Then Exit Sub
    SelectedMove(SelPoke) = Index + 1
    Call RefreshMoveList
End Sub

Private Sub OpponentPokemon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Integer
'    If UseDX Then
'        If DX.Animating Then Exit Sub
'    End If
    If BattleCurrent(ONum).No = 0 Or Display2 Then Exit Sub
    Display1 = False
    Display2 = True
    For Z = 0 To 5
        Display(Z) = False
    Next Z
    TT.Title = BattleCurrent(ONum).Nickname
    TT.TipText = StoredTT2
    TT.Style = TTStandard
    TT.Create BattleArea.hWnd
End Sub

Private Sub BattleArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim Z As Long
'    If UseDX Then
'        If DX.Animating Then Exit Sub
'    End If
    Z = CheckArea(X, Y)
    If Z > 0 Then
    
        With TT
            If BattleCurrent(Z).No = 0 Then
                If .Tag <> "" Then
                    .Destroy
                    .Tag = ""
                End If
            Else
                If .Tag <> "C" & Z Then
                    .Destroy
                    .Tag = "C" & Z
                    '.Style = TTBalloon
                    .Title = BattleCurrent(Z).Nickname
                    .TipText = CurrentTT(Z)
                    .Create BattleArea.hWnd
                End If
            End If
        End With
        
        If Z <> SelectedTarget And ValidTarget(Z) Then
            SelectedTarget = Z
            Call ClearFlash
            DoFlash(Z) = True
        End If
    Else
        If TT.Tag <> "" Then
            TT.Destroy
            TT.Tag = ""
        End If
    End If
End Sub

'Private Sub PKMNImage_Click(Index As Integer)
'    If KeyMode = 3 And ValidTarget(Index) Then Call Targeted(Index)
'End Sub

Private Sub PBTimer_Timer()
    Dim Temp As String
    Dim E As Boolean
    TimerLoops = TimerLoops + 1
    If TimerLoops >= Slider1.Value Then
        E = PBTimer.Enabled
        PBTimer.Enabled = False
        If Pos > UBound(ReplayCommand) Then Exit Sub
        TimerLoops = 0
        Do
            Call DoReplayCommand
            Pos = Pos + 1
            If Pos > UBound(ReplayCommand) Then Exit Do
        Loop Until Left(ReplayCommand(Pos), 5) = "BCMD:"
        If Pos = UBound(ReplayCommand) + 1 Then
            cmdPlay.Enabled = False
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            Started = False
            cmdNext.Visible = True
            cmdNext.Enabled = False
            Slider1.Enabled = False
            WaitList.Visible = False
            Call AddMessage(vbNewLine & "End of replay.", , , , True)
            If ThisBattle.BattleOver Then
                Call StartMusic
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
    WaitList.Caption = "Wait Time " & Slider1.Value & " Seconds"
End Sub

Private Sub SwitchTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Entry_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub SwitchTile_DblClick(Index As Integer)
    If Switch.Enabled And SelectedPKMN = Index + 1 Then Call Switch_Click
End Sub

Private Sub Entry_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Entry(Index).Picture = BadImage.Picture Then Exit Sub
    If SelectedPKMN = Index + 1 Then Exit Sub
    If IsReplayWindow Then Exit Sub
    SelectedPKMN = Index + 1
    RefreshPokeList
End Sub

Public Sub HPAnimTimer_Timer()
    Dim X As Byte
    
    For X = 1 To ThisBattle.ActNum
        If BattleCurrent(X).MaxHP = 0 Then
            If HPBar(X).Value > 0 Then
                HPBar(X).Value = 0
                HPBar(X).Max = 0
            End If
        Else
            If AnimOption = 0 Then
                HPBar(X).Max = BattleCurrent(X).MaxHP
                Call ChangeHPBar(X, BattleCurrent(X).HP)
            Else
                If HPBar(X).Max <> BattleCurrent(X).MaxHP Then
                    HPBar(X).Value = 0
                    HPBar(X).Max = BattleCurrent(X).MaxHP
                    Call ChangeHPBar(X, BattleCurrent(X).HP)
                End If
                If HPBar(X).Value <= BattleCurrent(X).HP - 10 Then
                    Call ChangeHPBar(X, 10, True)
                ElseIf HPBar(X).Value > BattleCurrent(X).HP - 10 And HPBar(X).Value < BattleCurrent(X).HP Then
                    Call ChangeHPBar(X, BattleCurrent(X).HP)
                ElseIf HPBar(X).Value >= BattleCurrent(X).HP + 10 Then
                    Call ChangeHPBar(X, -10, True)
                ElseIf HPBar(X).Value < BattleCurrent(X).HP + 10 And HPBar(X).Value > BattleCurrent(X).HP Then
                    Call ChangeHPBar(X, BattleCurrent(X).HP)
                End If
            End If
        End If
    Next
End Sub
Private Sub Messages_Click()
    HideCaret Messages.hWnd
End Sub

Private Sub messages_GotFocus()
    HideCaret Messages.hWnd
End Sub
  
Private Sub messages_LostFocus()
    ShowCaret Messages.hWnd
End Sub
  
Private Sub messages_KeyDown(KeyCode As Integer, Shift As Integer)
    'Scroll the TextBox if appropriate
    Select Case KeyCode
        Case vbKeyDown
            'Scroll the text up
            VScrollTextBox Messages, True, False
        Case vbKeyUp
            'Scroll the text down
            VScrollTextBox Messages, False, False
        Case vbKeyPageDown
            'Scroll the text up
            VScrollTextBox Messages, True, True
        Case vbKeyPageUp
            'Scroll the text down
            VScrollTextBox Messages, False, True
    End Select
End Sub
 
Public Sub VScrollTextBox(ByRef TBox As RichTextBox, ByVal ScrollDown As Boolean, ByVal PageMode As Boolean)
    Dim lParam As Long
 
    'Determine which scroll type to perform
    If PageMode Then
        If ScrollDown Then
            lParam = SB_PAGEDOWN
        Else
            lParam = SB_PAGEUP
        End If
    Else
        If ScrollDown Then
            lParam = SB_LINEDOWN
        Else
            lParam = SB_LINEUP
        End If
    End If
 
    'Scroll the TextBox
    Call SendMessage(TBox.hWnd, WM_VSCROLL, lParam, 0)
End Sub

Private Sub Attack_Click()
    Dim R As Integer
    Dim M As Integer
    Dim X As Integer
    Dim Y As Byte
    Dim Z As Byte
    Dim Temp As String
        
    'Just in case...
    If Not ThisBattle.CanAttack(SelPoke) Then Exit Sub
    
    'Use the selected move if PP remains
    If ThisBattle.StruggleOK(SelPoke) Then
        M = 0
        R = 210
    Else
        M = -1
        If BattleCurrent(SelPoke).Move(SelectedMove(SelPoke)) > 0 Then M = SelectedMove(SelPoke)
        If M = -1 Then
            MsgBox "Please select a move!", vbCritical, "Error!"
            Exit Sub
        End If
        Temp = ThisBattle.CanUseMove(SelPoke, M)
        If Temp <> "" Then
            MsgBox Temp, vbExclamation, "Illegal Move"
            Exit Sub
        End If
        R = BattleCurrent(SelPoke).Move(M)
    End If
    
    'User chose the move, now it's time for the target.
    'First makes sure there's more than one target to choose from.
    SelMove(SelPoke) = M
    Y = 0
    For X = 1 To 4
        If BattleCurrent(X).HP > 0 And X <> SelPoke Then Z = X: Y = Y + 1
    Next X
    If Y > 1 And BattleCondition(SelPoke).EncoreMove = 0 And (Moves(R).Target = nbSelectedTarget Or ((BattleCurrent(SelPoke).Type1 = 14 Or BattleCurrent(SelPoke).Type2 = 14) And R = 38)) Then
        Call ActivateTargetMode(SelPoke)
    Else
        Call ThisBattle.LoadMove(SelPoke, M, Z, True)
        SelTarg(SelPoke) = Z
        Call SendBattle
    End If
    
'    'What to do with the data
'    Call SendData("MOVE:" & M)
'    StatusBar1.Panels(3).Text = "Network Status: Sent command"
'    Attack.Enabled = False
'    Switch.Enabled = False
'    'Exit sub if the move was Struggle to prevent crashing
'    If M = 0 Then Exit Sub
'    'Baton Pass handling
'    '(Handled elsewhere now)
''    If Moves(BattleCurrent(PNum).Move(M)).SpecialEffect = 11 And Fainted < 5 Then
''        Switch.Enabled = True
''        StatusBar1.Panels(3).Text = "Please select the Pokémon to switch to"
''        Switch.Caption = "&Baton Pass"
''        ControlTab.Tab = 1
''        Exit Sub
''    End If
End Sub

Private Sub cmdLeave_Click()
    Dim Temp As Integer
    If ImJustWatching Then
        Temp = MsgBox("Are you sure you want to stop watching this battle?", vbYesNo + vbQuestion, "Are you sure?")
    Else
        If cmdLeave.Caption = "&Leave" Then
            If RelayServer Then
                Temp = MsgBox("Close the battle window and return to the server?", vbYesNo + vbDefaultButton2 + vbQuestion, "Are you sure?")
            Else
                Temp = MsgBox("Are you sure you want to kill the connection and end the game?", vbYesNo + vbDefaultButton2 + vbQuestion, "Are you sure?")
            End If
        Else
            If ThisBattle.Unrated Then
                Temp = MsgBox("Are you sure you want to forfeit?", vbYesNo + vbDefaultButton2 + vbQuestion, "Are you sure?")
            Else
                Temp = MsgBox("Are you sure you want to forfeit?" + vbNewLine + "This will count as a Loss on your battle record.", vbYesNo + vbDefaultButton2 + vbQuestion, "Are you sure?")
            End If
        End If
    End If
    If Temp = vbYes Then
        FinalExit = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim X As Long
    Dim Y As Long
    Dim sFile As String
    Dim Temp As String
    On Error Resume Next
    WriteDebugLog "Opening Battle Window"
    'Call DoInitialResize
    
    If OpenedAsReplay Then
        IsReplayWindow = True
        OpenedAsReplay = False
    Else
        IsReplayWindow = False
    End If
    
    WriteDebugLog "Creating classes"
    Set RTB = New RTBClass
    'RTB.SetRTBHook Messages, Chatbox, MinWidth, MinHeight
    RTB.SetRTBHook Messages, ChatBox
    Set ThisBattle = New BattleData
    Set ThisBattle.BattleWindow = Me
    Set ThisBattle.LogOutput = RTB
    Call ThisBattle.ResetBattle
    Set TT = New CTooltip
    TT.Icon = TTIconInfo
    TT.DelayTime = 400
    TT.VisibleTime = 32767
    
    'Set DX = New clsDX
    WriteDebugLog "Calling StopMusic"
    Call StopMusic
    WriteDebugLog "Setting BG"
    If UseBG Or UseDX Then
        Call MainContainer.DoPicture("bg1.gif")
        picTerrain.Picture = MainContainer.SwapSpace.Picture
    End If
    
    If UseDX Then
'        DX.InitDX BattleArea
'        DX.CreateSurfaceFromPBox picTerrain
'
'        'DX.CreateSolidColorSurface RGB(255, 0, 0), 112, 128
'        'DX.Surface(0).Move 1, 1
'        DX.CreateSolidColorSurface vbWhite, 64, 64, vbWhite
'        DX.CreateSolidColorSurface vbWhite, 64, 64, vbWhite
'        DX.CreateSolidColorSurface vbWhite, 64, 64, vbWhite
'        DX.CreateSolidColorSurface vbWhite, 64, 64, vbWhite
'        DX.CreateSurfaceFromPBox picShadow, -2
'        DX.CreateSurfaceFromPBox picShadow, -2
'        DX.Surface(5).Visible = False
'        DX.Surface(6).Visible = False
    Else
        WriteDebugLog "Making masks"
        Call CreateMask(picShadow, picShadowMask)
        picBuild.Width = BattleArea.Width
        picBuild.Height = BattleArea.Height
        For X = 1 To 4
            Load picPKMNImage(X)
            Load picPKMNMask(X)
            PKMNImage(X).Vis = False
            Shadow(X).Vis = False
        Next X
    End If

    
    If IsReplayWindow Then
        ThisBattle.ResetBattle
        PNum = 1
        ONum = 2
        Call RearrangeForm(2, False, True)
    ElseIf MasterServer.WatchID <> "" Then
        WriteDebugLog "Setting Player Data"
        ImJustWatching = True
        Temp = MasterServer.WatchID
        MasterServer.WatchID = ""
        WatchID = Dec(ChopString(Temp, 3))
        X = Dec(ChopString(Temp, 3))
        Me.WatchP1 = X
        ThisBattle.PlayerName(1) = Player(X).Name
        YourName = Player(X).Name
        ThisBattle.Player1 = X
        TIcon(0).Picture = MainContainer.Trainers.ListImages(Player(X).Picture).Picture
        X = Dec(ChopString(Temp, 3))
        Me.WatchP2 = X
        ThisBattle.PlayerName(2) = Player(X).Name
        OpponentName = Player(X).Name
        ThisBattle.Player2 = X
        TIcon(1).Picture = MainContainer.Trainers.ListImages(Player(X).Picture).Picture
        Me.Caption = YourName & " vs " & OpponentName
        mnuOptionsItem(4).Enabled = False
        X = Dec(ChopString(Temp, 1))
        ThisBattle.ActNum = X
'        If X = 0 Then
'            ChatBox.Text = ""
'            ChatBox.Enabled = True
'            SendMsg.Enabled = True
'        Else
'            ChatBox.Text = "Spectator chat has been disallowed for this battle."
'            ChatBox.Enabled = False
'            SendMsg.Enabled = False
'        End If
        ThisBattle.WatchFormID = Dec(ChopString(Temp, 3))
        PNum = 1
        ONum = 2
        'Set up the Watch Window
        'NOTE: It should know by this point whether this is a 1-on-1 or 2-on-2 battle.
        WriteDebugLog "Rearraging"
        Call RearrangeForm(ThisBattle.ActNum, True)
    Else
        WriteDebugLog "Setting Player Info"
        ImJustWatching = False
        BattleID = Asc(ChopString(BattleTemp, 1))
        PNum = Val(ChopString(BattleTemp, 1))
        ONum = OtherTeam(PNum)
        ThisBattle.ActNum = Val(ChopString(BattleTemp, 1))
        BattleTemp = 0
        ThisBattle.WatchFormID = 0
        ControlTab.Enabled = False
        mnuOptionsItem(4).Enabled = True
        WriteDebugLog "Clearing Poke Variables"
        For X = 1 To 6
            BattlePKMN(PNum, X) = PKMN(X)
            BattlePKMN(ONum, X) = BlankPKMN
        Next
        WriteDebugLog "Setting More Player Info"
        Battling = True
        ChallengePending = False
        TIcon(0).Picture = MainContainer.Trainers.ListImages(You.Picture).Picture
        YourName = Player(YourNumber).Name
        If LogPrompt = 1 Then mnuBattleItem(5).Checked = True
        If LogSave = 1 Then mnuBattleItem(5).Enabled = False
        If ReplayPrompt = 1 Then mnuBattleItem(4).Checked = True
        If Autosave = 1 Then mnuBattleItem(4).Enabled = False
        WriteDebugLog "Rearraging"
        Call RearrangeForm(ThisBattle.ActNum, False)
    End If
    WriteDebugLog "Centering"
    Call CenterWindow(Me)
    If DebugMode Then mnuDebug.Visible = True
    Call ThisBattle.IsReady(1)
    Call ThisBattle.IsReady(2)
    'Reset the variables because VB is stupid and doesn't seem to unload properly.
    WriteDebugLog "Resetting Main Variables"
    SelPoke = 0
    tmrDelay.Interval = MoveDelay
    tmrDelay.Enabled = False
    StatusBar1.Panels(1).Text = ""
    StoredRandTeam = ""
    SelectedPKMN = 1
    SyncString = ""
    Resuming = False
    If UseDX Then
        mnuOptionsItem(3).Checked = True
        mnuOptionsItem(3).Enabled = False
        SkipDelay = True
    Else
        SkipDelay = False
    End If
    ReDim BattleQueue(0)
    WriteDebugLog "Resetting SelectedMove"
    For X = 1 To 4
        SelectedMove(X) = 1
    Next X
    WriteDebugLog "Resetting More Variables"
    DrawCount = 0
    UnrateCount = 0
    NumPoke = 6
    TurnNumber = 0
    Attack.Enabled = False
    Switch.Enabled = False
    FinalExit = False
    ReceivedBOVER = False
    StartedTransfer = False
    cmdUndo.Enabled = False
    For X = 1 To 1024
        NetworkStuff(X) = ""
    Next X
    For X = 1 To 4
        BattleCurrent(X) = BlankPKMN
        BattleCondition(X) = BlankCondition
        BattleTC(X) = BlankTC
    Next X
    For X = 1 To 2
        For Y = 1 To 6
            BattlePKMN(X, Y) = BlankPKMN
        Next Y
    Next X
    
    BattleWeather = 0
    'ThisBattle.DoMessages = True
    PNum = 1
    WriteDebugLog "Finalizing"
    If Not ImJustWatching And Not IsReplayWindow Then
'        For X = 1 To 6
'            BattlePKMN(PNum, X) = PKMN(X)
'            BattlePKMN(PNum, X).HP = BattlePKMN(PNum, X).MaxHP
'        Next
'        BattleCurrent(PNum) = PKMN(1)
'        BattleCurrent(PNum).TeamNumber = 1
'        Call RefreshPokeList
'        Call RefreshMoveList
        'Play music and set up the menus
        'If MusicOption = 1 Then Call PlayMusic(4, True)
        If SoundOption = 1 Then mnuOptionsItem(0).Checked = True
        If MusicOption = 1 Then mnuOptionsItem(1).Checked = True
        'Do the connection
        MasterServer.mnuOptionsItem(5).Enabled = False
        cmdLeave.Caption = "&Forfeit"
        mnuBattle.Visible = True
        Call SendData("READY")
        WriteDebugLog "READY Sent"
        Unload ChallengeWindow
    ElseIf Not IsReplayWindow Then
        Call MasterServer.SendData("WWRDY")
    End If
    WriteDebugLog "Load complete"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Temp As Integer
    
    If IsReplayWindow Then Exit Sub
    If UnloadMode = 1 Then
        If FinalExit Then
            If RelayServer Then
                Call MasterServer.SendData("EXIT:")
                Unload MainContainer
            Else
                Call SendData("EXIT")
                Unload MainContainer
            End If
        End If
        Exit Sub
    Else
        If ImJustWatching Then
            Temp = MsgBox("Are you sure you want to stop watching this battle?", vbYesNo + vbQuestion, "Are you sure?")
        Else
            Temp = MsgBox("Are you sure you want to end the game?", vbYesNo + vbQuestion, "Are you sure?")
        End If
        Select Case Temp
            Case vbYes
                If FinalExit Then
                    If RelayServer Then
                        Call MasterServer.SendData("EXIT:")
                        End
                    Else
                        Call SendData("EXIT")
                        End
                    End If
                End If
                Unload Me
            Case vbNo
                Cancel = True
                Exit Sub
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Answer As Integer
    Dim X As Integer
    Dim Y As Long
    Dim Command As String
    Dim Data As String
    Dim SaveNum As Integer
    Dim Temp
    Dim SavedLog As Boolean
    Dim SavedReplay As Boolean
    
    If IsReplayWindow Then
        RTB.UnsetRTBHook
        Call StopMusic
        Close
        Loader.Show
        Exit Sub
    End If
    If ImJustWatching Then
        WatchLoaded(ThisBattle.WatchFormID) = False
        Call MasterServer.SendData("DONW:" & WatchID)
        WatchID = 0
        If mnuBattleItem(5).Checked Then
            With MainContainer.FileBox
                .DialogTitle = "Save Log"
                .Filter = "Log Files (*.txt)|*.txt"
                .Flags = cdlOFNOverwritePrompt
                .CancelError = True
                .DefaultExt = ".txt"
                .FileName = YourName & " vs " & OpponentName & ".txt"
                Temp = GetSetting("NetBattle", "Options", "InitDir", "")
                If Temp <> "" Then .InitDir = Temp
                On Error GoTo Cancelled
                .ShowSave
                Data = .FileName
                SaveSetting "NetBattle", "Options", "InitDir", Left$(Data, InStrRev(Data, "\"))
                If Data <> "" Then
                    If Right(Data, 4) <> ".txt" Then Data = Data & ".txt" 'It's been screwing up lately...
                    Call SaveLog(Data)
                End If
            End With
        End If
        Exit Sub
    End If
    On Error Resume Next
    MasterServer.UnloadingBattle = True
    SaveNum = FreeFile
    If LogSave And Autosave Then
        For Y = 0 To 9999998
            If Not FileExists(SlashPath & OpponentName & Format(Y, "0000000") & ".txt") And Not FileExists(SlashPath & OpponentName & Format(Y, "0000000") & ".btl") Then Exit For
        Next
        Call SaveLog(SlashPath & OpponentName & Format(Y, "0000000") & ".txt")
        Call MasterServer.AddMessage("Log saved to " & SlashPath & OpponentName & Format(Y, "0000000") & ".txt", , , , , True)
        Call SaveReplay(SlashPath & OpponentName & Format(Y, "0000000") & ".btl")
        Call MasterServer.AddMessage("Replay saved to " & SlashPath & OpponentName & Format(Y, "0000000") & ".btl", , , , , True)
        SavedLog = True
        SavedReplay = True
    ElseIf LogSave = 1 Then
        For Y = 0 To 9999998
            If Not FileExists(SlashPath & OpponentName & Format(Y, "0000000") & ".txt") Then Exit For
        Next
        Call SaveLog(SlashPath & OpponentName & Format(Y, "0000000") & ".txt")
        Call MasterServer.AddMessage("Log saved to " & SlashPath & OpponentName & Format(Y, "0000000") & ".txt", , , , , True)
        SavedLog = True
    ElseIf Autosave Then
        For Y = 1 To 9999998
            If Not FileExists(SlashPath & OpponentName & Format(Y, "0000000") & ".btl") Then Exit For
        Next
        Close #ReplayNum
        Call SaveReplay(SlashPath & OpponentName & Format(Y, "0000000") & ".btl")
        Call MasterServer.AddMessage("Replay saved to " & SlashPath & OpponentName & Format(Y, "0000000") & ".btl", , , , , True)
        SavedReplay = True
    End If
    If mnuBattleItem(5).Checked And Not SavedLog Then
        With MainContainer.FileBox
            .DialogTitle = "Save Log"
            .Filter = "Log Files (*.txt)|*.txt"
            .Flags = cdlOFNOverwritePrompt
            .CancelError = True
            .DefaultExt = ".txt"
            .FileName = ""
            Temp = GetSetting("NetBattle", "Options", "InitDir", "")
            If Temp <> "" Then .InitDir = Temp
            On Error GoTo Cancelled
            .ShowSave
            Data = .FileName
            SaveSetting "NetBattle", "Options", "InitDir", Left$(Data, InStrRev(Data, "\"))
            If Data <> "" Then
                If Right(Data, 4) <> ".txt" Then Data = Data & ".txt" 'It's been screwing up lately...
                Call SaveLog(Data)
            End If
        End With
    End If
    If mnuBattleItem(4).Checked And Not SavedReplay Then
        Close #ReplayNum
        With MainContainer.FileBox
            .DialogTitle = "Save Replay"
            .Filter = "Battle Files (*.btl)|*.btl"
            .Flags = cdlOFNOverwritePrompt
            .CancelError = True
            .DefaultExt = ".btl"
            .FileName = ""
            Temp = GetSetting("NetBattle", "Options", "InitDir", "")
            If Temp <> "" Then .InitDir = Temp
            On Error GoTo Cancelled
            .ShowSave
            Data = .FileName
            SaveSetting "NetBattle", "Options", "InitDir", Left$(Data, InStrRev(Data, "\"))
            If Data <> "" Then
                If Right(Data, 4) <> ".btl" Then Data = Data & ".btl" 'It's been screwing up lately...
                Call SaveReplay(Data)
            End If
        End With
    End If
Cancelled:
    If Me.WindowState <> vbMinimized Then
        If Me.WindowState = vbMaximized Then
            SaveSetting "NetBattle", "Battle Window", "Maximized", True
        Else
            SaveSetting "NetBattle", "Battle Window", "Maximized", False
        End If
        SaveSetting "NetBattle", "Battle Window", "Width", Me.Width
        SaveSetting "NetBattle", "Battle Window", "Height", Me.Height
    End If
    tmrDelay.Enabled = False
    Err.Clear
    Err.Number = 0
    On Error Resume Next
    RTB.UnsetRTBHook
    Close #ReplayNum
    If FileExists(TFileName) And TFileName <> "" Then Kill TFileName
    StopMusic
    Battling = False
    MasterServer.mnuOptionsItem(5).Enabled = True
    'If Not FinalExit And Not ReceivedBOVER Then Call SendData("BOVER")
    Call MasterServer.SendData("BACK:")
    Unload Stadium
    MasterServer.mnuTeam.Enabled = True
    MasterServer.UnloadingBattle = False
    'Set DX = Nothing
    Set ThisBattle.BattleWindow = Nothing
    Set ThisBattle = Nothing
    If FinalExit Then Unload MainContainer
End Sub

Private Sub KillConn_Click()
    Dim Temp As Integer
    
    Temp = MsgBox("Are you sure you want to kill the connection and exit the program?", vbYesNo + vbDefaultButton2 + vbQuestion, "Are you sure?")
    Select Case Temp
        Case vbYes
            FinalExit = True
            Call MasterServer.SendData("EXIT:")
            End
        Case vbNo
            Exit Sub
    End Select
End Sub

Private Sub mnuBattleItem_Click(Index As Integer)
    Dim Temp As Integer

    Select Case Index
        'Forfeit
        Case 0
            Call cmdLeave_Click
        'Draw
        Case 1
            If DrawCount > 2 Then MsgBox "You may only ask for a Tie 3 times.": Exit Sub
            If MsgBox("Propose a that the game end in a Tie?", vbYesNo + vbDefaultButton2, "Propose Tie") = vbYes Then
                DrawCount = DrawCount + 1
                If cmdCancel.Visible Then Call cmdCancel_Click
                Call SendData("REQTI")
                Call LockBattle
            End If
        'Unrate
        Case 2
            If UnrateCount > 2 Then MsgBox "You may only ask to unrate the battle 3 times.": Exit Sub
            If MsgBox("If a battle is unrated, the outcome will not affect either players' battle record." + vbNewLine + "However, both players must agree.  Ask your opponent to unrate the battle?", vbYesNo + vbDefaultButton2, "Unrate Battle") = vbYes Then
                UnrateCount = UnrateCount + 1
                If cmdCancel.Visible Then Call cmdCancel_Click
                Call SendData("REQUN")
                Call LockBattle
            End If
        'Separator
        Case 3
        'Save Replay
        Case 4
            mnuBattleItem(4).Checked = Not mnuBattleItem(4).Checked
            ReplayPrompt = Abs(mnuBattleItem(4).Checked)
            SaveSetting "NetBattle", "Options", "Replay Prompt", ReplayPrompt
        'Save Log
        Case 5
            mnuBattleItem(5).Checked = Not mnuBattleItem(5).Checked
            LogPrompt = Abs(mnuBattleItem(5).Checked)
            SaveSetting "NetBattle", "Options", "Log Prompt", LogPrompt
    End Select
End Sub

Private Sub mnuDebugItem_Click(Index As Integer)
    Dim X As Byte
    Dim Temp As String
    Dim Worked As Boolean
    
    Select Case Index
        Case 0
            For X = 1 To 4
                BattleCurrent(1).PP(X) = 0
                BattleCurrent(2).PP(X) = 0
            Next
        Case 1
            For X = 1 To 6
                BattlePKMN(2, X) = BattlePKMN(PNum, X)
            Next
            BattleCurrent(1) = BattlePKMN(PNum, 1)
            BattleCurrent(2) = BattlePKMN(2, 1)
        Case 2
            Call UpdateImages
            Call UpdateStats
        Case 3
            Call AddMessage("Test Message", True, , vbRed, True)
        Case 4
            'Call AIBotMove
        Case 5
'            If PNum = 0 Then PNum = 1: ONum = 2
'            Worked = ThisBattle.SetVer(PNum, 1)
'            For X = 1 To 6
'                Worked = ThisBattle.SetPKMN(PNum, X, Pkmn2Str(PKMN(X)))
'                If Worked Then
'                    MsgBox Pkmn2Str(PKMN(X)), vbInformation, "#" & X & " - Ok"
'                Else
'                    MsgBox Pkmn2Str(PKMN(X)), vbCritical, "#" & X & " - Failed"
'                End If
'            Next
'            MsgBox ThisBattle.GetTeam(PNum), vbInformation, "Team Data - " & Len(ThisBattle.GetTeam(PNum)) & " Characters"
'            mnuDebugItem(6).Enabled = True
'            mnuDebugItem(7).Enabled = True
'        Case 6
'            For X = 1 To 6
'                Worked = ThisBattle.SetPKMN(ONum, X, Pkmn2Str(PKMN(X)))
'            Next
'            Call BattleSync
'            Call UpdateImages
'            Call UpdateStats
'        Case 7
'            Temp = ThisBattle.GetTeam(PNum)
'            Worked = ThisBattle.SetTeam(ONum, Temp)
'            If Worked Then
'                MsgBox "No problems detected.", vbInformation, "Debug"
'            Else
'                MsgBox "Error in transfer!", vbCritical, "Debug"
'            End If
    End Select
End Sub
Private Sub mnuFileItem_Click(Index As Integer)
    Dim Temp As Integer
    Dim FileToUse As String
    Dim X As Integer
    
    Select Case Index
        Case 0
            Call SaveLog
        Case 1
            Call SaveDump
        Case 3
            Temp = MsgBox("Are you sure you want to end the battle?", vbYesNo + vbQuestion, "Are you sure?")
            Select Case Temp
                Case vbYes
                    FinalExit = False
                    Unload Me
                Case vbNo
                    Exit Sub
            End Select
        Case 4
            Temp = MsgBox("Are you sure you want to exit the program?", vbYesNo + vbQuestion, "Are you sure?")
            Select Case Temp
                Case vbYes
                    FinalExit = True
                    Unload Me
                Case vbNo
                    Exit Sub
            End Select
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute 0, vbNullString, "http://www.netbattle.net", vbNullString, vbNullString, 0
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Private Sub mnuOptionsItem_Click(Index As Integer)
    Dim X As Integer
    
    Select Case Index
        Case 0
            If mnuOptionsItem(0).Checked Then
                mnuOptionsItem(0).Checked = False
                SoundOption = 0
            Else
                mnuOptionsItem(0).Checked = True
                SoundOption = 1
            End If
            SaveSetting "NetBattle", "Options", "ServerSound", SoundOption
        Case 1
            If mnuOptionsItem(1).Checked Then
                mnuOptionsItem(1).Checked = False
                MusicOption = 0
                Call StopMusic
            Else
                mnuOptionsItem(1).Checked = True
                MusicOption = 1
                If Not IsReplayWindow Then
                    Call StartMusic
                Else
                    If CmdFile = "" Then Exit Sub Else Call StartMusic
                End If
            End If
            SaveSetting "NetBattle", "Options", "Music", MusicOption
        Case 3
            mnuOptionsItem(3).Checked = Not mnuOptionsItem(3).Checked
            SkipDelay = mnuOptionsItem(3).Checked
            tmrDelay.Interval = IIf(SkipDelay, 0, MoveDelay)
            If SkipDelay And tmrDelay.Enabled Then Call tmrDelay_Timer
        Case 4
            mnuOptionsItem(4).Enabled = False
            Call SendData("IGWAT")
    End Select
End Sub

Private Sub MoveName_Click(Index As Integer)
    StatusBar1.Panels(3).Text = Moves(BattleCurrent(PNum).Move(Index + 1)).Text
End Sub






Private Sub SendMsg_Click()
    Dim Build As String
    Build = RTrim$(FilterIllegalChars(ChatBox.Text, True))
    If Len(Build) = 0 Then Exit Sub
    If txtWatchIgnore.Visible Then Exit Sub
    If ImJustWatching Then
        Call SendData("WMSG:" & Build)
    Else
        Call AddToReplay("CHAT:" & YourName & ": " & ChatBox.Text)
        Call SendData("CMSG:" & Build)
    End If
    If Left$(Build, 4) = "/me " Then
        Call AddMessage("*** " & Player(YourNumber).Name & " " & Right$(Build, Len(Build) - 4), False, , &HC000C0)
    Else
        Call AddMessage(Player(YourNumber).Name & ": " & Build, False, ":", vbRed, True, False)
    End If
    ChatBox.Text = ""
End Sub



Private Sub Switch_Click()
    Dim p As Integer
    
    If IsReplayWindow Then Exit Sub
    'Just in case...
    If Not ThisBattle.CanSwitch(SelPoke) Then Exit Sub

    p = SelectedPKMN
    If p = BattleCurrent(PNum).TeamNumber _
    Or p = BattleCurrent(PNum + 2).TeamNumber Then
        MsgBox "That Pokémon is already out!", , "Error"
        Exit Sub
    End If
    If BattlePKMN(PNum, p).HP = 0 Then
        MsgBox "There's no will to fight!", , "Error!"
        Exit Sub
    End If
    If ThisBattle.LoadSwitch(SelPoke, p) = False Then
        MsgBox BattlePKMN(PNum, p).Nickname & " has already been selected!", , "Error!"
        Exit Sub
    Else
        SelSwitch(SelPoke) = p
        Call SendBattle
    End If
    
'    Call SendData("PKSW:" & P)
'    StatusBar1.Panels(3).Text = "Network Status: Sent command"
'    TabSwitch = True
'    Switch.Enabled = False
'    Attack.Enabled = False
'    Switch.Caption = "&Switch"
End Sub

Public Sub DoIncoming(ByVal Temp As String)
    'Process incoming data
    Dim Command As String * 5
    Dim Data As String
    Dim OData As String
    Dim Worked As Boolean
    Dim Temp2 As String
    Dim Temp3 As String
    Dim X As Integer
    Dim Y As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim Fainted As Integer
    Dim AttackStatus As Boolean
    Dim SwitchStatus As Boolean
    Dim Answer As Integer
    Dim FileToUse As String
    Dim TempWatch() As String
    Dim TempEffect(1 To 12) As Byte
    tmrBattleQueue.Enabled = False
    Call NetworkLog(Temp, False)
    Call AddMessage("Rcd: " & Temp, True)
    Command = ChopString(Temp, 5)
    Data = Temp
    Call WriteDebugLog("Battle.DoIncoming: " & Command & Data)
    Select Case Command
        'Random Team: For Challenge Cup Mode
        Case "RAND:"
            ThisBattle.RandBat = True
            StoredRandTeam = Data
        'Info: Basically everything.
        Case "INFO:"
            Call InitTFile
            Call AddToReplay("VNUM:" & You.ProgVersion)
            OData = Data
            'ThisBattle.BattleMode = CompatVersion(Player(YourNumber).GameVersion)
            Resuming = (ChopString(Data, 1) = 1)
            If Not Resuming Then ThisBattle.InitLogFile
            ThisBattle.BattleMode = Dec(ChopString(Data, 1))
            ThisBattle.ActNum = Dec(ChopString(Data, 1))
            ThisBattle.Terrain = Dec(ChopString(Data, 1))
            If UseBG Then
                Call MainContainer.DoPicture("bg" & CStr(ThisBattle.Terrain) & ".gif")
                If UseDX Then
                    'DX.Surface(0).UpdatePicture MainContainer.SwapSpace
                Else
                    picTerrain.Picture = MainContainer.SwapSpace.Picture
                End If
            End If
            StatusBar1.Panels(2).Text = TerrainText(ThisBattle.Terrain)
            ThisBattle.Rules = ChopString(Data, 8)
            'If ThisBattle.EnabledRule(nbUnrated) Then mnuBattleItem(2).Enabled = False
            Temp2 = ChopString(Data, 13)
            PNum = Val(ChopString(Data, 1))
            ONum = OtherTeam(PNum)
'            If ThisBattle.EnabledRule(nbExactHP) Then
'                HPBar(ONum).Caption = nbExact
'                If ThisBattle.ActNum = 4 Then HPBar(ONum + 2).Caption = nbExact
'            End If
            X = Asc(ChopString(Data, 1))
            If Resuming Then
                Player(YourNumber).BattlingWith = X
                Call MasterServer.RefreshListing
            End If
            ThisBattle.ThisPNum = PNum
            If PNum = 1 Then
                ThisBattle.Player1 = YourNumber
                ThisBattle.Player2 = X
            Else
                ThisBattle.Player2 = YourNumber
                ThisBattle.Player1 = X
            End If
            Opponent.Picture = Player(X).Picture
            TIcon(1).Picture = MainContainer.Trainers.ListImages(Opponent.Picture).Picture
            Opponent.Version = Player(X).GFXVer
            OpponentName = Player(X).Name
            Call AddToReplay("XINF:" & FixedHex(You.Picture, 2) & FixedHex(You.Version, 2) & Pad(You.Name, 25) & FixedHex(Opponent.Picture, 2) & FixedHex(Opponent.Version, 2) & Pad(OpponentName, 25))
            Call AddToReplay(Command & OData)
            Worked = ThisBattle.SetVer(PNum, You.Version)
            Worked = ThisBattle.SetVer(ONum, Opponent.Version)
            ThisBattle.PlayerName(PNum) = YourName
            ThisBattle.PlayerName(ONum) = OpponentName
            If StoredRandTeam = "" Then
                Temp3 = ""
                For X = 1 To 6
                    Temp3 = Temp3 & PKMN2Str(PKMN(X))
                Next X
                Worked = ThisBattle.SetTeam(PNum, Temp3)
            Else
                Worked = ThisBattle.SetTeam(PNum, StoredRandTeam)
            End If
            Worked = ThisBattle.SetTeam(ONum, Data)
            Call ThisBattle.SetTrace(Dec(ChopString(Temp2, 1)))
            For X = 1 To 12
                TempEffect(X) = Asc(ChopString(Temp2, 1))
            Next X
            Call ThisBattle.SetItemEffect(TempEffect)
            Me.Caption = "Battling with " & OpponentName
            Call AddToReplay("MYTM:" & Pad(YourName, 20) & FixedHex(You.Picture, 2) & FixedHex(You.Version, 1) & ThisBattle.GetTeam(PNum))
            RBYMode = (ThisBattle.BattleMode = nbRBYBattle)
            If Resuming Then
                ThisBattle.StartBattle
            Else
                If ThisBattle.StadiumMode And Not Resuming Then
                    Stadium.Show vbModeless, MainContainer
                Else
                    Call StartTheMatch
                End If
            End If
            If Resuming Then
                ThisBattle.LogFile = SlashPath & "bl" & MD5(ServerAddress & You.Name) & ".log"
                ThisBattle.RestoreLog
            End If
        'Stadium Pokemon
        Case "SPKM:"
            Call AddToReplay("SPKM:" & ONum & Data)
            Worked = ThisBattle.SetSPoke(ONum, Data)
            If Not ThisBattle.NeedsStadiumSelect(PNum) Then
                Worked = ThisBattle.DoThreePKMN
                Call StartTheMatch
            End If
        'Start
        Case "BCMD:"
            Cancellable = False
            cmdUndo.Enabled = False
            Call AddToReplay(Command & Data)
            Worked = ThisBattle.ParseBattle(Data)
        Case "ENDTN"
            Call UpdateImages
            Call ResetTurn
            Call UpdateStats
            If ThisBattle.BattleOver Then
                P1 = 0
                P2 = 0
                For X = 1 To NumPoke
                    If BattlePKMN(1, X).HP <> 0 Then P1 = P1 + 1
                    If BattlePKMN(2, X).HP <> 0 Then P2 = P2 + 1
                Next X
                Call AddMessage("")
                Select Case ThisBattle.Winner
                    Case PNum
                        Call ThisBattle.AddMessage(35, nbPlayer, PNum)
                        Call ThisBattle.AddMessage(36, nbNumber, P1, nbNumber, P2)
                        If You.WinMess <> "" And Not ImJustWatching Then
                            Call AddMessage(YourName & ": " & You.WinMess, False, ":", vbRed, True, False)
                            Call SendData("CMSG:" & You.WinMess)
                        End If
                    Case ONum
                        Call ThisBattle.AddMessage(35, nbPlayer, ONum)
                        Call ThisBattle.AddMessage(36, nbNumber, P1, nbNumber, P2)
                        If You.LoseMess <> "" And Not ImJustWatching Then
                            Call AddMessage(YourName & ": " & You.LoseMess, False, ":", vbRed, True, False)
                            Call SendData("CMSG:" & You.LoseMess)
                        End If
                    Case 3
                        Call ThisBattle.AddMessage(31)
                        Call ThisBattle.AddMessage(36, nbNumber, P1, nbNumber, P2)
                        If You.LoseMess <> "" And Not ImJustWatching Then
                            Call AddMessage(YourName & ": " & You.LoseMess, False, ":", vbRed, True, False)
                            Call SendData("CMSG:" & You.LoseMess)
                        End If
                End Select
                Call DoEndBattle
            Else
                Attack.Enabled = ThisBattle.CanAttack(PNum)
                Switch.Enabled = ThisBattle.CanSwitch(PNum)
                If ThisBattle.NeedSwitch(PNum) Or ThisBattle.NeedSwitch(PNum + 2) Then
                    ControlTab.Tab = 1
                Else
                    ControlTab.Tab = 0
                End If
                If BattleCondition(PNum).BatonPassing Or BattleCondition(PNum + 2).BatonPassing Then
                    Switch.Caption = "&Baton Pass"
                    StatusBar1.Panels(3).Text = "Please select the Pokémon to Baton Pass to"
                Else
                    Switch.Caption = "&Switch"
                End If
                TabSwitch = False
            End If
            If Not ImJustWatching Then Call SendBattle 'Call SendData("RCVB:")
        Case "SYNC:" 'Sync Packet for Watchers
            SyncString = SyncString & Left(Data, 230)
            If Len(Data) <= 230 Then
                ThisBattle.VariableSync = SyncString
                SyncString = vbNullString
                If UseBG Then
                    Call MainContainer.DoPicture("bg" & CStr(ThisBattle.Terrain) & ".gif")
                    picTerrain.Picture = MainContainer.SwapSpace.Picture
                End If
                For X = 1 To 2
                    For Y = 1 To 6
                        BattlePKMN(X, Y) = GetClassPKMN(ThisBattle, CByte(X), CByte(Y))
                    Next Y
                Next X
                If ThisBattle.EnabledRule(nbExactHP) Then
                    For X = 1 To ThisBattle.ActNum
                        HPBar(X).Caption = nbExact
                    Next X
                End If
                If Resuming Then
                    StartTheMatch
                    Resuming = False
                    Call SendBattle
                Else
                    PokeCenter.Visible = False
                    Computer.Visible = False
                    Call BattleSync
                    Call UpdateImages
                    Call UpdateStats
                    Call ThisBattle.AddMessage(34, nbPlayer, 1, nbPlayer, 2)
                End If
                StatusBar1.Panels(3).Text = "Network Status: Battle sync complete"
            End If
'        Case "TLOG:"
'            SyncString = SyncString & Left(Data, 230)
'            If Len(Data) <= 230 Then ThisBattle.DecompressLog SyncString
        'Chat messages
        Case "CANCL"
            With ThisBattle
                For X = PNum To .ActNum Step 2
                    If .GetMoved(X) Or .GetSwitchTo(X) > 0 Then
                        .UnloadMove X
                        .UnloadSwitch X
                    End If
                Next X
            End With
            Call SendBattle
                
        Case "BACK:"
            If Len(Data) = 0 Then
                Call AddMessage("Your opponent has returned.", False, , vbRed, True)
                cmdLeave.Caption = "Forfeit"
            Else
                Call AddMessage(ThisBattle.PlayerName(Val(Data)) & " has returned.", False, , vbRed, True)
            End If
        Case "CMSG:"
            Call AddToReplay("CHAT:" & Data)
            If Left$(Data, 4) = "*** " Then
                Data = ApplyCSFilter(Data)
                Call AddMessage(Data, False, , &HC000C0)
            Else
                X = InStr(1, Data, ":")
                Mid(Data, X) = ApplyCSFilter(Mid$(Data, X))
                Call AddMessage(Data, False, ":", vbBlue, True, False)
            End If
        Case "WMSG:"
            Call AddToReplay("CHAT:" & Data)
            If Left$(Data, 4) = "*** " Then
                Data = ApplyCSFilter(Data)
                Call AddMessage(Data, False, , &HC000C0)
            Else
                X = InStr(1, Data, ":")
                Mid(Data, X) = ApplyCSFilter(Mid$(Data, X))
                Call AddMessage(Data, False, ":", vbDarkGreen, True, False)
            End If
        Case "WTCH:"
            Call AddMessage(Player(Dec(Data)).Name & " has started watching.", , , vbDarkGreen, True)
            Call ThisBattle.NewWatcher(Dec(Data))
            TempWatch = Split(ThisBattle.GetWatchers, ";")
            StatusBar1.Panels(1).Text = "Spectators: " & CStr(UBound(TempWatch) + 1)
            For X = 0 To UBound(TempWatch)
                TempWatch(X) = Player(Val(TempWatch(X))).Name
            Next X
            StatusBar1.Panels(1).ToolTipText = Join(TempWatch, ", ")
        Case "DONW:"
            If ChopString(Data, 1) = "1" Then
                Call ThisBattle.RemoveWatcher(Asc(ChopString(Data, 1)))
                Call AddMessage(Data & " has left.", , , vbDarkGreen, True)
            Else
                Call AddMessage(Player(Dec(Data)).Name & " has left.", , , vbDarkGreen, True)
                Call ThisBattle.RemoveWatcher(Dec(Data))
            End If
            TempWatch = Split(ThisBattle.GetWatchers, ";")
            X = UBound(TempWatch) + 1
            If X = 0 Then Temp2 = "" Else Temp2 = "Spectators: " & CStr(X)
            StatusBar1.Panels(1).Text = Temp2
            For X = 0 To UBound(TempWatch)
                TempWatch(X) = Player(Val(TempWatch(X))).Name
            Next X
            StatusBar1.Panels(1).ToolTipText = Join(TempWatch, ", ")
        'Disconnect
        Case "EXIT"
            FinalExit = False
            Unload Me
        'Tried to connect to Master Server
        Case "MSRV"
            FinalExit = False
            Answer = MsgBox("There is a server at this address, not an individual game." & vbCrLf & "You must connect using the 'Connect to a Server' option." & vbCrLf & "Would you like to open the proper connection now?", vbExclamation + vbYesNo, "Error")
            If Answer = vbYes Then
                RelayServer = True
                GameType = 2
                MasterServer.Show
                Unload Me
            End If
        Case "WAIT:"
            StatusBar1.Panels(3).Text = "Waiting for your opponent to pick Pokémon..."
        'Server Message Handling - only processed by client
        Case "SMSG:"
            Call AddMessage(Data)
        'Battle hacked; force a loss for the guilty player
        Case "HACK:"
            Call AddToReplay("HACK:" & Data)
            Y = Val(Data)
            Call AddMessage(vbNewLine & "The server has detected illegal activity on the part of " & IIf(Y = PNum, YourName, OpponentName) & ".  If this detection was erroneous, please report it to a NetBattle staff member at www.netbattle.net.")
            If Not ImJustWatching Then
                If Y = PNum Then
                    Call AddMessage(YourName & ": " & You.LoseMess, , ":", vbRed, True)
                    Call SendData("CMSG:" & You.LoseMess)
                Else
                    Call AddMessage(YourName & ": " & You.WinMess, , ":", vbRed, True)
                    Call SendData("CMSG:" & You.WinMess)
                End If
            End If
            For X = 1 To NumPoke
                BattlePKMN(Y, X).HP = 0
                BattlePKMN(Y, X).Condition = 8
            Next X
            If Y = PNum Then Call RefreshPokeList
            Call DoEndBattle
        Case "HURRY"
            Call AddToReplay("HURRY")
            If Val(Data) = PNum Then
                Call AddMessage(YourName & " has 30 seconds remaining before the win goes to " & OpponentName & " by default!", , , vbRed)
            Else
                Call AddMessage(OpponentName & " has 30 seconds remaining before the win goes to " & YourName & " by default!", , , vbRed)
            End If
        Case "TIME:"
            Call AddToReplay("TIME:" & Data)
            Y = Val(Data)
            If Y = PNum Then
                Call AddMessage(vbNewLine & YourName & " has been inactive for 5 minutes.  " & OpponentName & " wins by default!")
            Else
                Call AddMessage(vbNewLine & OpponentName & " has been inactive for 5 minutes.  " & YourName & " wins by default!")
            End If
            If Not ImJustWatching Then
                If Y = PNum Then
                    If Len(You.LoseMess) > 0 Then
                        Call AddMessage(YourName & ": " & You.LoseMess, , ":", vbRed, True)
                        Call SendData("CMSG:" & You.LoseMess)
                    End If
                Else
                    If Len(You.WinMess) > 0 Then
                        Call AddMessage(YourName & ": " & You.WinMess, , ":", vbRed, True)
                        Call SendData("CMSG:" & You.WinMess)
                    End If
                End If
            End If
            For X = 1 To NumPoke
                BattlePKMN(Y, X).HP = 0
                BattlePKMN(Y, X).Condition = 8
            Next X
            If Y = PNum Then RefreshPokeList
            DoEndBattle
        Case "DUMP:"
            Call AddMessage(OpponentName & " has saved a Battle Dump.", , , vbRed)
        Case "REQUN" 'REQuest UNrate
            If MsgBox("Your opponent has requested that the battle be unrated.  If a battle is unrated," + vbNewLine + "the outcome will not affect either players' battle record.  Do you accept?", vbYesNo + vbDefaultButton2 + vbQuestion, "Unrate Battle") = vbYes Then
                Call AddToQueue("UNACC")
                Call SendData("UNACC")
            Else
                Call SendData("REFUS")
            End If
        Case "UNACC" ' UNrate request ACCepted
            Call AddMessage("The battle is now unrated.", , , vbRed, True)
            mnuBattleItem(2).Enabled = False
            Battle.Caption = Battle.Caption & " - Unrated Battle"
            Call UnlockBattle
        Case "REQTI" 'REQuest TIe
            If MsgBox("Your opponent has proposed that the battle end in a Tie." + vbNewLine + "Do you accept?", vbYesNo + vbDefaultButton2 + vbQuestion, "Propose Tie") = vbYes Then
                Call AddToQueue("TIACC")
                Call SendData("TIACC")
            Else
                Call SendData("REFUS")
            End If
        Case "TIACC" 'TIe request ACCepted
            Call AddToReplay("TIACC")
            Call UnlockBattle
            Call DoTie
        Case "REFUS" 'Request REFUSed
            Call AddMessage("Your opponent has refused your request.", , , vbRed)
            Call UnlockBattle
        Case "BOVER"
            If ThisBattle.BattleOver Then
                MsgBox OpponentName & " has ended the battle.", vbInformation
            Else
                MsgBox OpponentName & " has forfeited.", vbInformation
            End If
            ReDim BattleQueue(0)
            Unload Me
            Exit Sub
        Case "IGWAT"
            X = Val(Data)
            If X = 1 Then
                ThisBattle.WatchIgnore1 = Not ThisBattle.WatchIgnore1
                Call AddMessage(Player(ThisBattle.Player1).Name & IIf(ThisBattle.WatchIgnore1, " is ", " has stopped ") & "ignoring spectator chat.", , , 38400, True)
            Else
                ThisBattle.WatchIgnore2 = Not ThisBattle.WatchIgnore2
                Call AddMessage(Player(ThisBattle.Player2).Name & IIf(ThisBattle.WatchIgnore2, " is ", " has stopped ") & "ignoring spectator chat.", , , 38400, True)
            End If
            If X = PNum Then
                mnuOptionsItem(4).Enabled = True
                mnuOptionsItem(4).Checked = Not mnuOptionsItem(4).Checked
            End If
            Call WatchChatLock
    End Select
    tmrBattleQueue.Enabled = Not tmrDelay.Enabled
End Sub
Sub WatchChatLock()
    If ThisBattle.WatchIgnore1 And ThisBattle.WatchIgnore2 And ImJustWatching Then
        txtWatchIgnore.Visible = True
        SendMsg.Enabled = False
        ChatBox.Text = ""
    Else
        txtWatchIgnore.Visible = False
        SendMsg.Enabled = True
    End If
End Sub
Public Function IsWatching(ByVal PNum As Integer)
    IsWatching = ThisBattle.IsWatching(PNum)
End Function
Sub UpdateImages()
    Dim TempVar As String
    Dim X As Integer
    Dim Y As Integer
    
    If UseDX Then Call UpdateImagesDX: Exit Sub
    
    Call BattleSync
    For X = 1 To ThisBattle.ActNum
        If X Mod 2 = PNum - 1 Then
            If BattleCondition(X).Substitute = 0 Then
                TempVar = ChooseImage(BattleCurrent(X), ThisBattle.GetVer(ONum), , ThisBattle.CastformImage(X))
            Else
                TempVar = "subst.gif"
            End If
            If picPKMNImage(X).Tag <> TempVar Then
                Call MainContainer.DoPicture(TempVar)
                picPKMNImage(X).Picture = MainContainer.SwapSpace.Picture
                picPKMNMask(X).Picture = Nothing
                Call CreateMask(picPKMNImage(X), picPKMNMask(X))
                picPKMNImage(X).Tag = TempVar
            End If
            If BattleCondition(X).Substitute = 0 And (InStr(1, TempVar, "rs") > 0 Or InStr(1, TempVar, "fl") > 0 Or InStr(1, TempVar, "e") > 0 Or Left(TempVar, 5) = "unown") Then
                PKMNImage(X).Y = OriginalTop(X) + (960 - picPKMNImage(X).Height) - (Screen.TwipsPerPixelY * BasePKMN(BattleCurrent(X).No).Offset)
                Shadow(X).Vis = (BasePKMN(BattleCurrent(X).No).Offset > 0)
            Else
                PKMNImage(X).Y = OriginalTop(X) + (960 - picPKMNImage(X).Height)
                Shadow(X).Vis = False
            End If
            PKMNImage(X).X = OriginalLeft(X) + ((960 - picPKMNImage(X).Width) / 2)
        Else
            If BattleCondition(X).Substitute = 0 Then
                TempVar = ChooseImage(BattleCurrent(X), ThisBattle.GetVer(PNum), True, ThisBattle.CastformImage(X))
            Else
                TempVar = "substb.gif"
            End If
            If picPKMNImage(X).Tag <> TempVar Then
                Call MainContainer.DoPicture(TempVar)
                picPKMNImage(X).Picture = MainContainer.SwapSpace.Picture
                picPKMNMask(X).Picture = Nothing
                Call CreateMask(picPKMNImage(X), picPKMNMask(X))
                picPKMNImage(X).Tag = TempVar
            End If
            PKMNImage(X).Y = (960 - picPKMNImage(X).Height) + OriginalTop(X)
            PKMNImage(X).X = ((960 - picPKMNImage(X).Width) / 2) + OriginalLeft(X)
        End If
        PKMNImage(X).Vis = (BattleCurrent(X).HP > 0 And BattleCondition(X).SemiInvul = 0 And Not BattleCondition(X).BatonPassing)
        If PKMNImage(X).Vis = False Then Shadow(X).Vis = False
    Next
    For X = 1 To 4
        PKMNImage(X).OrigVis = PKMNImage(X).Vis
        Shadow(X).OrigVis = Shadow(X).Vis
    Next X
    Call RepaintBattleArea
    
    For X = 1 To NumPoke
        Select Case BattlePKMN(ONum, X).Condition
            Case 1
                OpponentStat(X - 1).Picture = PKBall(0).Picture
            Case 8
                OpponentStat(X - 1).Picture = PKBall(2).Picture
            Case Else
                OpponentStat(X - 1).Picture = PKBall(1).Picture
        End Select
    Next X
    For X = 1 To NumPoke
        Select Case BattlePKMN(PNum, X).Condition
            Case 1
                OpponentStat(X + 5).Picture = PKBall(0).Picture
            Case 8
                OpponentStat(X + 5).Picture = PKBall(2).Picture
            Case Else
                OpponentStat(X + 5).Picture = PKBall(1).Picture
        End Select
    Next X
End Sub
Sub UpdateImagesDX()
'    Dim X As Long
'    Dim Y As Long
'    Dim s As Long
'    Dim TempVar As String
'    Call BattleSync
'    For X = 1 To ThisBattle.ActNum
'        With DX.Surface(X)
'            If X Mod 2 = PNum - 1 Then
'                If X < 3 Then s = 5 Else s = 6
'                If BattleCondition(X).Substitute = 0 Then
'                    TempVar = ChooseImage(BattleCurrent(X), ThisBattle.GetVer(ONum), , ThisBattle.CastformImage(X))
'                Else
'                    TempVar = "subst.gif"
'                End If
'                If .Tag <> TempVar Then
'                    Call MainContainer.DoPicture(TempVar)
'                    .UpdatePicture MainContainer.SwapSpace
'                    .Trans = -2
'                    .Tag = TempVar
'                End If
'                If BattleCondition(X).Substitute = 0 And (InStr(1, TempVar, "rs") > 0 Or InStr(1, TempVar, "fl") > 0 Or Left(TempVar, 5) = "unown") Then
'                    .Top = OriginalTop(X) + (64 - .Height) - BasePKMN(BattleCurrent(X).No).Offset
'                    DX.Surface(s).Visible = (BasePKMN(BattleCurrent(X).No).Offset > 0)
'                Else
'                    .Top = OriginalTop(X) + (64 - .Height)
'                    DX.Surface(s).Visible = False
'                End If
'                .Left = OriginalLeft(X) + ((64 - .Width) \ 2)
'            Else
'                s = 0
'                If BattleCondition(X).Substitute = 0 Then
'                    TempVar = ChooseImage(BattleCurrent(X), ThisBattle.GetVer(PNum), True, ThisBattle.CastformImage(X))
'                Else
'                    TempVar = "substb.gif"
'                End If
'                If .Tag <> TempVar Then
'                    Call MainContainer.DoPicture(TempVar)
'                    .UpdatePicture MainContainer.SwapSpace
'                    .Trans = -2
'                    .Tag = TempVar
'                End If
'                .Top = (64 - .Height) + OriginalTop(X)
'                .Left = ((64 - .Width) \ 2) + OriginalLeft(X)
'            End If
'            .Visible = (BattleCurrent(X).HP > 0 And BattleCondition(X).SemiInvul = 0 And Not BattleCondition(X).BatonPassing)
'            If s > 0 And Not .Visible Then DX.Surface(s).Visible = False
'            .RealVis = .Visible
'            If s > 0 Then DX.Surface(s).RealVis = DX.Surface(s).Visible
'        End With
'    Next
'
'    For X = 1 To NumPoke
'        Select Case BattlePKMN(ONum, X).Condition
'            Case 1
'                OpponentStat(X - 1).Picture = PKBall(0).Picture
'            Case 8
'                OpponentStat(X - 1).Picture = PKBall(2).Picture
'            Case Else
'                OpponentStat(X - 1).Picture = PKBall(1).Picture
'        End Select
'    Next X
'    For X = 1 To NumPoke
'        Select Case BattlePKMN(PNum, X).Condition
'            Case 1
'                OpponentStat(X + 5).Picture = PKBall(0).Picture
'            Case 8
'                OpponentStat(X + 5).Picture = PKBall(2).Picture
'            Case Else
'                OpponentStat(X + 5).Picture = PKBall(1).Picture
'        End Select
'    Next X
'
End Sub

Sub UpdateStats()
    Dim X As Byte
    Dim PPTotal As Integer
    Dim PKMNNo As Integer
    Dim TTString As String
    Dim TempStat As Integer
    Dim LBPika As Boolean
    Dim TCWak As Boolean
    Dim MPDitto As Boolean
    Dim ThisName As String
    
    If IsReplayWindow Then Exit Sub
    Call BattleSync
    HPAnimTimer.Enabled = False
    For X = 1 To ThisBattle.ActNum
        With BattleCurrent(X)
            ThisName = BattlePKMN(TeamNum(X), .TeamNumber).Name
            If ThisBattle.ActNum = 2 Then
                PokeText(X).Caption = IIf(X = PNum, YourName, OpponentName) & "'s " & .Nickname & vbNewLine & "Lv." & .Level & IIf(ThisBattle.BattleMode <> nbRBYBattle, " " & Gender(.Gender), "") & " " & ThisName
            Else
                If .Nickname = ThisName Then
                    PokeText(X).Caption = ThisName & " L" & .Level
                Else
                    PokeText(X).Caption = .Nickname & " (" & ThisName & ")" & " L" & .Level
                End If
                If ThisBattle.BattleMode <> nbRBYBattle Then
                    Select Case .Gender
                    Case 1: PokeText(X).Caption = PokeText(X).Caption & " M"
                    Case 2: PokeText(X).Caption = PokeText(X).Caption & " F"
                    End Select
                End If
            End If
            Select Case .Condition
                Case Is <= 1
                    PokeCond(X).Picture = Nothing
                    PokeCond(X).ToolTipText = ""
                Case Else
                    PokeCond(X).Picture = MainContainer.Conditions.ListImages(.Condition).Picture
                    PokeCond(X).ToolTipText = Condition(X)
            End Select
        End With
    Next
    
    Call RefreshMoveList
    If SelPoke > 0 Then
        If ThisBattle.StruggleOK(SelPoke) Then Attack.Caption = "S&truggle" Else Attack.Caption = "&Attack!"
        Attack.Enabled = ThisBattle.CanAttack(SelPoke)
        Switch.Enabled = ThisBattle.CanSwitch(SelPoke)
    End If
    Call RefreshPokeList
    If ThisBattle.RealWeather = 0 Then StatusBar1.Panels(1).Text = "Normal" Else StatusBar1.Panels(1).Text = Weather(ThisBattle.RealWeather)
    
    For X = 1 To ThisBattle.ActNum
        CurrentTT(X) = ThisBattle.GetTT(X, 0, (TeamNum(X) = ONum) Or ImJustWatching)
    Next X
    For X = 0 To 5
        StoredTT(X) = ThisBattle.GetTT(PNum, X + 1)
    Next X
    If UseDX Then
        TT.Tag = ""
    Else
        If Mid(TT.Tag, 1, 1) = "C" Then
            TT.TipText = CurrentTT(Val(Mid(TT.Tag, 2, 1)))
        ElseIf Mid(TT.Tag, 1, 1) = "S" Then
            TT.TipText = StoredTT(Val(Mid(TT.Tag, 2, 1)))
        End If
    End If
    
    HPAnimTimer.Enabled = True
End Sub


Private Sub SwitchTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Integer
    If TT.Tag = "S" & Index Then Exit Sub
    If UseDX Then
        'If DX.Animating Then Exit Sub
    End If
    TT.Destroy
    TT.Tag = "S" & Index
    TT.Title = BattlePKMN(PNum, Index + 1).Nickname
    TT.TipText = StoredTT(Index)
    TT.Style = TTStandard
    TT.Create Entry(Index).hWnd
End Sub

'Private Sub YourPokemon_DblClick()
'    DebugMode = Not DebugMode
'    mnuDebug.Visible = DebugMode
'End Sub

Sub SendData(ByVal SendMe As String)
    Dim Temp As Integer
    If ThisBattle.BattleOver Then
        Select Case Left$(SendMe, 5)
        Case "CMSG:", "WMSG:"
        Case Else: Exit Sub
        End Select
    End If
    If IsReplayWindow Then Exit Sub
    On Error Resume Next
    Call AddMessage(SendMe, True)
    Call NetworkLog(SendMe, True)
    If ImJustWatching Then
        Call MasterServer.SendData("RELAY" & Chr$(WatchID) & SendMe)
    Else
        Call MasterServer.SendData("RELAY" & Chr$(BattleID) & SendMe)
    End If
    Call WriteDebugLog("Battle.SendData: " & SendMe)
End Sub

Public Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    If DebugMessage And Not DebugMode Then Exit Sub
    Call RTB.AddMessage(Message, BreakChar, Color, Bold, Italic)
End Sub

Sub NetworkLog(ByVal NetData As String, Optional ByVal Sender As Boolean = False)
    Dim UseThis As Integer
    Dim X As Integer
    'To further prevent cheating with Dumps:
    If Left(NetData, 5) = "INFO:" Then NetData = "INFO: {Opponent Team Data}"
    
    UseThis = 1025
    For X = 1024 To 1 Step -1
        If NetworkStuff(X) = "" Then UseThis = X
    Next X
    If UseThis < 1024 Then
        If Sender Then
            NetworkStuff(UseThis) = "Sent: " & NetData
        Else
            NetworkStuff(UseThis) = "Rcd.: " & NetData
        End If
    End If
End Sub

Public Sub SaveLog(Optional FileName As String)
    Dim FileToUse As String
    Dim X As Integer
    Dim Y As Integer
    Dim FileNum As Integer
    Dim Temp As String
    Dim TempPKMN() As Pokemon
    
    If FileName = "" Then
        With MainContainer.FileBox
            .DialogTitle = "Save Log File"
            .Flags = cdlOFNOverwritePrompt
            .CancelError = True
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .DefaultExt = ".txt"
            .FileName = ""
            Temp = GetSetting("NetBattle", "Options", "InitDir", "")
            If Temp <> "" Then .InitDir = Temp
            On Error GoTo CancelledLog
            .ShowSave
            FileToUse = .FileName
            SaveSetting "NetBattle", "Options", "InitDir", Left$(FileToUse, InStrRev(FileToUse, "\"))
        End With
    Else
        FileToUse = FileName
    End If
    FileNum = FreeFile
    Open FileToUse For Output As #FileNum
    If ThisBattle.RandBat Then
        ReDim TempPKMN(1 To 6)
        For Y = 1 To 6
            TempPKMN(Y) = BattlePKMN(PNum, Y)
        Next Y
        Print #FileNum, YourName & "'s Team:"
        Print #FileNum, MakeTeamText(TempPKMN)
        Print #FileNum, ""
        For Y = 1 To 6
            TempPKMN(Y) = BattlePKMN(ONum, Y)
        Next Y
        Print #FileNum, OpponentName & "'s Team:"
        Print #FileNum, MakeTeamText(TempPKMN)
        Print #FileNum, ""
    End If
    Print #FileNum, "Battle Log:"
    Print #FileNum, Messages.Text
    Print #FileNum, ""
    Print #FileNum, "NetBattle v" & You.ProgVersion
    Print #FileNum, "Log saved " & Date & " at " & Time
    Close #FileNum
    MsgBox "Log saved to " & FileToUse, , "Save Complete"
CancelledLog:
End Sub

Sub SaveDump()
    Dim FileToUse As String
    Dim X As Integer
    Dim FileNum As Integer
    Dim Temp As String
    FileNum = FreeFile
    MainContainer.FileBox.DialogTitle = "Save Data Dump"
    MainContainer.FileBox.Flags = cdlOFNOverwritePrompt
    MainContainer.FileBox.CancelError = True
    MainContainer.FileBox.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    With MainContainer.FileBox
        .DialogTitle = "Save Data Dump"
        .Flags = cdlOFNOverwritePrompt
        .CancelError = True
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .DefaultExt = ".txt"
        .FileName = ""
        Temp = GetSetting("NetBattle", "Options", "InitDir", "")
        If Temp <> "" Then .InitDir = Temp
        On Error GoTo CancelledDump
        .ShowSave
        FileToUse = .FileName
        SaveSetting "NetBattle", "Options", "InitDir", Left$(FileToUse, InStrRev(FileToUse, "\"))
    End With
    Open FileToUse For Output As #FileNum
'    Print #FileNum, "Player Num.:" & PNum
'    Print #FileNum, "Player Name:" & YourName
'    Print #FileNum, "Player Info:" & You.Extra
'    Print #FileNum, "Player Win: " & You.WinMess
'    Print #FileNum, "Player Lose:" & You.LoseMess
'    Print #FileNum, "Player Ver.:" & You.ProgVersion
'    Print #FileNum, "Player Pic.:" & You.Picture
'    Print #FileNum, "Enemy Name: " & OpponentName
'    Print #FileNum, "Enemy Info: " & Opponent.Extra
'    Print #FileNum, "Enemy Win:  " & Opponent.WinMess
'    Print #FileNum, "Enemy Lose: " & Opponent.LoseMess
'    Print #FileNum, "Enemy Ver.: " & Opponent.ProgVersion
'    Print #FileNum, "Enemy Pic.: " & Opponent.Picture
'    For X = 1 To 6
'        Print #FileNum, "Player Pokemon #" & X
'        Print #FileNum, "No:        " & BattlePKMN(PNum, X).No
'        Print #FileNum, "Species:   " & BattlePKMN(PNum, X).Name
'        Print #FileNum, "Nickname:  " & BattlePKMN(PNum, X).Nickname
'        Print #FileNum, "Condition: " & Condition(PKMN(X).Condition)
'        Print #FileNum, "Count:     " & BattlePKMN(PNum, X).ConditionCount
'        Print #FileNum, "Item:      " & Item(PKMN(X).Item)
'        Print #FileNum, "Level:     " & BattlePKMN(PNum, X).Level
'        Print #FileNum, "HP:        " & BattlePKMN(PNum, X).HP
'        Print #FileNum, "Max HP:    " & BattlePKMN(PNum, X).MaxHP
'        Print #FileNum, "Base HP:   " & BasePKMN(PKMN(X).No).MaxHP
'        Print #FileNum, "HP DV      " & BattlePKMN(PNum, X).DV_HP
'        Print #FileNum, "Attack:    " & BattlePKMN(PNum, X).Attack
'        Print #FileNum, "Attack DV: " & BattlePKMN(PNum, X).DV_Atk
'        Print #FileNum, "Base ATK:  " & BasePKMN(PKMN(X).No).Attack
'        Print #FileNum, "Defense:   " & BattlePKMN(PNum, X).Defense
'        Print #FileNum, "Defense DV:" & BattlePKMN(PNum, X).DV_Def
'        Print #FileNum, "Base DEF:  " & BasePKMN(PKMN(X).No).Defense
'        Print #FileNum, "Speed:     " & BattlePKMN(PNum, X).Speed
'        Print #FileNum, "Speed DV:  " & BattlePKMN(PNum, X).DV_Spd
'        Print #FileNum, "Base SPD:  " & BasePKMN(PKMN(X).No).Speed
'        Print #FileNum, "Sp.Attack: " & BattlePKMN(PNum, X).SpecialAttack
'        Print #FileNum, "Sp.Defense:" & BattlePKMN(PNum, X).SpecialDefense
'        Print #FileNum, "Special DV:" & BattlePKMN(PNum, X).DV_SAtk
'        Print #FileNum, "Base SATK: " & BasePKMN(PKMN(X).No).SpecialAttack
'        Print #FileNum, "Base SDEF: " & BasePKMN(PKMN(X).No).SpecialDefense
'        Print #FileNum, "Move #1:   " & Moves(PKMN(X).Move(1)).Name
'        Print #FileNum, "Move #2:   " & Moves(PKMN(X).Move(2)).Name
'        Print #FileNum, "Move #3:   " & Moves(PKMN(X).Move(3)).Name
'        Print #FileNum, "Move #4:   " & Moves(PKMN(X).Move(4)).Name
'        Print #FileNum, "Move 1 PP: " & BattlePKMN(PNum, X).PP(1)
'        Print #FileNum, "Move 2 PP: " & BattlePKMN(PNum, X).PP(2)
'        Print #FileNum, "Move 3 PP: " & BattlePKMN(PNum, X).PP(3)
'        Print #FileNum, "Move 4 PP: " & BattlePKMN(PNum, X).PP(4)
'    Next
'    For X = 1 To 6
'        Print #FileNum, "Opponent Pokemon #" & X
'        Print #FileNum, "No:        " & BattlePKMN(ONum, X).No
'        Print #FileNum, "Species:   " & BattlePKMN(ONum, X).Name
'        Print #FileNum, "Nickname:  " & BattlePKMN(ONum, X).Nickname
'        Print #FileNum, "Condition: " & Condition(BattlePKMN(ONum, X).Condition)
'        Print #FileNum, "Count:     " & BattlePKMN(ONum, X).ConditionCount
'        Print #FileNum, "Item:      " & Item(BattlePKMN(ONum, X).Item)
'        Print #FileNum, "Level:     " & BattlePKMN(ONum, X).Level
'        Print #FileNum, "HP:        " & BattlePKMN(ONum, X).HP
'        Print #FileNum, "Max HP:    " & BattlePKMN(ONum, X).MaxHP
'        Print #FileNum, "Base HP:   " & BasePKMN(BattlePKMN(ONum, X).No).MaxHP
'        Print #FileNum, "HP DV      " & BattlePKMN(ONum, X).DV_HP
'        Print #FileNum, "Attack:    " & BattlePKMN(ONum, X).Attack
'        Print #FileNum, "Attack DV: " & BattlePKMN(ONum, X).DV_Atk
'        Print #FileNum, "Base ATK:  " & BasePKMN(BattlePKMN(ONum, X).No).Attack
'        Print #FileNum, "Defense:   " & BattlePKMN(ONum, X).Defense
'        Print #FileNum, "Defense DV:" & BattlePKMN(ONum, X).DV_Def
'        Print #FileNum, "Base DEF:  " & BasePKMN(BattlePKMN(ONum, X).No).Defense
'        Print #FileNum, "Speed:     " & BattlePKMN(ONum, X).Speed
'        Print #FileNum, "Speed DV:  " & BattlePKMN(ONum, X).DV_Spd
'        Print #FileNum, "Base SPD:  " & BasePKMN(BattlePKMN(ONum, X).No).Speed
'        Print #FileNum, "Sp.Attack: " & BattlePKMN(ONum, X).SpecialAttack
'        Print #FileNum, "Sp.Defense:" & BattlePKMN(ONum, X).SpecialDefense
'        Print #FileNum, "Special DV:" & BattlePKMN(ONum, X).DV_SAtk
'        Print #FileNum, "Base SATK: " & BasePKMN(BattlePKMN(ONum, X).No).SpecialAttack
'        Print #FileNum, "Base SDEF: " & BasePKMN(BattlePKMN(ONum, X).No).SpecialDefense
'        Print #FileNum, "Move #1:   " & Moves(BattlePKMN(ONum, X).Move(1)).Name
'        Print #FileNum, "Move #2:   " & Moves(BattlePKMN(ONum, X).Move(2)).Name
'        Print #FileNum, "Move #3:   " & Moves(BattlePKMN(ONum, X).Move(3)).Name
'        Print #FileNum, "Move #4:   " & Moves(BattlePKMN(ONum, X).Move(4)).Name
'        Print #FileNum, "Move 1 PP: " & BattlePKMN(ONum, X).PP(1)
'        Print #FileNum, "Move 2 PP: " & BattlePKMN(ONum, X).PP(2)
'        Print #FileNum, "Move 3 PP: " & BattlePKMN(ONum, X).PP(3)
'        Print #FileNum, "Move 4 PP: " & BattlePKMN(ONum, X).PP(4)
'    Next
    Print #FileNum, "BattleConditions:"
    Print #FileNum, "Player's battlecurrent:"
    Print #FileNum, "Accuracy Mod: " & BattleCondition(PNum).AccuracyChange
    Print #FileNum, "Evasion Mod:  " & BattleCondition(PNum).EvadeChange
    Print #FileNum, "Attack Mod:   " & BattleCondition(PNum).AttackChange
    Print #FileNum, "Defense Mod:  " & BattleCondition(PNum).DefenseChange
    Print #FileNum, "Speed Mod:    " & BattleCondition(PNum).SpeedChange
    Print #FileNum, "S.Attack Mod: " & BattleCondition(PNum).SAttackChange
    Print #FileNum, "S.Defense Mod:" & BattleCondition(PNum).SDefenseChange
    Print #FileNum, "Attract:      " & BattleCondition(PNum).Attract
    Print #FileNum, "Bide Turns:   " & BattleCondition(PNum).BideCount
    Print #FileNum, "Bide Damage:  " & BattleCondition(PNum).BideDamage
    Print #FileNum, "Charging:     " & BattleCondition(PNum).Charging
    Print #FileNum, "Confused:     " & BattleCondition(PNum).Confuse
    Print #FileNum, "Conf. Count:  " & BattleCondition(PNum).ConfuseCounter
    Print #FileNum, "Cursed:       " & BattleCondition(PNum).Curse
    Print #FileNum, "Defense Curl: " & BattleCondition(PNum).DefenseCurl
    Print #FileNum, "Destiny Bond: " & BattleCondition(PNum).DestinyBond
    Print #FileNum, "Disabled Move:" & BattleCondition(PNum).DisabledMove
    Print #FileNum, "Disable Count:" & BattleCondition(PNum).DisableCount
    Print #FileNum, "Encore:       " & BattleCondition(PNum).Encore
    Print #FileNum, "Encore Length:" & BattleCondition(PNum).EncoreDuration
    Print #FileNum, "Encore Move:  " & BattleCondition(PNum).EncoreMove
    Print #FileNum, "Foresight:    " & BattleCondition(PNum).Foresight
    Print #FileNum, "Fury Cutter:  " & BattleCondition(PNum).FuryCutter
    'Print #FileNum, "Last Damage:  " & BattleCondition(PNum).LastDamage
    'Print #FileNum, "Last S.Damage:" & BattleCondition(PNum).LastSDamage
    Print #FileNum, "Leech Seed:   " & BattleCondition(PNum).LeechSeed
    Print #FileNum, "Light Screen: " & BattleTC(PNum).LightScreenCount
    'Print #FileNum, "Locked:       " & BattleCondition(PNum).Locked
    Print #FileNum, "Lock-On:      " & BattleCondition(PNum).LockOn
    Print #FileNum, "Mimiced Move: " & BattleCondition(PNum).MimicedMove
    Print #FileNum, "Minimize:     " & BattleCondition(PNum).Minimize
    Print #FileNum, "Mist:         " & BattleCondition(PNum).Mist
    Print #FileNum, "Nightmare:    " & BattleCondition(PNum).Nightmare
    Print #FileNum, "Perish Song:  " & BattleCondition(PNum).PerishSong
    Print #FileNum, "Protect %:    " & BattleCondition(PNum).ProtectPercent
    Print #FileNum, "Rage Counter: " & BattleCondition(PNum).RageCounter
    Print #FileNum, "Recharging:   " & BattleCondition(PNum).Recharging
    Print #FileNum, "Reflect:      " & BattleTC(PNum).ReflectCount
    Print #FileNum, "Repeat Move:  " & BattleCondition(PNum).RepeatMove
    Print #FileNum, "Repeat Count: " & BattleCondition(PNum).RepeatCount
    Print #FileNum, "Rollout:      " & BattleCondition(PNum).Rollout
    Print #FileNum, "Safeguard:    " & BattleTC(PNum).SafeguardCount
    Print #FileNum, "Spikes:       " & BattleTC(PNum).Spikes
    Print #FileNum, "Substitute:   " & BattleCondition(PNum).Substitute
    Print #FileNum, "Toxic Count:  " & BattleCondition(PNum).ToxicCount
    Print #FileNum, "Opponent's battlecurrent:"
    Print #FileNum, "Accuracy Mod: " & BattleCondition(ONum).AccuracyChange
    Print #FileNum, "Evasion Mod:  " & BattleCondition(ONum).EvadeChange
    Print #FileNum, "Attack Mod:   " & BattleCondition(ONum).AttackChange
    Print #FileNum, "Defense Mod:  " & BattleCondition(ONum).DefenseChange
    Print #FileNum, "Speed Mod:    " & BattleCondition(ONum).SpeedChange
    Print #FileNum, "S.Attack Mod: " & BattleCondition(ONum).SAttackChange
    Print #FileNum, "S.Defense Mod:" & BattleCondition(ONum).SDefenseChange
    Print #FileNum, "Attract:      " & BattleCondition(ONum).Attract
    Print #FileNum, "Bide Turns:   " & BattleCondition(ONum).BideCount
    Print #FileNum, "Bide Damage:  " & BattleCondition(ONum).BideDamage
    Print #FileNum, "Charging:     " & BattleCondition(ONum).Charging
    Print #FileNum, "Confused:     " & BattleCondition(ONum).Confuse
    Print #FileNum, "Conf. Count:  " & BattleCondition(ONum).ConfuseCounter
    Print #FileNum, "Cursed:       " & BattleCondition(ONum).Curse
    Print #FileNum, "Defense Curl: " & BattleCondition(ONum).DefenseCurl
    Print #FileNum, "Destiny Bond: " & BattleCondition(ONum).DestinyBond
    Print #FileNum, "Disabled Move:" & BattleCondition(ONum).DisabledMove
    Print #FileNum, "Disable Count:" & BattleCondition(ONum).DisableCount
    Print #FileNum, "Encore:       " & BattleCondition(ONum).Encore
    Print #FileNum, "Encore Length:" & BattleCondition(ONum).EncoreDuration
    Print #FileNum, "Encore Move:  " & BattleCondition(ONum).EncoreMove
    Print #FileNum, "Foresight:    " & BattleCondition(ONum).Foresight
    Print #FileNum, "Fury Cutter:  " & BattleCondition(ONum).FuryCutter
    'Print #FileNum, "Last Damage:  " & BattleCondition(ONum).LastDamage
    'Print #FileNum, "Last S.Damage:" & BattleCondition(ONum).LastSDamage
    Print #FileNum, "Leech Seed:   " & BattleCondition(ONum).LeechSeed
    Print #FileNum, "Light Screen: " & BattleTC(ONum).LightScreenCount
    'Print #FileNum, "Locked:       " & BattleCondition(ONum).Locked
    Print #FileNum, "Lock-On:      " & BattleCondition(ONum).LockOn
    Print #FileNum, "Mimiced Move: " & BattleCondition(ONum).MimicedMove
    Print #FileNum, "Minimize:     " & BattleCondition(ONum).Minimize
    Print #FileNum, "Mist:         " & BattleCondition(ONum).Mist
    Print #FileNum, "Nightmare:    " & BattleCondition(ONum).Nightmare
    Print #FileNum, "Perish Song:  " & BattleCondition(ONum).PerishSong
    Print #FileNum, "Protect %:    " & BattleCondition(ONum).ProtectPercent
    Print #FileNum, "Rage Counter: " & BattleCondition(ONum).RageCounter
    Print #FileNum, "Recharging:   " & BattleCondition(ONum).Recharging
    Print #FileNum, "Reflect:      " & BattleTC(ONum).ReflectCount
    Print #FileNum, "Repeat Move:  " & BattleCondition(ONum).RepeatMove
    Print #FileNum, "Repeat Count: " & BattleCondition(ONum).RepeatCount
    Print #FileNum, "Rollout:      " & BattleCondition(ONum).Rollout
    Print #FileNum, "Safeguard:    " & BattleTC(ONum).SafeguardCount
    Print #FileNum, "Spikes:       " & BattleTC(ONum).Spikes
    Print #FileNum, "Substitute:   " & BattleCondition(ONum).Substitute
    Print #FileNum, "Toxic Count:  " & BattleCondition(ONum).ToxicCount
    Print #FileNum, "Player moves used:"
    For X = 1 To 10
        Print #FileNum, X & ") " & Moves(BattleCondition(PNum).MoveUsed(X)).Name
    Next
    Print #FileNum, "Enemy moves used:"
    For X = 1 To 10
        Print #FileNum, X & ") " & Moves(BattleCondition(ONum).MoveUsed(X)).Name
    Next
    Print #FileNum, "Recent network traffic:"
    For X = 1 To 1024
        If NetworkStuff(X) <> "" Then
            Print #FileNum, X & " - " & NetworkStuff(X)
        End If
    Next
    Print #FileNum, Messages.Text
    Close #FileNum
    Call SendData("DUMP:")
    Call AddMessage(YourName & " has saved a Battle Dump.", , , vbRed)
    MsgBox "Battle battleconditions saved to " & FileToUse & vbNewLine & "Please send to netbattle@tvsian.com or shiningmasamune@yahoo.com for debugging purposes.", , "Dump Complete"
CancelledDump:
End Sub

Sub BottomScroll()
    Messages.SelStart = Len(Messages.Text)
    Messages.SelLength = 0
    Call SendMessage(Messages.hWnd, &HB7, 0, 0)
End Sub


Function Rollover(ByVal Value As Integer) As Integer
    If StadiumMode Then
        Rollover = Value
        Exit Function
    End If
    If Value <= 1024 Then
        Rollover = Value
    ElseIf Value > 1024 And Value <= 2048 Then
        Rollover = Value - 1024
    Else
        Rollover = Value - 2048
    End If
End Function

'Sync the display variables to the battle variables
Sub BattleSync()
    Dim X As Byte
    Dim Y As Byte
    For X = 1 To ThisBattle.ActNum
        BattleCurrent(X) = GetClassPKMN(ThisBattle, X, 0)
        BattleCondition(X) = GetClassBC(ThisBattle, X)
        If BattleCurrent(X).No = 0 Then
            MsgBox "Load Error!", vbCritical, "Error!"
            If InVBMode Then Stop
        End If
    Next X
    For X = 1 To 2
        BattleTC(X) = GetClassTC(ThisBattle, X)
        For Y = 1 To 6
            BattlePKMN(X, Y) = GetClassPKMN(ThisBattle, X, Y)
        Next
    Next X
End Sub

Public Sub GetPKMN(ByVal Team As Byte, ByVal Num As Byte)
    Call Code.GetClassPKMN(ThisBattle, Team, Num)
End Sub

Public Sub SetStadium(ByVal PokeString As String)
    Dim GoAhead As Boolean
    Battle.Enabled = True
    Battle.SetFocus
    Call SendData("SPKM:" & PokeString)
    Call AddToReplay("SPKM:" & PNum & PokeString)
    GoAhead = ThisBattle.SetSPoke(PNum, PokeString)
    If Not ThisBattle.NeedsStadiumSelect(ONum) Then
        GoAhead = ThisBattle.DoThreePKMN
        Call StartTheMatch
    End If
End Sub

Public Sub InitTFile()
    Dim X As Long
    If IsReplayWindow Then Exit Sub
    ReplayNum = FreeFile
    X = Int(Rnd * 65536)
    TFileName = SlashPath & "TEMP" & FixedHex(X, 4) & ".tmp"
    Open TFileName For Binary Access Write As #ReplayNum
    Call AddMessage("Temp File: " & TFileName, True)
End Sub

Sub ChangeHPBar(ByVal Bar As Byte, ByVal Value As Integer, Optional ByVal Relative = False)
    If Relative Then Value = HPBar(Bar).Value + Value
    If Value > HPBar(Bar).Max Then Value = HPBar(Bar).Max
    HPBar(Bar).Value = Value
End Sub

Sub RefreshPokeList()
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    On Error Resume Next
    If ImJustWatching Then Exit Sub
    If BattlePKMN(PNum, SelectedPKMN).HP = 0 Then
        For X = 1 To 6
            If BattlePKMN(PNum, X).HP > 0 Then
                SelectedPKMN = X
                Exit For
            End If
        Next X
    End If
    For X = 0 To ThisBattle.NumPoke - 1
        With BattlePKMN(PNum, X + 1)
            PKName(X).Caption = .Nickname
            PKHP(X).Caption = .HP & "/" & .MaxHP
            Select Case .Gender
                Case 0
                    PKGender(X).Visible = False
                Case Else
                    PKGender(X).Picture = GenderImage(.Gender).Picture
                    PKGender(X).Visible = Not ThisBattle.BattleMode = nbRBYBattle
            End Select
            StatIcon(X) = MainContainer.Conditions.ListImages(.Condition).Picture
'            Temp = ""
'            For Y = 1 To 4
'                If .Move(Y) <> 0 Then
'                    Temp = Temp & Moves(.Move(Y)).Name & " (" & .PP(Y) & "/" & .MaxPP(Y) & ") • "
'                End If
'            Next Y
'            Temp = Temp & "[" & IIf(.Item = 0, "No Item", Item(.Item)) & "]"
            Entry(X).ToolTipText = Temp
            SwitchTile(X).ToolTipText = Temp
            If .Level < 100 Then
                PKLevel(X).Caption = "L" & .Level
            Else
                PKLevel(X).Caption = .Level
            End If
            If .HP = 0 Then
                PKName(X).ForeColor = RGB(200, 200, 200)
                PKLevel(X).ForeColor = RGB(200, 200, 200)
                PKHP(X).ForeColor = RGB(200, 200, 200)
                If Not IsReplayWindow Then Entry(X).Picture = BadImage.Picture
            ElseIf X = SelectedPKMN - 1 Then
                PKName(X).ForeColor = vbWhite
                PKLevel(X).ForeColor = vbWhite
                PKHP(X).ForeColor = vbWhite
                If Not IsReplayWindow Then Entry(X).Picture = SelImage.Picture
            Else
                PKName(X).ForeColor = vbBlack
                PKLevel(X).ForeColor = vbBlack
                PKHP(X).ForeColor = vbBlack
                If Not IsReplayWindow Then Entry(X).Picture = GoodImage.Picture
            End If
        End With
    Next
    If IsReplayWindow Then
        For X = 0 To 5
            If BattleCurrent(PNum).TeamNumber = X + 1 Then
                Entry(X).Picture = SelImage.Picture
            ElseIf ThisBattle.ActNum = 4 And BattleCurrent(PNum + 2).TeamNumber = X + 1 Then
                Entry(X).Picture = SelImage.Picture
            ElseIf BattlePKMN(PNum, X + 1).HP = 0 Then
                Entry(X).Picture = BadImage.Picture
            Else
                Entry(X).Picture = GoodImage.Picture
            End If
        Next
    End If
End Sub
Private Sub DoIcons()
    Dim X As Byte
    For X = 0 To ThisBattle.NumPoke - 1
        Call MainContainer.DoPicture(ChooseImage(BattlePKMN(PNum, X + 1), nbGFXSml))
        PokeIcon(X).Picture = MainContainer.SwapSpace.Picture
    Next X
End Sub

Sub RefreshMoveList()
    Dim X As Integer
    Dim Y As Integer
    Dim NoSel As Boolean
    Dim TempMove As Move
    
    If ImJustWatching Or SelPoke = 0 Then Exit Sub
    NoSel = ThisBattle.StruggleOK(SelPoke)
    If ThisBattle.BattleMode = nbRBYBattle Then
        If BattleCondition(SelPoke).RepeatMove > 0 Then
            For X = 1 To 4
                If BattleCurrent(SelPoke).Move(X) = BattleCondition(SelPoke).RepeatMove Then Exit For
            Next X
            If X < 5 Then SelectedMove(SelPoke) = X
        End If
        If BattleCondition(OtherTeam(SelPoke)).RepeatMove > 0 Then NoSel = True
    End If
    If ThisBattle.BattleMode = nbRBYBattle And (BattleCurrent(SelPoke).Condition = nbSlp Or BattleCurrent(SelPoke).Condition = nbFrz) Then NoSel = True
    If BattleCondition(SelPoke).EncoreMove <> 0 Then SelectedMove(SelPoke) = BattleCondition(SelPoke).EncoreMove
    If BattleCondition(SelPoke).BideCount <> 0 Then SelectedMove(SelPoke) = BattleCondition(SelPoke).LastMoveSlot
    For X = 0 To 3
        TempMove = ConvertMove(Moves(BattleCurrent(SelPoke).Move(X + 1)), ThisBattle.BattleMode)
        If TempMove.ID = 91 And BattleCurrent(SelPoke).No > 0 Then
            If ThisBattle.BattleMode = nbAdvBattle Then
                TempMove.Power = HiddenPowerStrengthAdv(BattleCurrent(SelPoke))
                TempMove.Type = HiddenPowerTypeAdv(BattleCurrent(SelPoke))
            Else
                With BattleCurrent(SelPoke)
                    TempMove.Power = HiddenPowerStrength(.DV_Atk, .DV_Def, .DV_Spd, .DV_SAtk)
                    TempMove.Type = HiddenPowerType(.DV_Atk, .DV_Def)
                End With
            End If
        End If
        If TempMove.SpecialEffect = 178 Then
            Select Case ThisBattle.CurrentWeather
            Case 1
                TempMove.Type = 3
            Case 2
                TempMove.Type = 2
            Case 3
                TempMove.Type = 13
            Case 4
                TempMove.Type = 6
            End Select
            If ThisBattle.CurrentWeather <> 0 Then TempMove.Power = TempMove.Power * 2
        End If
        With TempMove
            If .ID = 0 Then
                PPLabel(X).ForeColor = RGB(200, 200, 200)
                MoveNameLabel(X).ForeColor = RGB(200, 200, 200)
                MoveNameLabel(X).Caption = "-----"
                PPLabel(X).Caption = "--/--"
                MoveType(X).Picture = Nothing
                MoveSel(X).Picture = BadImage.Picture
            ElseIf X = SelectedMove(SelPoke) - 1 And Not NoSel Then
                PPLabel(X).ForeColor = vbWhite
                MoveNameLabel(X).ForeColor = vbWhite
                MoveSel(X).Picture = SelImage.Picture
                MoveDesc.Text = ""
                If .Power > 0 Then MoveDesc.Text = MoveDesc.Text & "Power: " & .Power & vbCrLf
                If .Accuracy > 0 Then MoveDesc.Text = MoveDesc.Text & "Accuracy: " & .Accuracy & vbCrLf
                MoveDesc.Text = MoveDesc.Text & .Text
            Else
                PPLabel(X).ForeColor = vbBlack
                MoveNameLabel(X).ForeColor = vbBlack
                MoveSel(X).Picture = GoodImage.Picture
            End If
            If .ID > 0 Then
                MoveNameLabel(X).Caption = .Name
                MoveType(X).Picture = MainContainer.Types.ListImages(.Type).Picture
                PPLabel(X).Caption = BattleCurrent(SelPoke).PP(X + 1) & "/" & BattleCurrent(SelPoke).MaxPP(X + 1)
            End If
        End With
    Next
End Sub
Private Sub StartTheMatch()
    Dim X As Integer
    PokeCenter.Visible = False
    Computer.Visible = False
    If Not Resuming Then ThisBattle.StartBattle
    Cancellable = True
    ControlTab.Enabled = True
    Call UpdateImages
    Call ResetTurn
    Call UpdateStats
    Call DoIcons
    For X = ThisBattle.NumPoke To 5
        Entry(X).Picture = BadImage.Picture
        PKGender(X).Visible = False
        PKName(X).Visible = False
        PKHP(X).Visible = False
        PKLevel(X).Visible = False
        PokeIcon(X).Visible = False
        StatIcon(X).Visible = False
        OpponentStat(X).Visible = False
        OpponentStat(X + 6).Visible = False
        OpponentStat(X).Top = OpponentStat(X).Top - 120
        OpponentStat(X + 6).Top = OpponentStat(X + 6).Top - 120
    Next X
    'Attack.Enabled = True
    'Switch.Enabled = True
    ControlTab.Tab = 0
    If MusicOption = 1 Then Call StartMusic
End Sub

Private Sub ResetTurn()
    Dim X As Byte
    SelPoke = 0
    If BattleCurrent(PNum).HP = 0 And BattleCurrent(PNum + 2).HP > 0 Then
        Call SetSelPoke(PNum + 2)
    Else
        Call SetSelPoke(PNum)
    End If
    For X = 1 To 4
        Call ThisBattle.UnloadMove(X)
        SelTarg(X) = 0
        SelMove(X) = 0
        SelSwitch(X) = 0
    Next X
End Sub

Private Sub SendBattle()
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Send As Boolean
    Dim Build As String
    Dim NS() As Boolean
    For X = PNum To ThisBattle.ActNum Step 2
        If Not ThisBattle.Ready(X) Then
            Cancellable = True
            Call SetSelPoke(X)
            Exit Sub
        End If
    Next X
    ReDim NS(1 To 4)
    Y = 0
    For X = 1 To ThisBattle.ActNum
        NS(X) = ThisBattle.NeedSwitch(X)
        If NS(X) Then Y = Y + 1
    Next X
    If Y = 1 Then Cancellable = False
    
    Send = True
    With ThisBattle
        If (NS(ONum) Or NS(ONum + 2)) And Not (NS(PNum) Or NS(PNum + 2)) Then Send = False
    End With
    'Debug.Print "Send Battle: " & Timer
    If Send Then
        For X = PNum To 4 Step 2
            If SelSwitch(X) > 0 Then
                Build = Build & "0" & Dec2Bin(SelSwitch(X), 5)
            Else
                If SelTarg(X) > X Then SelTarg(X) = SelTarg(X) - 1
                Build = Build & "1" & Dec2Bin(SelMove(X), 3) & Dec2Bin(SelTarg(X), 2)
            End If
        Next X
        Call SendData("MOVE:" & Bin2Chr(Build))
    End If
    For X = 1 To 4
        Call ThisBattle.UnloadMove(X)
        SelMove(X) = 0
        SelTarg(X) = 0
        SelSwitch(X) = 0
    Next X
    Call SetSelPoke(0)
    Attack.Enabled = False
    Switch.Enabled = False
    cmdUndo.Enabled = Cancellable
    StatusBar1.Panels(3).Text = "Network Status: Link Standby..."
End Sub

Private Sub AddToReplay(Text As String)
    On Error Resume Next
    If Not ImJustWatching And Not IsReplayWindow And Not ThisBattle.BattleOver Then
        Put #ReplayNum, , FormatPacket(Text, True)
    End If
End Sub

Private Sub SaveReplay(FileName As String)
    Dim X As Long
    Dim BArray() As Byte
    Dim C1 As Long
    Dim C2 As Long
    Dim C3 As Long
    Dim C4 As Long
    Dim tmp As String
    Dim FileNum As Integer
    Dim Worked As Boolean
    If IsReplayWindow Or ImJustWatching Then Exit Sub
    On Error GoTo SaveError
    If FileExists(FileName) Then Kill FileName
    FileNum = FreeFile
    Open TFileName For Binary Access Read As #FileNum
    ReDim BArray(LOF(FileNum) - 1)
    Get #FileNum, , BArray()
    Close #FileNum
    'Generate a few checksums...
    For X = 0 To UBound(BArray)
        C1 = C1 + BArray(X)
        If C1 > 10000 Then C1 = C1 Mod 89
        C2 = C2 + BArray(X) - 55
        If Abs(C2) > 10000 Then C2 = C2 Mod 75
        If X <> 0 Then C3 = C3 + (BArray(X - 1) Xor BArray(X)) - 10
        If Abs(C3) > 100000 Then C3 = C3 Mod 101
        C4 = C4 + BArray(X) - BArray(UBound(BArray) - X) + 10
        If Abs(C4) > 1000 Then C3 = C3 Mod 39
    Next X
    C1 = Abs(C1) Mod 256
    C2 = Abs(C2) Mod 256
    C3 = Abs(C3) Mod 256
    C4 = Abs(C4) Mod 256
    ReDim Preserve BArray(X + 3)
    BArray(X) = C1
    BArray(X + 1) = C2
    BArray(X + 2) = C3
    BArray(X + 3) = C4
    tmp = "header" & FixedHex(Int(Rnd * 65536), 4) & ".tmp"
    Open SlashPath & tmp For Output As #FileNum
    Worked = MainContainer.Compressor.CompressData(BArray())
    Write #FileNum, MainContainer.Compressor.OriginalSize
    Close #FileNum
    'Debug.Print MainContainer.Compressor.OriginalSize
    'Debug.Print UBound(BArray())
    ReDim HeaderBytes(FileLen(SlashPath & tmp) - 1) As Byte
    Open SlashPath & tmp For Binary Access Read As #FileNum
    Get #FileNum, , HeaderBytes()
    Close #FileNum
    Kill SlashPath & tmp
    Open FileName For Binary Access Write As #FileNum
    Put #FileNum, , HeaderBytes()
    Put #FileNum, , BArray
    Close #FileNum
    Exit Sub
SaveError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error saving replay."
    If InVBMode Then
        Stop
        Resume
    End If
End Sub
Public Function BattleOver() As Boolean
    BattleOver = ThisBattle.BattleOver
End Function

Private Sub YourPokemon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Integer
    If UseDX Then
        'If DX.Animating Then Exit Sub
    End If
    If BattleCurrent(PNum).No = 0 Or Display1 Then Exit Sub
    Display1 = True
    Display2 = False
    For Z = 0 To 5
        Display(Z) = False
    Next Z
    TT.Title = BattleCurrent(PNum).Nickname
    TT.TipText = StoredTT1
    'TT.Style = TTStandard
    TT.Create BattleArea.hWnd
End Sub

Private Sub RearrangeForm(ByVal BattleType As Byte, ByVal Watching As Boolean, Optional ByVal Replay As Boolean = False)
    Dim X As Long
    Dim Y As Long
    Dim Z As Byte
    'This part arranges the stuff inside the BattleArea PictureBox.
    'No sense putting it in twice - it's the same between battle and spectator modes.
    picStatus(PNum).Move 0, 3000
    picStatus(ONum).Move 840, 30
    
    ActNum = ThisBattle.ActNum
    
    If UseDX Then
'        With DX
'            .SetZOrder 5, 2
'            .SetZOrder 6, 3
'            If PNum = 1 Then
'                .SetZOrder 2, 7
'            Else
'                .SetZOrder 1, 7
'                .SetZOrder 3, 6
'            End If
'        End With
    End If
      
    Select Case BattleType
        '1v1
        Case 2
            If UseDX Then
'                With DX
'                    .Surface(PNum).Move 41, 48
'                    .Surface(ONum).Move 144, 8
'                    .Surface(5).Move 160, 68
'                    .Surface(3).Left = 300
'                    .Surface(4).Left = 300
'                End With
            Else
                PKMNImage(PNum).X = 480: PKMNImage(PNum).Y = 720
                PKMNImage(ONum).X = 2160: PKMNImage(ONum).Y = 120
                PKMNImage(3).X = 4000
                PKMNImage(4).X = 4000
                Shadow(ONum).X = 2400: Shadow(ONum).Y = 1020
            End If
            For X = 1 To 2
                HPBar(X).Move 60, 675, 2775, 375
                PokeText(X).Move 90, 90, 2595, 495
                PokeCond(X).Move 2580, 420
                SelPKMN(X).Visible = False
            Next X
            For X = 3 To 4
                PokeCond(X).Visible = False
                HPBar(X).Visible = False
                SelPKMN(X).Visible = False
                PokeText(X).Visible = False
            Next X
        '2v2
        Case 4
            If UseDX Then
'                With DX
'                    .Surface(PNum).Move 0, 48
'                    .Surface(ONum).Move 168, 8
'                    .Surface(PNum + 2).Move 56, 49
'                    .Surface(ONum + 2).Move 120, 0
'                    .Surface(5).Move 186, 68
'                    .Surface(6).Move 138, 60
'                End With
            Else
                PKMNImage(PNum).X = 0: PKMNImage(PNum).Y = 700
                PKMNImage(ONum).X = 2525: PKMNImage(ONum).Y = 120
                PKMNImage(PNum + 2).X = 840: PKMNImage(PNum + 2).Y = 820
                PKMNImage(ONum + 2).X = 1800: PKMNImage(ONum + 2).Y = 0
                Shadow(ONum).X = 2790: Shadow(ONum).Y = 1020
                Shadow(ONum + 2).X = 2070: Shadow(ONum + 2).Y = 905
            End If
            For Z = 1 To 4
                If (Z = 1 Or Z = 4) Xor (PNum = 1) Then Y = 600 Else Y = 0
                If Y = 0 Then X = 0 Else X = 120
                SelPKMN(Z).Move 60 + X, 30 + Y
                PokeCond(Z).Move 2475 + X, 30 + Y
                PokeText(Z).Move 105 + X, 60 + Y, 2595, 255
                HPBar(Z).Move 75 + X, 270 + Y, 2655, 270
            Next Z
    End Select
    
    For X = 1 To 4
        If UseDX Then
'            With DX.Surface(X)
'                OriginalTop(X) = .Top
'                OriginalLeft(X) = .Left
'                .sngX = .Left
'                .sngY = .Top
'                .Visible = False
'            End With
        Else
            OriginalTop(X) = PKMNImage(X).Y
            OriginalLeft(X) = PKMNImage(X).X
        End If
        If (X Mod 2 = PNum - 1) Or Watching Then
            HPBar(X).Caption = nbPercent
        Else
            HPBar(X).Caption = nbExact
        End If
    Next X
    
    'Hide controls if watching
    If Watching Then
        ControlTab.Visible = False
        Me.Height = 5370
        mnuBattleItem(0).Visible = False
        mnuBattleItem(1).Visible = False
        mnuBattleItem(2).Visible = False
        mnuBattleItem(3).Visible = False
        mnuBattleItem(4).Visible = False
        Exit Sub
    Else
        Me.Height = 7105
        mnuBattleItem(0).Visible = True
        mnuBattleItem(1).Visible = True
        mnuBattleItem(2).Visible = True
        mnuBattleItem(3).Visible = True
        mnuBattleItem(4).Visible = True
    End If
    
    'Show Replay controls if Replaying
    If Replay Then
        ReplayControls.Visible = True
        'KillConn.Visible = False
        SendMsg.Visible = False
        cmdLeave.Visible = False
        ChatBox.Left = Battle.Width + 120
        Messages.Height = 3015
        'PlayerFrame.Visible = True
        mnuBattle.Visible = False
        mnuFile.Visible = False
        mnuReplayFile.Visible = True
        mnuOptions.Visible = True
        'PokeTiles.Container = Battle
        ControlTab.Enabled = False
        Switch.Visible = False
    Else
        ReplayControls.Visible = False
        'KillConn.Visible = True
        SendMsg.Visible = True
        cmdLeave.Visible = True
        ChatBox.Left = 3760
        Messages.Height = 3375
        'PlayerFrame.Visible = False
        mnuBattle.Visible = False
        mnuFile.Visible = True
        mnuReplayFile.Visible = False
        mnuOptions.Visible = True
        'PokeTiles.Container = ControlTab
        ControlTab.Enabled = True
        Switch.Visible = True
    End If
    
    If Battling Then
        If OldInterface Then
            OldAttackFrame.Visible = True
            OldSwitchFrame.Visible = True
            ControlTab.Visible = False
        Else
            OldAttackFrame.Visible = False
            OldSwitchFrame.Visible = False
            ControlTab.Visible = True
        End If
    End If
End Sub

Sub ActivateTargetMode(ByVal UsingPKMN As Byte)
    Dim X As Integer
    
    KeyMode = 3
    For X = 1 To 4
        If X = UsingPKMN Or BattleCurrent(X).HP = 0 Then ValidTarget(X) = False Else ValidTarget(X) = True
    Next
    Select Case UsingPKMN
        Case 1, 3
            If ValidTarget(2) Then SelectedTarget = 2 Else SelectedTarget = 4
        Case 2, 4
            If ValidTarget(1) Then SelectedTarget = 1 Else SelectedTarget = 3
    End Select
    If Not OldInterface Then
        cmdCancel.Visible = True
        cmdTarget(0).Tag = ONum + 2
        cmdTarget(1).Tag = ONum
        cmdTarget(2).Tag = ThisBattle.AllyNum(UsingPKMN)
        For X = 0 To 2
            With cmdTarget(X)
                .Caption = "&" & CStr(X + 1) & ": " & IIf(X = 2, "Ally ", "Foe ") & BattleCurrent(.Tag).Name
                .Enabled = ValidTarget(.Tag)
                .Visible = True
            End With
        Next X
        MoveDesc.Visible = False
    Else
        'You'll have to work out something here yourself...
        cmdOldCancel.Visible = True
    End If
    ControlTab.Enabled = False
    DoFlash(SelectedTarget) = True
    StatusBar1.Panels(3).Text = "Please click a target to continue."
    If OldInterface Then
        For X = 1 To 4
            OldAttack(X - 1).Caption = BattleCurrent(X).Nickname
            OldPP(X - 1).Caption = ""
            If ValidTarget(X) Then OldAttack(X - 1).Enabled = True Else OldAttack(X - 1).Enabled = False
        Next
        OldAttack(SelectedTarget - 1).Value = True
    End If
End Sub

Sub DeactivateTargetMode()
    Dim X As Byte
    KeyMode = 0
    Call ClearFlash
    For X = 1 To 4
        ValidTarget(X) = False
    Next X
    For X = 0 To 2
        cmdTarget(X).Visible = False
    Next X
    SelectedTarget = 0
    ControlTab.Enabled = True
    MoveDesc.Visible = True
    cmdCancel.Visible = False
    cmdOldCancel.Visible = False
    Call RefreshOldAttacks
End Sub

Sub Targeted(ByVal TargetPKMN As Byte)
    SelTarg(SelPoke) = TargetPKMN
    Call ThisBattle.LoadMove(SelPoke, SelMove(SelPoke), SelTarg(SelPoke), True)
    Call DeactivateTargetMode
    Call SendBattle
End Sub

Sub SetSelPoke(ByVal NewPoke As Byte)
    Dim X As Integer
    Dim B As Boolean
    If ImJustWatching Then NewPoke = 0
    If SelPoke = NewPoke Then Exit Sub
    If SelPoke > 0 Then
        If UseDX Then
            'DX.Surface(SelPoke).Top = (64 - DX.Surface(SelPoke).Height) + OriginalTop(SelPoke)
        Else
            X = (960 - picPKMNImage(SelPoke).Height) + OriginalTop(SelPoke)
            If PKMNImage(SelPoke).Y <> X Then
                PKMNImage(SelPoke).Y = X
                Call RepaintBattleArea
            End If
        End If
    End If
    SelPoke = NewPoke
    If SelPoke > 0 Then
        'If UseDX Then B = DX.Surface(SelPoke).Visible Else
        B = PKMNImage(SelPoke).Vis
        If B Then
            ControlTab.Tab = 0
            StatusBar1.Panels(3).Text = "What should " & BattleCurrent(NewPoke).Nickname & " do?"
        Else
            ControlTab.Tab = 1
            StatusBar1.Panels(3).Text = "Choose a Pokémon to send out."
        End If
        Call UpdateStats
        Call RefreshMoveList
    Else
        If Not ImJustWatching Then StatusBar1.Panels(3).Text = "Network Status: Link Standby..."
    End If
End Sub

Sub RepaintBattleArea()
    Dim Z As Integer
    If UseDX Then Exit Sub
    picBuild.Picture = Nothing
    'First the terrain
    If UseBG Then picBuild.PaintPicture picTerrain.Picture, 0, 0
    'Second any visible shadows
    For Z = 1 To 4
        If Shadow(Z).Vis Then
            Call PaintPictureTrans(picBuild, picShadow, picShadowMask, Shadow(Z).X, Shadow(Z).Y)
        End If
    Next Z
    'Third the actual pokes.
    If PNum = 1 Then
        If PKMNImage(1).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(1), picPKMNMask(1), PKMNImage(1).X, PKMNImage(1).Y)
        If PKMNImage(3).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(3), picPKMNMask(3), PKMNImage(3).X, PKMNImage(3).Y)
        If PKMNImage(4).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(4), picPKMNMask(4), PKMNImage(4).X, PKMNImage(4).Y)
        If PKMNImage(2).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(2), picPKMNMask(2), PKMNImage(2).X, PKMNImage(2).Y)
    Else
        If PKMNImage(2).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(2), picPKMNMask(2), PKMNImage(2).X, PKMNImage(2).Y)
        If PKMNImage(4).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(4), picPKMNMask(4), PKMNImage(4).X, PKMNImage(4).Y)
        If PKMNImage(3).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(3), picPKMNMask(3), PKMNImage(3).X, PKMNImage(3).Y)
        If PKMNImage(1).Vis Then Call PaintPictureTrans(picBuild, picPKMNImage(1), picPKMNMask(1), PKMNImage(1).X, PKMNImage(1).Y)
    End If
    picBuild.Picture = picBuild.Image
    
    'Lastly, copy the complete image to the BattleArea.
    BattleArea.PaintPicture picBuild.Picture, 0, 0
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
    'Debug.Print Final
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
        Case "CHAT:", "BCMD:", "INFO:", "MYTM:", "SPKM:", "HACK:", "TIME:", "TIACC:", "VNUM:", "RAND:", "XINF:"
        Case Else
            MsgBox "This replay file is invalid.", vbCritical, "Invalid Replay"
            Exit Sub
        End Select
        Y = Y + 1
    Loop
    Pos = 1
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

Sub RefreshOldAttacks()
    Dim X As Integer
    If ImJustWatching Or SelPoke = 0 Or Not OldInterface Then Exit Sub
    With BattleCurrent(SelPoke)
        For X = 1 To 4
            If .Move(X) = 0 Then
                OldAttack(X - 1).Caption = "---"
                OldPP(X - 1).Caption = "--/--"
                OldAttack(X - 1).Enabled = False
            Else
                OldAttack(X - 1).Caption = Moves(.Move(X)).Name
                OldPP(X - 1).Caption = .PP(X) & "/" & .MaxPP(X)
                If .PP(X) = 0 Then OldAttack(X - 1).Enabled = False Else OldAttack(X - 1).Enabled = True
            End If
        Next
    End With
End Sub

Sub StartMusic()
    If MusicOption = 0 Then Exit Sub
    Call StopMusic
    
    If ThisBattle.BattleOver Then
        If ThisBattle.Winner = PNum Then
            Call PlayMusic(nbMusicVictory, True)
        Else
            Call PlayMusic(nbMusicLost, True)
        End If
    Else
        Select Case ThisBattle.BattleMode
            Case nbRBYBattle
                Call PlayMusic(nbMusicRBY, True)
            Case nbGSCBattle
                Call PlayMusic(nbMusicGSC, True)
            Case nbAdvBattle
                Call PlayMusic(nbMusicRuSa, True)
        End Select
    End If
End Sub

Sub DoReplayCommand()
    Dim X As Byte
    Dim Worked As Boolean
    Dim Command As String
    Dim Temp As String
    Dim Answer As Integer
    Dim Data As String
    Dim Temp2 As String
    Dim TempEffect(1 To 12) As Byte
    Dim ThrowAway As Variant
    
    Data = ReplayCommand(Pos)
    Command = ChopString(Data, 5)
    Call WriteDebugLog("Replay command processing: " & Command & Data)
    With ThisBattle
        Select Case Command
            Case "VNUM:"
                ReplayVersion = Data
                'Call AddMessage("Recorded in NetBattle v" & Data, , , vbRed, True)
            Case "XINF:"
                RePlayerTemp(1).Picture = Dec(ChopString(Data, 2))
                RePlayerTemp(1).Version = Dec(ChopString(Data, 2))
                RePlayerTemp(1).Name = Trim(ChopString(Data, 25))
                RePlayerTemp(2).Picture = Dec(ChopString(Data, 2))
                RePlayerTemp(2).Version = Dec(ChopString(Data, 2))
                RePlayerTemp(2).Name = Trim(ChopString(Data, 25))
            Case "INFO:"
                If ReplayVersion <> You.ProgVersion Then
                    Answer = MsgBox("This replay was recorded in a different version of NetBattle (v" & ReplayVersion & ", you are using v" & You.ProgVersion & ")." & vbCrLf & "This replay will not run properly.  Do you want to try anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Version Error")
                    If Answer = vbNo Then Call cmdStop_Click: Exit Sub
                End If
                .BattleMode = Dec(ChopString(Data, 1))
                .ActNum = Dec(ChopString(Data, 1))
                .Terrain = Dec(ChopString(Data, 1))
                If UseBG Then
                    Call MainContainer.DoPicture("bg" & CStr(.Terrain) & ".gif")
                    picTerrain.Picture = MainContainer.SwapSpace.Picture
                End If
                .Rules = ChopString(Data, 8)
                Temp2 = ChopString(Data, 13)
                PNum = Val(ChopString(Data, 1))
                ONum = OtherTeam(PNum)
                X = Asc(ChopString(Data, 1))
                Select Case PNum
                    Case 1
                        RePlayer(1).Picture = RePlayerTemp(1).Picture
                        RePlayer(1).Version = RePlayerTemp(1).Version
                        RePlayer(1).Name = RePlayerTemp(1).Name
                        RePlayer(2).Picture = RePlayerTemp(2).Picture
                        RePlayer(2).Version = RePlayerTemp(2).Version
                        RePlayer(2).Name = RePlayerTemp(2).Name
                    Case 2
                        RePlayer(1).Picture = RePlayerTemp(2).Picture
                        RePlayer(1).Version = RePlayerTemp(2).Version
                        RePlayer(1).Name = RePlayerTemp(2).Name
                        RePlayer(2).Picture = RePlayerTemp(1).Picture
                        RePlayer(2).Version = RePlayerTemp(1).Version
                        RePlayer(2).Name = RePlayerTemp(1).Name
                End Select
                Worked = ThisBattle.SetTeam(ONum, Data)
                Call ThisBattle.SetTrace(Dec(ChopString(Temp2, 1)))
                For X = 1 To 12
                    TempEffect(X) = Asc(ChopString(Temp2, 1))
                Next X
                Call ThisBattle.SetItemEffect(TempEffect)
            Case "MYTM:"
                RePlayer(PNum).Name = Trim(ChopString(Data, 20))
                ThrowAway = Dec(ChopString(Data, 2))
                RePlayer(PNum).Version = Dec(ChopString(Data, 1))
                Worked = .SetTeam(PNum, Data)
                Worked = ThisBattle.SetVer(PNum, RePlayer(PNum).Version)
                Worked = ThisBattle.SetVer(ONum, RePlayer(ONum).Version)
                ThisBattle.PlayerName(PNum) = RePlayer(PNum).Name
                ThisBattle.PlayerName(ONum) = RePlayer(ONum).Name
                If Not .StadiumMode Then
                    .StartBattle
                    Call BattleSync
                    Call ReplayPKMNTiles
                    Call ReplayRefresh
                    PokeCenter.Visible = False
                    Computer.Visible = False
                    Call StartMusic
                End If
            Case "SPKM:"
                X = Val(ChopString(Data, 1))
                Worked = .SetSPoke(X, Data)
                If Not .NeedsStadiumSelect(OtherTeam(X)) Then
                    Worked = .DoThreePKMN
                    .StartBattle
                    Call BattleSync
                    Call ReplayPKMNTiles
                    Call ReplayRefresh
                    PokeCenter.Visible = False
                    Computer.Visible = False
                    Call StartMusic
                End If
            Case "CHAT:"
                Temp = Left(Data, InStr(1, Data, ":") - 1)
                If Temp = RePlayer(PNum).Name Then
                    Call AddMessage(Data, , ":", vbRed, True)
                ElseIf Temp = RePlayer(ONum).Name Then
                    Call AddMessage(Data, , ":", vbBlue, True)
                Else
                    Call AddMessage(Data, , ":", vbDarkGreen, True)
                End If
            Case "BCMD:"
                cmdNext.Enabled = False
                PBTimer.Enabled = False
                Worked = .ParseBattle(Data)
                Call ReplayRefresh
                cmdNext.Enabled = True
                PBTimer.Enabled = True
                If ThisBattle.BattleOver Then Call StartMusic
            Case "HACK:"
                Call ThisBattle.ForceLoss(Val(Data))
                BattleCurrent(Val(Data)).HP = 0
                BattleCurrent(Val(Data)).Condition = 8
                For X = 1 To 6
                    BattlePKMN(Val(Data), X).HP = 0
                    BattlePKMN(Val(Data), X).Condition = 8
                Next X
                Call ReplayRefresh
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
            Case Else
                Call AddMessage("Unhandled Replay command - see Debug log")
                Call WriteDebugLog("^^^ Unhandled Command ^^^")
        End Select
    End With
End Sub

Private Sub ReplayRefresh()
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Percent As Integer
    
    Call BattleSync
    Call RefreshPokeList
    Call UpdateImages
    HPAnimTimer.Enabled = False
    For X = 1 To ThisBattle.ActNum
        With BattleCurrent(X)
            If ThisBattle.ActNum = 2 Then
                PokeText(X).Caption = RePlayer(X).Name & "'s " & .Nickname & vbNewLine & "Lv." & .Level & " " & Gender(.Gender) & " " & .Name
            Else
                If .Nickname = .Name Then
                    PokeText(X).Caption = .Name & " L" & .Level
                Else
                    PokeText(X).Caption = .Nickname & " (" & .Name & ")" & " L" & .Level
                End If
                Select Case .Gender
                Case 1: PokeText(X).Caption = PokeText(X).Caption & " M"
                Case 2: PokeText(X).Caption = PokeText(X).Caption & " F"
                End Select
            End If
            Select Case .Condition
                Case Is <= 1
                    PokeCond(X).Picture = Nothing
                    PokeCond(X).ToolTipText = ""
                Case Else
                    PokeCond(X).Picture = MainContainer.Conditions.ListImages(.Condition).Picture
                    PokeCond(X).ToolTipText = Condition(X)
            End Select
        End With
    Next
    HPAnimTimer.Enabled = True
End Sub

Sub ReplayPKMNTiles()
    Dim X As Byte
    Dim Y As Byte
    
    'Okay, it actually does player data, too.
    'I just didn't feel like renaming the sub after I added in the functions.
    TIcon(PNum - 1).Picture = MainContainer.Trainers.ListImages(RePlayer(PNum).Picture).Picture
    TIcon(ONum - 1).Picture = MainContainer.Trainers.ListImages(RePlayer(ONum).Picture).Picture
    Call RearrangeForm(ThisBattle.ActNum, False, True)
    'Sticking this in here, just because it's easier, and I'm lazy.
    For X = 1 To ThisBattle.ActNum
        HPBar(X).Caption = nbExact
    Next
    'Okay, now let's handle the icons.
    If ThisBattle.StadiumMode Then
        Y = 2
        Entry(3).Visible = False
        Entry(4).Visible = False
        Entry(5).Visible = False
    Else
        Y = 5
        Entry(3).Visible = True
        Entry(4).Visible = True
        Entry(5).Visible = True
    End If
    For X = 0 To Y
        Call MainContainer.DoPicture(ChooseImage(BattlePKMN(PNum, X + 1), nbGFXSml))
        PokeIcon(X).Picture = MainContainer.SwapSpace.Picture
    Next
End Sub
Public Sub DoEndBattle()
    mnuFileItem(0).Enabled = True
    Attack.Enabled = False
    Switch.Enabled = False
    cmdLeave.Caption = "&Leave"
    mnuBattle.Enabled = False
    Call StartMusic
End Sub
Private Sub LockBattle()
    mnuFile.Enabled = False
    mnuBattle.Enabled = False
    ControlTab.Enabled = False
    StatusBar1.Panels(3).Text = "Waiting for response..."
End Sub
Private Sub UnlockBattle()
    mnuFile.Enabled = True
    mnuBattle.Enabled = True
    ControlTab.Enabled = True
    StatusBar1.Panels(3).Text = ""
End Sub

Private Sub tmrBattleQueue_Timer()
    Dim X As Long
    If UBound(BattleQueue) = 0 Then Exit Sub
    Call DoIncoming(BattleQueue(1))
    For X = 2 To UBound(BattleQueue)
        BattleQueue(X - 1) = BattleQueue(X)
    Next X
    ReDim Preserve BattleQueue(X - 2)
End Sub
Public Sub AddToQueue(NewItem As String)
    ReDim Preserve BattleQueue(UBound(BattleQueue) + 1)
    BattleQueue(UBound(BattleQueue)) = NewItem
End Sub

Private Sub tmrDelay_Timer()
    Dim B As Boolean
    B = True
    If UseDX Then
        'If DX.Animating Then B = False
    End If
    If B Then Call ThisBattle.ParseBattle("")
End Sub


Private Sub tmrDX_Timer()
    If UseDX Then BattleArea.Picture = BattleArea.Image
End Sub
Public Sub OppDiscon(Optional PNum As Long = 0)
    If PNum = 0 Then
        Call AddMessage("Your opponent has disconnected.", , , vbRed, True) ' from the server  You may leave, or wait for him or her to return."
        cmdLeave.Caption = "&Leave"
    Else
        Call AddMessage(ThisBattle.PlayerName(PNum) & " has disconnected.", , , vbRed, True)
    End If
End Sub
Private Sub mnuPokedexItem_Click(Index As Integer)
    MasterDex.Show
    MasterDex.SetMode Index
    MasterDex.SetVer ThisBattle.BattleMode
    If BattleCurrent(ONum).No > 0 Then MasterDex.SetPoke BattleCurrent(ONum).No
End Sub

