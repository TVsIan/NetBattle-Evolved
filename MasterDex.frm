VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form MasterDex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetBattle PokéDex  -  Ruby/Sapphire"
   ClientHeight    =   6735
   ClientLeft      =   2775
   ClientTop       =   2565
   ClientWidth     =   9270
   Icon            =   "MasterDex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSwap 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -10000
      MultiLine       =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Text            =   "MasterDex.frx":1272
      Top             =   6840
      Width           =   2655
   End
   Begin VB.PictureBox picSwap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   -10000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   6840
      Width           =   375
   End
   Begin MSComctlLib.ImageList EvoImages 
      Left            =   8640
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":12A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":2522
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":267C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":27D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":2ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":3464
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":39FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":3F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":40F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":424C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":43A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":4500
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":465A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":47B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":490E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MasterDex.frx":4EA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7860
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin TabDlg.SSTab MasterTab 
      Height          =   6555
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Pokémon"
      TabPicture(0)   =   "MasterDex.frx":5442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(12)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblInfo(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblInfo(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInfo(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInfo(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblInfo(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblInfo(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblInfo(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblInfo(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblInfo(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblInfo(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblInfo(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblInfo(9)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "NatDexNum"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DexNum"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "hscPokemon"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MovePool"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "EvoTree"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "picPokeImage"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPokemon"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "picStatBox"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "fraDexText(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "fraDexText(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "fraDexText(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "fraDexText(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "fraDexText(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Attacks, Items, && Abilities"
      TabPicture(1)   =   "MasterDex.frx":545E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMoves"
      Tab(1).Control(1)=   "TraitFrame"
      Tab(1).Control(2)=   "ItemFrame"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Type Matching Chart"
      TabPicture(2)   =   "MasterDex.frx":547A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Command1"
      Tab(2).Control(2)=   "Picture2"
      Tab(2).Control(3)=   "picTypeChart"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Damage Calculator"
      TabPicture(3)   =   "MasterDex.frx":5496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmResults"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "frmDefender"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "frmAttacker"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame frmResults 
         Caption         =   "Results"
         Height          =   3615
         Left            =   -70080
         TabIndex        =   122
         Top             =   420
         Width           =   4095
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   3360
            Top             =   3120
         End
         Begin VB.ComboBox cmbPoke 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   123
            Top             =   600
            Width           =   1815
         End
         Begin VB.PictureBox picDamage 
            BorderStyle     =   0  'None
            Height          =   2895
            Left            =   120
            ScaleHeight     =   2895
            ScaleWidth      =   3915
            TabIndex        =   124
            Top             =   600
            Visible         =   0   'False
            Width           =   3915
            Begin VB.PictureBox picEV 
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   0
               ScaleHeight     =   615
               ScaleWidth      =   3855
               TabIndex        =   136
               Top             =   480
               Width           =   3855
               Begin MSComctlLib.Slider Slider1 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   137
                  Top             =   360
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   450
                  _Version        =   393216
                  LargeChange     =   16
                  SmallChange     =   4
                  Max             =   255
                  SelStart        =   255
                  TickFrequency   =   16
                  Value           =   255
               End
               Begin MSComctlLib.Slider Slider2 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   138
                  Top             =   0
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   450
                  _Version        =   393216
                  LargeChange     =   16
                  SmallChange     =   4
                  Max             =   252
                  SelStart        =   252
                  TickFrequency   =   16
                  Value           =   252
               End
               Begin VB.Label lblDamageCalc 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP EV"
                  Height          =   255
                  Index           =   2
                  Left            =   -240
                  TabIndex        =   142
                  Top             =   0
                  Width           =   855
               End
               Begin VB.Label lblDamageCalc 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Def EV"
                  Height          =   255
                  Index           =   3
                  Left            =   -240
                  TabIndex        =   141
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label lblDamageCalc 
                  BackStyle       =   0  'Transparent
                  Caption         =   "252"
                  Height          =   255
                  Index           =   8
                  Left            =   3540
                  TabIndex        =   140
                  Top             =   0
                  Width           =   975
               End
               Begin VB.Label lblDamageCalc 
                  BackStyle       =   0  'Transparent
                  Caption         =   "252"
                  Height          =   255
                  Index           =   9
                  Left            =   3540
                  TabIndex        =   139
                  Top             =   360
                  Width           =   975
               End
            End
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   1080
               ScaleHeight     =   735
               ScaleWidth      =   1575
               TabIndex        =   128
               Top             =   1200
               Width           =   1575
               Begin VB.OptionButton optNature 
                  Caption         =   "Nature Reduction"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   131
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.OptionButton optNature 
                  Caption         =   "Nature Neutral"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   130
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton optNature 
                  Caption         =   "Nature Boost"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   129
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin VB.TextBox txtDamageCalc 
               Height          =   285
               Index           =   4
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   127
               Text            =   "100"
               Top             =   0
               Width           =   495
            End
            Begin VB.PictureBox Picture3 
               Height          =   240
               Left            =   3180
               ScaleHeight     =   180
               ScaleWidth      =   315
               TabIndex        =   125
               Top             =   1440
               Width           =   375
               Begin VB.Label lblBattleMod 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "0"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   -30
                  TabIndex        =   126
                  Top             =   0
                  Width           =   375
               End
            End
            Begin NetBattle.ColorProgress DemoBar 
               Height          =   375
               Left            =   0
               TabIndex        =   132
               Top             =   2520
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   661
            End
            Begin VB.HScrollBar HScroll1 
               Height          =   240
               Left            =   2940
               Max             =   6
               Min             =   -6
               TabIndex        =   135
               Top             =   1440
               Width           =   855
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Sp. Def"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   134
               Top             =   1440
               Width           =   1455
            End
            Begin VB.OptionButton optDef 
               Caption         =   "Defense"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   133
               Top             =   1200
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.Label lblDamageCalc 
               BackStyle       =   0  'Transparent
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   145
               Top             =   2280
               Width           =   2535
            End
            Begin VB.Label lblDamageCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "Level"
               Height          =   255
               Index           =   4
               Left            =   2880
               TabIndex        =   144
               Top             =   60
               Width           =   975
            End
            Begin VB.Image imgDefender 
               Height          =   960
               Left            =   2880
               Tag             =   "0"
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label lblDamageCalc 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Battle Mod"
               Height          =   255
               Index           =   14
               Left            =   2820
               TabIndex        =   143
               Top             =   1200
               Width           =   1095
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Damage:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Select a Pokémon from the list to see a simulation."
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   146
            Top             =   1800
            Width           =   3615
         End
      End
      Begin VB.Frame frmDefender 
         Caption         =   "Defender"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   112
         Top             =   2340
         Width           =   4695
         Begin VB.CheckBox chkDamageCalc 
            Caption         =   "Critical Hit"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   118
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "MasterDex.frx":54B2
            Left            =   120
            List            =   "MasterDex.frx":54C5
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton optReflect 
            Caption         =   "None"
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   116
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optReflect 
            Caption         =   "Single Battle"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   115
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optReflect 
            Caption         =   "Double Battle"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   114
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtDamageCalc 
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   4
            TabIndex        =   113
            Text            =   "200"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Type Effectiveness"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   121
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Light Screen/Reflect"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   120
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Defense/Sp. Def"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame frmAttacker 
         Caption         =   "Attacker"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   99
         Top             =   420
         Width           =   4695
         Begin VB.CheckBox chkDamageCalc 
            Caption         =   "Low-HP Trait Bonus"
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   107
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox cmbWeather 
            Height          =   315
            ItemData        =   "MasterDex.frx":550F
            Left            =   2280
            List            =   "MasterDex.frx":551C
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox chkDamageCalc 
            Caption         =   "Type Boosting Item Bonus"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   105
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtDamageCalc 
            Height          =   315
            Index           =   3
            Left            =   120
            MaxLength       =   3
            TabIndex        =   104
            Text            =   "0"
            Top             =   1440
            Width           =   495
         End
         Begin VB.ComboBox cmbMove 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   103
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtDamageCalc 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   3
            TabIndex        =   102
            Text            =   "100"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtDamageCalc 
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   4
            TabIndex        =   101
            Text            =   "200"
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkDamageCalc 
            Caption         =   "Same Type Attack Bonus"
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   100
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Weather Modifier"
            Height          =   255
            Index           =   13
            Left            =   2280
            TabIndex        =   111
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Power"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   110
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblDamageCalc 
            BackStyle       =   0  'Transparent
            Caption         =   "Attack/Sp. Atk"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   108
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraDexText 
         Caption         =   "Ruby Text"
         Height          =   1335
         Index           =   4
         Left            =   4620
         TabIndex        =   45
         Top             =   2340
         Width           =   4395
         Begin VB.Label lblDexText 
            BackStyle       =   0  'Transparent
            Height          =   975
            Index           =   4
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   4155
         End
      End
      Begin VB.Frame fraDexText 
         Caption         =   "Ruby Text"
         Height          =   1335
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   2340
         Width           =   4395
         Begin VB.Label lblDexText 
            BackStyle       =   0  'Transparent
            Height          =   975
            Index           =   3
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   4155
         End
      End
      Begin VB.Frame ItemFrame 
         Caption         =   "Items"
         Height          =   2655
         Left            =   -69720
         TabIndex        =   71
         Top             =   480
         Width           =   3735
         Begin VB.TextBox txtItem 
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1575
         End
         Begin MSComctlLib.ListView ItemList 
            Height          =   1935
            Left            =   120
            TabIndex        =   94
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2187
            EndProperty
         End
         Begin VB.Label lblItemDesc 
            Caption         =   "When damage is done to a Pokémon holding this item that would Faint it, there is a 10% chance of the Pokémon surviving with 1HP."
            Height          =   1575
            Left            =   1800
            TabIndex        =   81
            Top             =   240
            Width           =   1815
         End
         Begin VB.Image imgMoveCompat 
            Height          =   495
            Index           =   4
            Left            =   2880
            ToolTipText     =   "True RBY"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Image imgMoveCompat 
            Height          =   495
            Index           =   3
            Left            =   2040
            ToolTipText     =   "True RBY"
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Frame TraitFrame 
         Caption         =   "Traits"
         Height          =   2655
         Left            =   -69720
         TabIndex        =   69
         Top             =   3240
         Width           =   3735
         Begin MSComctlLib.ListView TraitList 
            Height          =   1935
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2187
            EndProperty
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Search for Pokémon that have this Trait..."
            Height          =   495
            Left            =   1800
            TabIndex        =   79
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtTrait 
            Height          =   285
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblTraitDesc 
            Caption         =   $"MasterDex.frx":5549
            Height          =   1815
            Left            =   1800
            TabIndex        =   78
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame fraMoves 
         Caption         =   "Attacks"
         Height          =   5415
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   5055
         Begin VB.CommandButton Command5 
            Caption         =   "Search for Pokémon that learn this Move..."
            Height          =   495
            Left            =   2520
            TabIndex        =   80
            Top             =   4800
            Width           =   2175
         End
         Begin VB.Frame Frame5 
            Caption         =   "Details"
            Height          =   1335
            Left            =   2280
            TabIndex        =   73
            Top             =   3240
            Width           =   2655
            Begin VB.Image imgMoveInfo 
               Height          =   255
               Index           =   3
               Left            =   240
               Top             =   960
               Width           =   255
            End
            Begin VB.Image imgMoveInfo 
               Height          =   255
               Index           =   2
               Left            =   240
               Top             =   720
               Width           =   255
            End
            Begin VB.Image imgMoveInfo 
               Height          =   255
               Index           =   1
               Left            =   240
               Top             =   480
               Width           =   255
            End
            Begin VB.Image imgMoveInfo 
               Height          =   255
               Index           =   0
               Left            =   240
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblMoveInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "BrightPowder"
               Height          =   255
               Index           =   15
               Left            =   600
               TabIndex        =   77
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label lblMoveInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "King's Rock"
               Height          =   255
               Index           =   14
               Left            =   600
               TabIndex        =   76
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label lblMoveInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Move"
               Height          =   255
               Index           =   13
               Left            =   600
               TabIndex        =   75
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblMoveInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Sound Move"
               Height          =   255
               Index           =   12
               Left            =   600
               TabIndex        =   74
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.TextBox txtMove 
            Height          =   285
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1935
         End
         Begin MSComctlLib.ListView MoveList 
            Height          =   4695
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   8281
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
            EndProperty
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "30% Chance"
            Height          =   255
            Index           =   17
            Left            =   3120
            TabIndex        =   98
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Effect:"
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
            Index           =   16
            Left            =   2280
            TabIndex        =   97
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Image imgMoveCompat 
            Height          =   495
            Index           =   2
            Left            =   4200
            ToolTipText     =   "True RBY"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Image imgMoveCompat 
            Height          =   495
            Index           =   1
            Left            =   3360
            ToolTipText     =   "True RBY"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Image imgMoveCompat 
            Height          =   495
            Index           =   0
            Left            =   2520
            ToolTipText     =   "True RBY"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lblMoveDesc 
            Caption         =   "Raises Special Attack and Special Defense by focusing the mind."
            Height          =   615
            Left            =   2280
            TabIndex        =   68
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Reaction Target"
            Height          =   255
            Index           =   11
            Left            =   3120
            TabIndex        =   67
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "10 (16)"
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   66
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "95%"
            Height          =   255
            Index           =   9
            Left            =   3120
            TabIndex        =   65
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "120"
            Height          =   255
            Index           =   8
            Left            =   3120
            TabIndex        =   64
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Physical"
            Height          =   255
            Index           =   7
            Left            =   3120
            TabIndex        =   63
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Normal"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   62
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Target:"
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
            Index           =   5
            Left            =   2280
            TabIndex        =   61
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
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
            Index           =   4
            Left            =   2280
            TabIndex        =   60
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Acc.:"
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
            Index           =   3
            Left            =   2280
            TabIndex        =   59
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
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
            Index           =   2
            Left            =   2280
            TabIndex        =   58
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Line Line1 
            X1              =   2160
            X2              =   2160
            Y1              =   240
            Y2              =   5280
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Base:"
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
            Index           =   1
            Left            =   2280
            TabIndex        =   57
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblMoveInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
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
            Index           =   0
            Left            =   2280
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame fraDexText 
         Caption         =   "Cystal Text"
         Height          =   1335
         Index           =   2
         Left            =   6120
         TabIndex        =   51
         Top             =   2340
         Width           =   2895
         Begin VB.Label lblDexText 
            BackStyle       =   0  'Transparent
            Height          =   975
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame fraDexText 
         Caption         =   "Silver Text"
         Height          =   1335
         Index           =   1
         Left            =   3120
         TabIndex        =   49
         Top             =   2340
         Width           =   2895
         Begin VB.Label lblDexText 
            BackStyle       =   0  'Transparent
            Height          =   975
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame fraDexText 
         Caption         =   "Gold Text"
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   2340
         Width           =   2895
         Begin VB.Label lblDexText 
            BackStyle       =   0  'Transparent
            Caption         =   $"MasterDex.frx":55EF
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search..."
         Height          =   375
         Left            =   7800
         TabIndex        =   29
         Top             =   5520
         Width           =   1215
      End
      Begin VB.PictureBox picStatBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2595
         ScaleWidth      =   4335
         TabIndex        =   16
         Top             =   3780
         Width           =   4395
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "60"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   660
            TabIndex        =   92
            Top             =   2220
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "60"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   660
            TabIndex        =   91
            Top             =   1860
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "60"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   660
            TabIndex        =   90
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "60"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   660
            TabIndex        =   89
            Top             =   1140
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "60"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   660
            TabIndex        =   88
            Top             =   780
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "255"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   660
            TabIndex        =   87
            Top             =   420
            Width           =   375
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Min - Max"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   3360
            TabIndex        =   86
            Top             =   60
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Base"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   720
            TabIndex        =   85
            Top             =   60
            Width           =   2415
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Stat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   60
            TabIndex        =   84
            Top             =   60
            Width           =   495
         End
         Begin VB.Line Line3 
            X1              =   3240
            X2              =   3240
            Y1              =   120
            Y2              =   2520
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "700 - 550"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   28
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "700 - 550"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   27
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "700 - 550"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   26
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "700 - 550"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   3360
            TabIndex        =   25
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "700 - 550"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   24
            Top             =   780
            Width           =   855
         End
         Begin VB.Label lblStat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "444"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   23
            Top             =   420
            Width           =   855
         End
         Begin VB.Line Line2 
            X1              =   600
            X2              =   600
            Y1              =   120
            Y2              =   2520
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "SDef"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   22
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "SAtk"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   21
            Top             =   1860
            Width           =   495
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Spd"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   20
            Top             =   1500
            Width           =   495
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Def"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   19
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Atk"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   18
            Top             =   780
            Width           =   495
         End
         Begin VB.Label lblStat 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "HP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.TextBox txtPokemon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   420
         TabIndex        =   15
         Text            =   "Pokemon"
         Top             =   1920
         Width           =   1740
      End
      Begin VB.PictureBox picPokeImage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   180
         ScaleHeight     =   1035
         ScaleWidth      =   2145
         TabIndex        =   13
         Top             =   540
         Width           =   2200
         Begin VB.Image imgFrontImage 
            Height          =   825
            Left            =   1200
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Image imgBackImage 
            Height          =   825
            Left            =   120
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.PictureBox picTypeChart 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   6000
         Left            =   -74880
         ScaleHeight     =   5940
         ScaleWidth      =   5940
         TabIndex        =   8
         Top             =   420
         Width           =   6000
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   10
            ToolTipText     =   "Defender"
            Top             =   0
            Width           =   195
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   9
            ToolTipText     =   "Attacker"
            Top             =   180
            Width           =   135
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   6480
            X2              =   6480
            Y1              =   0
            Y2              =   6540
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   -68640
         ScaleHeight     =   975
         ScaleWidth      =   2475
         TabIndex        =   3
         Top             =   420
         Width           =   2535
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Attack will do 2x damage"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   660
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Attack will do 1/2 damage"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Attack will do no damage"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   4
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
      Begin VB.CommandButton Command1 
         Caption         =   "Dual &Type..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67200
         TabIndex        =   2
         Top             =   5520
         Width           =   1215
      End
      Begin MSComctlLib.TreeView EvoTree 
         Height          =   1575
         Left            =   6480
         TabIndex        =   11
         Top             =   540
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2778
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   5
         Scroll          =   0   'False
         ImageList       =   "EvoImages"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView MovePool 
         Height          =   2655
         Left            =   4620
         TabIndex        =   14
         Top             =   3780
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Move"
            Object.Width           =   2884
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Learned By"
            Object.Width           =   1958
         EndProperty
      End
      Begin VB.HScrollBar hscPokemon 
         Height          =   285
         Left            =   180
         Max             =   389
         TabIndex        =   12
         Top             =   1920
         Value           =   1
         Width           =   2200
      End
      Begin VB.Label DexNum 
         Caption         =   "#100"
         Height          =   255
         Left            =   240
         TabIndex        =   95
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label NatDexNum 
         Alignment       =   1  'Right Justify
         Caption         =   "National #100"
         Height          =   255
         Left            =   1080
         TabIndex        =   96
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "[No Second Trait]"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   42
         Top             =   1950
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Compound Eyes"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   41
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Water 1 / Water 2"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   40
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Egg Group:"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   39
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Water / Dark"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   38
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
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
         Index           =   0
         Left            =   2760
         TabIndex        =   37
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trait(s):"
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
         Index           =   6
         Left            =   2760
         TabIndex        =   35
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         Index           =   5
         Left            =   2760
         TabIndex        =   34
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Weight:"
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
         Index           =   4
         Left            =   2760
         TabIndex        =   33
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   32
         Top             =   780
         Width           =   975
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
         Left            =   -68520
         TabIndex        =   7
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "88.5m  (23' 10"")"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   36
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "555.0 kg  (222.3 lbs)"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   30
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Male 12.5%  |  Female 87.5%"
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   31
         Top             =   1260
         Width           =   2175
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Exit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "&Version"
      Begin VB.Menu mnuVersionItem 
         Caption         =   "&Red/Blue/Yellow"
         Index           =   0
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "&Gold/Silver/Crystal"
         Index           =   1
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Ruby/&Sapphire"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MasterDex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type PokeListType
    Listing() As String
    Index(0 To 255) As Integer
End Type
Private Const EM_GETLINECOUNT = &HBA
Dim ImageNum As Byte
Dim KeyChangeOK(3) As Boolean
Dim PokeList(2) As PokeListType
Dim CurrentPoke As Integer
Dim CurrentMove As Integer
Dim CurrentItem As Integer
Dim CurrentTrait As Integer
Dim MaxBaseStat(6) As Integer
Dim hKeyDown As Boolean
Dim EggGroupText(14) As String
Public CurrentMode As Byte
Public SearchOpen As Boolean

Dim Cycle As Boolean
Dim Min As Long
Dim Max As Long
Dim Loading As Boolean


Sub SetMode(ByVal NewMode As Byte)
    If NewMode < 0 Or NewMode > 3 Then Exit Sub
    'CurrentMode = SetMode
    MasterTab.Tab = NewMode
End Sub

Sub SetVer(ByVal NewVer As Integer)
    CurrentMode = 3
    If NewVer < 0 Or NewVer > 2 Then Exit Sub
    Call mnuVersionItem_Click(NewVer)
End Sub

Sub SetPoke(ByVal NewPoke As Integer)
    If NewPoke < 1 Then Exit Sub
    If CurrentMode = 0 And NewPoke > 151 Then Exit Sub
    If CurrentMode = 1 And NewPoke > 251 Then Exit Sub
    Call DoPoke(NewPoke)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Not SearchOpen Then
        SearchOpen = True
        Search.Show vbModeless, Me
    End If
End Sub

Private Sub Command4_Click()
    If Not SearchOpen Then
        SearchOpen = True
        Search.Show vbModeless, Me
    End If
    Call Search.Reset
    Search.cmbSearch(3).ListIndex = TraitList.SelectedItem.Index - 1
    Search.SCheck(3).Value = 1
    Call Search.DoSearch
    MasterTab.Tab = 0
End Sub
Public Sub FillSearchEGs()
    Dim X As Long
    With Search
        .cmbSearch(5).AddItem "(None)", 0
        For X = 1 To 14
            .cmbSearch(4).AddItem EggGroupText(X), X - 1
            .cmbSearch(5).AddItem EggGroupText(X), X
        Next X
    End With
End Sub
Private Sub Command5_Click()
    If Not SearchOpen Then
        SearchOpen = True
        Search.Show vbModeless, Me
    End If
    Call Search.Reset
    With Search.MoveList.ListItems(MoveList.SelectedItem.Key)
        .Checked = True
        .Selected = True
        .EnsureVisible
    End With
    Search.SCheck(9).Value = 1
    Call Search.DoSearch
    MasterTab.Tab = 0
End Sub

Private Sub EvoTree_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Expanded = True
End Sub

Private Sub EvoTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Call DoPoke(GetPokeNum(Node.Text))
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim CurrentIcon As Integer
    Loading = True
    EggGroupText(0) = "No Eggs"
    EggGroupText(1) = "Monster"
    EggGroupText(2) = "Water 1"
    EggGroupText(3) = "Bug"
    EggGroupText(4) = "Flying"
    EggGroupText(5) = "Ground"
    EggGroupText(6) = "Fairy"
    EggGroupText(7) = "Plant"
    EggGroupText(8) = "Humanshape"
    EggGroupText(9) = "Water 3"
    EggGroupText(10) = "Mineral"
    EggGroupText(11) = "Indeterminate"
    EggGroupText(12) = "Water 2"
    EggGroupText(13) = "Ditto"
    EggGroupText(14) = "Dragon"
    
    imgMoveCompat(0).Picture = LoadResPicture("RBY", vbResIcon)
    imgMoveCompat(1).Picture = LoadResPicture("GSC", vbResIcon)
    imgMoveCompat(2).Picture = LoadResPicture("ADV", vbResIcon)
    imgMoveCompat(3).Picture = LoadResPicture("GSC", vbResIcon)
    imgMoveCompat(4).Picture = LoadResPicture("ADV", vbResIcon)

    ReDim PokeList(0).Listing(1 To 151)
    ReDim PokeList(1).Listing(1 To 251)
    ReDim PokeList(2).Listing(1 To 389)
    For X = 1 To 389
        If X < 152 Then PokeList(0).Listing(X) = BasePKMN(X).Name
        If X < 252 Then PokeList(1).Listing(X) = BasePKMN(X).Name
        PokeList(2).Listing(X) = BasePKMN(X).Name
    Next X
    For X = 0 To 2
        Call SortStringArray(PokeList(X).Listing)
        Z = 0
        For Y = 1 To UBound(PokeList(X).Listing)
            A = Asc(LCase(Left(PokeList(X).Listing(Y), 1)))
            If Z <> A Then
                PokeList(X).Index(A) = Y
                Z = A
            End If
        Next Y
    Next X
    Call GetMaxes
    For X = 1 To UBound(AttributeText)
        TraitList.ListItems.Add , "#" & Format(X, "00"), AttributeText(X)
    Next X
    TraitList.Sorted = True
    TraitList.ListItems(1).Selected = True
    Call TraitList_ItemClick(TraitList.SelectedItem)
    Call DrawBattleDex
    
    cmbType.ListIndex = 2
    cmbWeather.ListIndex = 1
    cmbMove.Visible = False
    For X = 1 To 354
        If Moves(X).Power > 0 Then cmbMove.AddItem Moves(X).Name
    Next X
    cmbMove.Visible = True
    cmbPoke.Visible = True
    For X = 1 To 386
        cmbPoke.AddItem BasePKMN(X).Name
    Next X
    cmbPoke.Visible = True
    DemoBar.Caption = nbExact
    
    Call FillMoveList
    Call FillItemList
    'Me.Show
    Call SetVer(GetSetting("NetBattle", "Options", "PokeDexMode", 2))
    'Call DoPoke(1, True)
    picDamage.Visible = False
    cmbPoke.Text = vbNullString
    Loading = False
End Sub

Private Sub DrawBattleDex()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Long
    If CurrentMode = 0 Then
        A = 15
        picTypeChart.Width = 5340
        picTypeChart.Height = 5340
    Else
        A = 17
        picTypeChart.Width = 6000
        picTypeChart.Height = 6000
    End If
    picTypeChart.Picture = Nothing
    DispImg(0) = MainContainer.Conditions.ListImages(10).Picture
    DispImg(1) = MainContainer.Conditions.ListImages(9).Picture
    DispImg(2) = MainContainer.Conditions.ListImages(1).Picture
    Z = 330 '(picTypeChart.Height - 100) \ 18
    picTypeChart.Line (0, Z * 2)-(picTypeChart.Width, Z * 7), picTypeChart.FillColor, BF
    picTypeChart.Line (0, Z * 11)-(picTypeChart.Width, Z * 12), picTypeChart.FillColor, BF
    picTypeChart.Line (0, Z * 15)-(picTypeChart.Width, Z * 17), picTypeChart.FillColor, BF
    For X = 1 To A + 1
        picTypeChart.Line (X * Z, 0)-(X * Z, picTypeChart.Height), vbBlack
        picTypeChart.Line (0, X * Z)-(picTypeChart.Width, X * Z), vbBlack
    Next X
    picTypeChart.Line (0, 0)-(Z, Z), vbBlack
    For X = 1 To A
        picTypeChart.PaintPicture MainContainer.Types.ListImages(X).Picture, X * Z + 60, 60
        picTypeChart.PaintPicture MainContainer.Types.ListImages(X).Picture, 60, X * Z + 60
        For Y = 1 To A
            Select Case BattleMatrixEx(X, Y, (CurrentMode = 0))
            Case 0
                picTypeChart.PaintPicture MainContainer.Conditions.ListImages(10).Picture, Y * Z + 60, X * Z + 60
            Case 0.5
                picTypeChart.PaintPicture MainContainer.Conditions.ListImages(9).Picture, Y * Z + 60, X * Z + 60
            Case 2
                picTypeChart.PaintPicture MainContainer.Conditions.ListImages(1).Picture, Y * Z + 60, X * Z + 60
            End Select
        Next Y
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "NetBattle", "Options", "PokeDexMode", CurrentMode
End Sub

Private Sub hscPokemon_Change()
    If hscPokemon.Value = 0 Then hscPokemon.Value = hscPokemon.Max - 1
    If hscPokemon.Value = hscPokemon.Max Then hscPokemon.Value = 1
    Call DoPoke(hscPokemon.Value)
End Sub

Private Sub hscPokemon_KeyDown(KeyCode As Integer, Shift As Integer)
    hKeyDown = True
End Sub

Private Sub hscPokemon_KeyUp(KeyCode As Integer, Shift As Integer)
    hKeyDown = False
End Sub

Private Sub ItemList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim X As Long
    txtItem.Text = Item.Text
    X = Val(Right(Item.Key, 2))
    lblItemDesc.Caption = ItemDesc(X)
    imgMoveCompat(3).Visible = (X <= 41)
    imgMoveCompat(4).Visible = AdvItem(X)
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Unload Me
'    Select Case Index
'        Case 0
'            'Insert print code here
'        Case 2
'            Unload Me
'        Case 3
'            End
'    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute 0, vbNullString, "http://www.netbattle.net", vbNullString, vbNullString, 0
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Private Sub Label16_Click()

End Sub

Private Sub mnuVersionItem_Click(Index As Integer)
    Dim X As Long
    Dim Y As Long
    Dim Temp As String
    If CurrentMode = Index Then Exit Sub
    CurrentMode = Index
    'SaveSetting "NetBattle", "Options", "PokeDexMode", Index
    mnuVersionItem(0).Checked = (Index = 0)
    mnuVersionItem(1).Checked = (Index = 0)
    mnuVersionItem(2).Checked = (Index = 0)
    If CurrentMode = 0 And CurrentPoke > 151 Then CurrentPoke = 1
    If CurrentMode = 1 And CurrentPoke > 251 Then CurrentPoke = 1
    If CurrentMode = 2 Then TraitFrame.Visible = True Else TraitFrame.Visible = False
    If CurrentMode = 0 Then ItemFrame.Visible = False Else ItemFrame.Visible = True
    ImageNum = 0
    Call GetMaxes
    If CurrentPoke = 0 Then CurrentPoke = 1
    Call DoPoke(CurrentPoke, True)
    Call FillMoveList
    Call FillItemList
    If SearchOpen Then Call Search.RefreshMode
    hscPokemon.Max = UBound(PokeList(CurrentMode).Listing) + 1
    lblInfo(2).Visible = (CurrentMode > 0)
    lblInfo(5).Visible = (CurrentMode > 0)
    lblInfo(7).Visible = (CurrentMode > 0)
    lblInfo(12).Visible = (CurrentMode > 0)
    lblInfo(6).Visible = (CurrentMode = 2)
    lblInfo(8).Visible = (CurrentMode = 2)
    lblInfo(9).Visible = (CurrentMode = 2 And BasePKMN(CurrentPoke).PAtt(1) <> nbNoTrait)
    lblMoveInfo(12).Visible = (CurrentMode = 2)
    lblMoveInfo(13).Visible = (CurrentMode = 2)
    imgMoveInfo(0).Visible = (CurrentMode = 2)
    imgMoveInfo(1).Visible = (CurrentMode = 2)
    Frame5.Visible = (CurrentMode > 0)
    
    chkDamageCalc(1).Enabled = (CurrentMode > 0)
    lblDamageCalc(13).Enabled = (CurrentMode > 0)
    cmbWeather.Enabled = (CurrentMode > 0)
    chkDamageCalc(3).Enabled = (CurrentMode = 2)
    optReflect(2).Enabled = (CurrentMode = 2)
    Slider1.Enabled = (CurrentMode = 2)
    Slider2.Enabled = (CurrentMode = 2)
    optNature(0).Enabled = (CurrentMode = 2)
    optNature(1).Enabled = (CurrentMode = 2)
    optNature(2).Enabled = (CurrentMode = 2)
    optDef(1).Caption = IIf(CurrentMode = 0, "Special", "Sp. Def")
    Select Case CurrentMode
    Case 0: X = 151
    Case 1: X = 251
    Case 2: X = 386
    End Select
    If CurrentMode < 2 Then
        optNature(1).Value = True
        If optReflect(2).Value Then optReflect(1).Value = True
        chkDamageCalc(3).Value = 0
    End If
    If CurrentMode < 1 Then
        chkDamageCalc(1).Value = 0
        cmbWeather.ListIndex = 1
    End If
    cmbMove.Visible = False
    Temp = cmbMove.List(cmbMove.ListIndex)
    cmbMove.Clear
    For X = 1 To 354
        If Moves(X).Power > 0 Then
            Y = 0
            Select Case CurrentMode
            Case 0: If Moves(X).RBYMove Then Y = 1
            Case 1: If Moves(X).GSCMove Then Y = 1
            Case 2: If Moves(X).AdvMove Then Y = 1
            End Select
            If Y = 1 Then
                cmbMove.AddItem Moves(X).Name
                If Moves(X).Name = Temp Then
                    For Y = 0 To cmbMove.ListCount - 1
                        If cmbMove.List(Y) = Temp Then cmbMove.ListIndex = Y
                    Next Y
                End If
            End If
        End If
    Next X
    cmbMove.Visible = True
    Select Case CurrentMode
    Case 0: X = 151
    Case 1: X = 251
    Case 2: X = 386
    End Select
    cmbPoke.Visible = True
    Temp = cmbPoke.List(cmbPoke.ListIndex)
    cmbPoke.Clear
    For X = 1 To X
        cmbPoke.AddItem BasePKMN(X).Name
        If BasePKMN(X).Name = Temp Then
            For Y = 0 To cmbMove.ListCount - 1
                If cmbPoke.List(Y) = Temp Then cmbPoke.ListIndex = Y
            Next Y
        End If
    Next X
    If cmbPoke.ListIndex = -1 Then cmbPoke.ListIndex = 0
    cmbPoke.Visible = True
    Call Slider2_Click
    Call UpdateDefense
    Call DoCalc

    Call DrawBattleDex
End Sub

Public Sub FillMoveList()
    Dim X As Integer
    Dim Y As Boolean
    Dim Temp As String
    On Error Resume Next
    Temp = MoveList.SelectedItem.Key
    MoveList.Visible = False
    MoveList.ListItems.Clear
    For X = 1 To 354
        Select Case CurrentMode
        Case 0: Y = Moves(X).RBYMove
        Case 1: Y = Moves(X).GSCMove
        Case 2: Y = Moves(X).AdvMove
        End Select
        With ConvertMove(Moves(X), CurrentMode)
            If Y Then MoveList.ListItems.Add , "#" & Format(X, "000"), .Name, .Type, .Type
        End With
    Next X
    MoveList.Sorted = True
    MoveList.ListItems(Temp).Selected = True
    MoveList.SelectedItem.EnsureVisible
    MoveList.Visible = True
    Call MoveList_ItemClick(MoveList.SelectedItem)
    If SearchOpen Then Call Search.CopyMoveList
End Sub

Private Sub MoveList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    X = Val(Right(Item.Key, 3))
    'If X = CurrentMove Then Exit Sub
    CurrentMove = X
    With ConvertMove(Moves(X), CurrentMode)
        txtMove.Text = .Name
        lblMoveDesc.Caption = .Text
        lblMoveInfo(6).Caption = Element(.Type)
        Select Case .Type
        Case 2 To 6, 11, 15, 16
            Temp = "Special"
        Case Else
            Temp = "Physical"
        End Select
        If .ID = 91 Then Temp = "Varies"
        lblMoveInfo(7).Caption = Temp
        lblMoveInfo(8).Caption = IIf(.Power = 0, "---", .Power)
        lblMoveInfo(9).Caption = IIf(.Accuracy = 0, "---", .Accuracy & "%")
        lblMoveInfo(17).Caption = IIf(.SpecialPercent = 0, "---", .SpecialPercent & "% Chance")
        Y = Int(.PP * 1.6)
        If CurrentMode <> 2 And Y = 64 Then Y = 61
        lblMoveInfo(10).Caption = .PP & " (" & Y & ")"
        Select Case .Target
        Case nbGlobal: lblMoveInfo(11).Caption = "N/A"
        Case nbSelectedTarget: lblMoveInfo(11).Caption = "Selected Target"
        Case nbBothEnemies: lblMoveInfo(11).Caption = "Both Foes"
        Case nbSelfAffecting: lblMoveInfo(11).Caption = "Self Affecting"
        Case nbTeamAffecting: lblMoveInfo(11).Caption = "Team Affecting"
        Case nbEveryoneElse: lblMoveInfo(11).Caption = "Everyone Else"
        Case nbRandomEnemy: lblMoveInfo(11).Caption = "Random Enemy"
        Case nbReactionTarget: lblMoveInfo(11).Caption = "Reaction Target"
        Case nbMoveCalled: lblMoveInfo(11).Caption = "Move Called"
        End Select
        imgMoveCompat(0).Visible = .RBYMove
        imgMoveCompat(1).Visible = .GSCMove
        imgMoveCompat(2).Visible = .AdvMove
        imgMoveInfo(0).Picture = MainContainer.Conditions.ListImages(IIf(.PhysMove, 1, 9)).Picture
        imgMoveInfo(1).Picture = MainContainer.Conditions.ListImages(IIf(.SoundMove, 1, 9)).Picture
        imgMoveInfo(2).Picture = MainContainer.Conditions.ListImages(IIf(.KingsRock, 1, 9)).Picture
        imgMoveInfo(3).Picture = MainContainer.Conditions.ListImages(IIf(.BrightPowder, 1, 9)).Picture
    End With
End Sub

Private Sub MovePool_DblClick()
    On Error Resume Next
    MoveList.ListItems(MovePool.SelectedItem.Key).Selected = True
    MoveList.SelectedItem.EnsureVisible
    Call MoveList_ItemClick(MoveList.SelectedItem)
    MasterTab.Tab = 1
End Sub

Private Sub picPokeImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImageNum = (ImageNum + 1) Mod IIf(CurrentMode = 0, 3, 4)
    Call RefreshPokeImages
End Sub

Private Sub picTypeChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim Z As Integer
    X1 = X \ 330
    Y1 = Y \ 330
    If X1 > 17 Or Y1 > 17 Then
        picTypeChart.ToolTipText = ""
        Exit Sub
    End If
    If X1 = 0 Or Y1 = 0 Then
        If X1 = 0 And Y1 = 0 Then
            picTypeChart.ToolTipText = ""
        ElseIf X1 = 0 Then
            picTypeChart.ToolTipText = Element(Y1)
        Else
            picTypeChart.ToolTipText = Element(X1)
        End If
    Else
        Select Case BattleMatrixEx(Y1, X1, (CurrentMode = 0))
        Case 0
            picTypeChart.ToolTipText = Element(X1) & " types are immune to " & Element(Y1) & " attacks"
        Case 0.5
            picTypeChart.ToolTipText = Element(X1) & " types are resistant to " & Element(Y1) & " attacks"
        Case 1
            picTypeChart.ToolTipText = ""
        Case 2
            picTypeChart.ToolTipText = Element(X1) & " types are weak to " & Element(Y1) & " attacks"
        End Select
    End If
End Sub
Public Sub DoPoke(No As Integer, Optional Refreshing As Boolean = False)
    Dim X As Single
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim F As Single
    Dim C As Long
    Dim C2 As Long
    Dim TempMove() As Integer
    Dim TempSource() As String
    Dim TempItem As ListItem
    On Error Resume Next
    
    If No = CurrentPoke And Not Refreshing Then Exit Sub
    CurrentPoke = No
    Select Case No
    Case Is < 1: Exit Sub
    Case Is > 151: If CurrentMode = 0 Then CurrentMode = 1
    Case Is > 251: If CurrentMode < 2 Then CurrentMode = 2
    Case Is > UBound(BasePKMN): Exit Sub
    End Select
    SetRedraw Me.hWnd, False
    For X = 0 To 2
        mnuVersionItem(X).Checked = (CurrentMode = X)
    Next X
    Me.Caption = "NetBattle PokéDex  -  " & Replace(mnuVersionItem(CurrentMode).Caption, "&", "")
    
    Call RefreshPokeImages
    Call DoEvoChart
    With BasePKMN(No)
        'Set the basic info
        hscPokemon.Value = No
        KeyChangeOK(0) = False
        txtPokemon.Text = .Name
        KeyChangeOK(0) = True
        lblInfo(1).Caption = Element(.Type1)
        If Not ((No = 81 Or No = 82) And CurrentMode = 0) Then
            If .Type2 > 0 Then lblInfo(1).Caption = lblInfo(1).Caption & " / " & Element(.Type2)
        End If
        X = .Height
        F = X * 0.32808
        Y = Int(F)
        Z = Round((F - Y) * 12)
        lblInfo(10).Caption = Format(X / 10, "0.0") & "m  (" & Y & "' " & Format(Z, "00") & Chr$(34) & ")"
        X = .Weight
        lblInfo(11).Caption = Format(X / 10, "0.0") & "kg  (" & Format(X * 0.22046, "0.0") & " lbs)"
        If .PercentFemale = -1 Then
            lblInfo(12).Caption = "Gender Unknown"
        Else
            X = 100 * .PercentFemale / 16
            lblInfo(12).Caption = "Male " & CStr(100 - X) & "%  |  Female " & CStr(X) & "%"
        End If
        lblInfo(7).Caption = EggGroupText(.EggGroup1)
        If .EggGroup2 > 0 Then lblInfo(7).Caption = lblInfo(7).Caption & " / " & EggGroupText(.EggGroup2)
        If CurrentMode = 2 Then
            lblInfo(8).Caption = AttributeText(.PAtt(0))
            If .PAtt(1) = nbNoTrait Then
                lblInfo(6).Top = 1740
                lblInfo(9).Visible = False
            Else
                lblInfo(6).Top = 1740
                lblInfo(9).Caption = AttributeText(.PAtt(1))
                lblInfo(9).Visible = True
            End If
        Else
            lblInfo(9).Visible = False
        End If
        'Set the PokeDex Text
        Select Case CurrentMode
        Case 0
            fraDexText(3).Caption = "Red/Blue Text"
            fraDexText(4).Caption = "Yellow Text"
            lblDexText(3).Caption = PDexText(No).RedBlue
            lblDexText(4).Caption = PDexText(No).Yellow
            Y = 0
        Case 1
            lblDexText(0).Caption = PDexText(No).Gold
            lblDexText(1).Caption = PDexText(No).Silver
            lblDexText(2).Caption = PDexText(No).Crystal
            Y = 1
        Case 2
            fraDexText(3).Caption = "Ruby Text"
            fraDexText(4).Caption = "Sapphire Text"
            lblDexText(3).Caption = PDexText(No).Ruby
            lblDexText(4).Caption = PDexText(No).Sapphire
            Y = 3
        End Select
        fraDexText(0).Visible = (CurrentMode = 1)
        fraDexText(1).Visible = (CurrentMode = 1)
        fraDexText(2).Visible = (CurrentMode = 1)
        fraDexText(3).Visible = (CurrentMode <> 1)
        fraDexText(4).Visible = (CurrentMode <> 1)
        For X = 0 To 4
            If fraDexText(X).Visible Then
                txtSwap.Width = lblDexText(X).Width
                txtSwap.Text = Replace(lblDexText(X).Caption, "-", "ii")
                Z = SendMessage(txtSwap.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
                lblDexText(X).Top = 240 + (lblDexText(X).Height - Me.TextHeight("A") * Z) \ 2
            End If
        Next X
        
        'Fill up the move pool
        Call MakeMoveArray(.No, Y, TempMove, TempSource)
        MovePool.ListItems.Clear
        For X = 1 To UBound(TempMove)
            Y = TempMove(X)
            If Y > 0 Then
                Set TempItem = MovePool.ListItems.Add(, "#" & Format(Y, "000"), Moves(Y).Name, Moves(Y).Type, Moves(Y).Type)
                TempItem.SubItems(1) = TempSource(X)
            End If
        Next X
        MovePool.SortKey = 0
        MovePool.Sorted = True
        
        'Now the crazy part... Stats
        lblStat(14) = IIf(CurrentMode = 2, "Min - Max", "Max")
        lblStat(4).Caption = IIf(CurrentMode = 0, "Spcl", "SAtk")
        lblStat(5).Visible = (CurrentMode <> 0)
        lblStat(11).Visible = (CurrentMode <> 0)
        lblStat(20).Visible = (CurrentMode <> 0)
        picStatBox.Picture = Nothing
        picStatBox.FillColor = RGB(200, 200, 255)
        For X = 1 To 6
            Select Case X
            Case 1: Y = .BaseHP
            Case 2: Y = .BaseAttack
            Case 3: Y = .BaseDefense
            Case 4: Y = .BaseSpeed
            Case 5: Y = IIf(CurrentMode = 0, .BaseSpecial, .BaseSAttack)
            Case 6: If CurrentMode = 0 Then Exit For Else Y = .BaseSDefense
            End Select
            'These are here for easier tweeking.  Substitute real values later
            Const Bar2Offset = 80
            Const BarLeft = 1080
            Const BarVOffset = 60
            Const BarVMulti = 360
            Const BarHMulti = 15.5
            Const BarHeight = 180
            'If Y > 127 Then
            '    C = vbBlue
            '    C2 = &HFFC0C0
            'ElseIf Y > 63 Then
            '    C = vbGreen
            '    C2 = &HC0FFC0
            'ElseIf Y > 31 Then
            '    C = &HC0C0&
            '    C2 = &H80FFFF
            'Else
            '    C = vbRed
            '    C2 = &HC0C0FF
            'End If
            C = vbBlue
            C2 = &HFFC0C0
            
            lblStat(X + 14).Caption = Y
            Z = BarLeft + Cap(Y, 127) * BarHMulti
            picStatBox.FillColor = C2
            picStatBox.Line (BarLeft, X * BarVMulti + BarVOffset)- _
                            (Z, X * BarVMulti + BarVOffset + BarHeight) _
                            , C, B
            If Y > 127 Then
                picStatBox.FillColor = C2
                'Z = BarLeft + (Y - 127) * BarHMulti + Bar2Offset
                Z = BarLeft + (Y - 128) * BarHMulti
                'picStatBox.Line (BarLeft + Bar2Offset, X * BarVMulti + BarVOffset + Bar2Offset)- _
                '                (Z, X * BarVMulti + BarVOffset + BarHeight + Bar2Offset) _
                '                , vbBlue, B
                picStatBox.Line (BarLeft, X * BarVMulti + BarVOffset + Bar2Offset)- _
                                (Z, X * BarVMulti + BarVOffset + BarHeight + Bar2Offset) _
                                , C, B
            End If
            If X = 1 Then
                If CurrentMode = 2 Then
                    lblStat(X + 5).Caption = GetAdvHP(Y, 31, 0, 100) & " - " & GetAdvHP(Y, 31, 255, 100)
                Else
                    lblStat(X + 5).Caption = GetHP(100, Y, 15)
                End If
            Else
                If CurrentMode = 2 Then
                    lblStat(X + 5).Caption = GetAdvStat(Y, 31, 0, 100, 0) & " - " & GetAdvStat(Y, 31, 255, 100, 0)
                Else
                    lblStat(X + 5).Caption = GetStat(100, Y, 15)
                End If
            End If
        Next X
    End With
    If No <= 389 Then
        Select Case CurrentMode
            Case 0
                DexNum.Caption = "#" & No
                NatDexNum.Caption = ""
            Case 1
                DexNum.Caption = "#" & BasePKMN(No).GSNo
                NatDexNum.Caption = "National #" & No
            Case 2
                DexNum.Caption = "#" & BasePKMN(No).AdvNo
                NatDexNum.Caption = "National #" & No
        End Select
    Else
        Select Case CurrentMode
            Case 0
                DexNum.Caption = "#" & BasePKMN(386).No
                NatDexNum.Caption = ""
            Case 1
                DexNum.Caption = "#" & BasePKMN(386).GSNo
                NatDexNum.Caption = "National #386"
            Case 0
                DexNum.Caption = "#" & BasePKMN(386).AdvNo
                NatDexNum.Caption = "National #386"
        End Select
    End If
    SetRedraw Me.hWnd, Not Loading
    If hKeyDown Then DoEvents
End Sub
Private Sub RefreshPokeImages()
    Dim TempPoke As Pokemon
    Dim X As Integer
    Dim Y As Integer
    Dim A As GFXModes
    Dim B As Integer
    TempPoke.No = CurrentPoke
    A = ImageNum \ 2
    Select Case CurrentMode
    Case 0
        Select Case ImageNum
        Case 0, 3: A = nbGFXRB
        Case 1: A = nbGFXYlo
        Case 2: A = nbGFXGrn
        End Select
    Case 1
        Select Case ImageNum \ 2
        Case 0: A = nbGFXGld
        Case 1: A = nbGFXSil
        End Select
        TempPoke.Shiny = (ImageNum Mod 2 = 1)
    Case 2
        Select Case ImageNum \ 2
        Case 0: A = nbGFXRS
        Case 1: A = nbGFXLF
        End Select
        TempPoke.Shiny = (ImageNum Mod 2 = 1)
    End Select
    
    picPokeImage.Picture = Nothing
    Call MainContainer.DoPicture(ChooseImage(TempPoke, A))
    picSwap.Picture = MainContainer.SwapSpace.Picture
    B = GetYOffset(picSwap) * Screen.TwipsPerPixelY
    Call SetTransPixels(picSwap, picPokeImage.BackColor)
    X = picPokeImage.Width * 3 \ 4 - picSwap.Width \ 2
    Y = (picPokeImage.Height - picSwap.Height - B) \ 2
    picPokeImage.PaintPicture picSwap.Picture, X, Y
    
    Call MainContainer.DoPicture(ChooseImage(TempPoke, A, True))
    picSwap.Picture = MainContainer.SwapSpace.Picture
    B = GetYOffset(picSwap) * Screen.TwipsPerPixelY
    Call SetTransPixels(picSwap, picPokeImage.BackColor)
    X = (picPokeImage.Width \ 4 - picSwap.Width \ 2)
    Y = (picPokeImage.Height - picSwap.Height - B) \ 2
    picPokeImage.PaintPicture picSwap.Picture, X, Y
End Sub
Sub DoEvoChart()
    Dim TopLevel(6) As Integer
    Dim MidLevel(6) As Integer
    Dim LstLevel(6) As Integer
    Dim TopLevelKey As String
    Dim MidLevelKey As String
    Dim X As Byte
    On Error Resume Next
    EvoTree.Nodes.Clear
    With BasePKMN(CurrentPoke)
        Select Case .MyStage
            Case 0
                EvoTree.Nodes.Add , , .Name, .Name, 1, 1
            Case 1
                TopLevel(0) = .No
            Case 2
                MidLevel(0) = .No
            Case 3
                LstLevel(0) = .No
        End Select
        If .No = 265 Then
            EvoTree.Nodes.Add , , .Name, .Name, 1, 1
            EvoTree.Nodes.Add .Name, tvwChild, BasePKMN(.Evo(1)).Name, BasePKMN(.Evo(1)).Name, 3, 3
            EvoTree.Nodes.Add BasePKMN(.Evo(1)).Name, tvwChild, BasePKMN(.Evo(2)).Name, BasePKMN(.Evo(2)).Name, 3, 3
            EvoTree.Nodes.Add .Name, tvwChild, BasePKMN(.Evo(3)).Name, BasePKMN(.Evo(3)).Name, 3, 3
            EvoTree.Nodes.Add BasePKMN(.Evo(3)).Name, tvwChild, BasePKMN(.Evo(4)).Name, BasePKMN(.Evo(4)).Name, 3, 3
            EvoTree.Nodes(.Name).Expanded = True
            EvoTree.Nodes(BasePKMN(.Evo(1)).Name).Expanded = True
            EvoTree.Nodes(BasePKMN(.Evo(3)).Name).Expanded = True
        Else
            If .MyStage > 0 Then
                For X = 1 To 5
                    If .Evo(X) > 0 Then
                        Select Case BasePKMN(.Evo(X)).MyStage
                            Case 1
                                TopLevel(X) = .Evo(X)
                            Case 2
                                MidLevel(X) = .Evo(X)
                            Case 3
                                LstLevel(X) = .Evo(X)
                        End Select
                    End If
                Next
                If TopLevel(0) > 0 Then
                    If .MyMethod = 0 Then
                        EvoTree.Nodes.Add , , .Name, .Name, 1, 1
                    Else
                        EvoTree.Nodes.Add , , .Name, .Name, .MyMethod + 2, .MyMethod + 2
                    End If
                    TopLevelKey = .Name
                End If
                For X = 1 To 6
                    If TopLevel(X) > 0 Then
                        If BasePKMN(TopLevel(X)).MyMethod = 0 Then
                            EvoTree.Nodes.Add , , BasePKMN(TopLevel(X)).Name, BasePKMN(TopLevel(X)).Name, 1, 1
                        Else
                            EvoTree.Nodes.Add , , BasePKMN(TopLevel(X)).Name, BasePKMN(TopLevel(X)).Name, BasePKMN(TopLevel(X)).MyMethod + 2, BasePKMN(TopLevel(X)).MyMethod + 2
                        End If
                        TopLevelKey = BasePKMN(TopLevel(X)).Name
                    End If
                Next
                If MidLevel(0) > 0 Then
                    EvoTree.Nodes.Add TopLevelKey, tvwChild, .Name, .Name, .MyMethod + 2, .MyMethod + 2
                    MidLevelKey = .Name
                End If
                For X = 1 To 6
                    If MidLevel(X) > 0 Then
                        EvoTree.Nodes.Add TopLevelKey, tvwChild, BasePKMN(MidLevel(X)).Name, BasePKMN(MidLevel(X)).Name, .EvoM(X) + 2, .EvoM(X) + 2
                        MidLevelKey = BasePKMN(MidLevel(X)).Name
                    End If
                Next
                If LstLevel(0) > 0 Then
                    EvoTree.Nodes.Add MidLevelKey, tvwChild, .Name, .Name, .MyMethod + 2, .MyMethod + 2
                End If
                For X = 1 To 6
                    If LstLevel(X) > 0 Then
                        EvoTree.Nodes.Add MidLevelKey, tvwChild, BasePKMN(LstLevel(X)).Name, BasePKMN(LstLevel(X)).Name, .EvoM(X) + 2, .EvoM(X) + 2
                    End If
                Next
                EvoTree.Nodes(TopLevelKey).Expanded = True
                EvoTree.Nodes(MidLevelKey).Expanded = True
            End If
            EvoTree.Nodes(.Name).Selected = True
        End If
    End With
End Sub

Private Sub Picture5_Click()

End Sub

Private Sub TraitList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblTraitDesc.Caption = AttributeDesc(Val(Right(Item.Key, 2)))
    txtTrait.Text = Item.Text
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    Temp = FutureText(txtItem, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With ItemList
        Y = Len(Temp)
        For X = 1 To .ListItems.count
            If LCase(Left(.ListItems(X).Text, Y)) = LCase(Temp) Then
                ItemList.ListItems(X).Selected = True
                ItemList.ListItems(X).EnsureVisible
                Call ItemList_ItemClick(.ListItems(X))
                txtItem.Text = .ListItems(X).Text
                txtItem.SelStart = Y
                txtItem.SelLength = Len(txtItem.Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtItem.SelStart + 1
            txtItem.Text = Temp
            txtItem.SelStart = X
        End If
    End With

End Sub

Private Sub txtMove_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    Temp = FutureText(txtMove, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With MoveList
        Y = Len(Temp)
        For X = 1 To .ListItems.count
            If LCase(Left(.ListItems(X).Text, Y)) = LCase(Temp) Then
                MoveList.ListItems(X).Selected = True
                MoveList.ListItems(X).EnsureVisible
                Call MoveList_ItemClick(MoveList.ListItems(X))
                txtMove.Text = .ListItems(X).Text
                txtMove.SelStart = Y
                txtMove.SelLength = Len(txtMove.Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtMove.SelStart + 1
            txtMove.Text = Temp
            txtMove.SelStart = X
        End If
    End With
End Sub

Private Sub txtPokemon_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    Temp = FutureText(txtPokemon, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With PokeList(CurrentMode)
        Z = Asc(LCase(Left(Temp, 1)))
        If .Index(Z) = 0 Then Exit Sub
        Y = Len(Temp)
        For X = .Index(Z) To UBound(.Listing)
            If Asc(LCase(Left(.Listing(X), 1))) <> Z Then Exit For
            If LCase(Left(.Listing(X), Y)) = LCase(Temp) Then
                If .Listing(X) <> BasePKMN(CurrentPoke).Name Then Call DoPoke(GetPokeNum(.Listing(X)))
                txtPokemon.Text = .Listing(X)
                txtPokemon.SelStart = Y
                txtPokemon.SelLength = Len(txtPokemon.Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtPokemon.SelStart + 1
            txtPokemon.Text = Temp
            txtPokemon.SelStart = X
        End If
    End With
End Sub


Private Sub txtTrait_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    Temp = FutureText(txtTrait, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With TraitList
        Y = Len(Temp)
        For X = 1 To .ListItems.count
            If LCase(Left(.ListItems(X).Text, Y)) = LCase(Temp) Then
                TraitList.ListItems(X).Selected = True
                TraitList.ListItems(X).EnsureVisible
                Call TraitList_ItemClick(TraitList.ListItems(X))
                txtTrait.Text = .ListItems(X).Text
                txtTrait.SelStart = Y
                txtTrait.SelLength = Len(txtTrait.Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtTrait.SelStart + 1
            txtTrait.Text = Temp
            txtTrait.SelStart = X
        End If
    End With
End Sub

Private Sub GetMaxes()
    Dim X As Integer
    Dim LastPoke As Integer
    
    Select Case CurrentMode
        Case 0
            LastPoke = 151
        Case 1
            LastPoke = 251
        Case 2
            LastPoke = 389
    End Select
    For X = 1 To 6
        MaxBaseStat(X) = 0
    Next
    For X = 1 To LastPoke
        If BasePKMN(X).BaseHP > MaxBaseStat(1) Then MaxBaseStat(1) = BasePKMN(X).BaseHP
        If BasePKMN(X).BaseAttack > MaxBaseStat(2) Then MaxBaseStat(2) = BasePKMN(X).BaseAttack
        If BasePKMN(X).BaseDefense > MaxBaseStat(3) Then MaxBaseStat(3) = BasePKMN(X).BaseDefense
        If BasePKMN(X).BaseSpeed > MaxBaseStat(4) Then MaxBaseStat(4) = BasePKMN(X).BaseSpeed
        If CurrentMode = 0 Then
            If BasePKMN(X).BaseSpecial > MaxBaseStat(5) Then MaxBaseStat(5) = BasePKMN(X).BaseSpecial
        Else
            If BasePKMN(X).BaseSAttack > MaxBaseStat(5) Then MaxBaseStat(5) = BasePKMN(X).BaseSAttack
            If BasePKMN(X).BaseSDefense > MaxBaseStat(6) Then MaxBaseStat(6) = BasePKMN(X).BaseSDefense
        End If
    Next
End Sub

Sub FillItemList()
    Dim X As Integer
    
    ItemList.ListItems.Clear
    Select Case CurrentMode
        Case 0, 1
            For X = 1 To 41
                ItemList.ListItems.Add , "#" & Format(X, "00"), Item(X)
            Next X
        Case 2
            For X = 1 To UBound(Item)
                If AdvItem2(X) Then ItemList.ListItems.Add , "#" & Format(X, "00"), Item(X)
            Next X
    End Select
    ItemList.Sorted = True
    If ItemList.ListItems.count > 0 Then
        ItemList.ListItems(1).Selected = True
        Call ItemList_ItemClick(ItemList.SelectedItem)
    End If
End Sub
Public Sub SetVersion(ByVal NewVersion As Integer)
    CurrentMode = 3
    Call mnuVersionItem_Click(NewVersion)
End Sub





Private Sub chkDamageCalc_Click(Index As Integer)
    Call DoCalc
End Sub

Private Sub cmbPoke_Click()
    Dim X As Long
    Dim Y As Long
    Dim s As Single
    X = GetPokeNum(cmbPoke.List(cmbPoke.ListIndex))
    If X = imgDefender.Tag Then Exit Sub
    imgDefender.Tag = X
    Call MainContainer.DoPicture(ChooseImage(BasePKMN(X), nbGFXRS))
    imgDefender.Picture = MainContainer.SwapSpace
    DemoBar.Max = GetAdvHP(BasePKMN(X).BaseHP, 31, Slider2.Value, Val(txtDamageCalc(4).Text))
    DemoBar.Value = DemoBar.Max
    DemoBar.RefreshBar
    If cmbMove.ListIndex > -1 Then
        Y = GetMoveNum(cmbMove.List(cmbMove.ListIndex))
        Y = ConvertMove(Moves(Y), CurrentMode).Type
        s = BattleMatrixEx(Y, BasePKMN(X).Type1, (CurrentMode = nbRBYBattle))
        If BasePKMN(X).Type2 > nbNoType Then
            s = s * BattleMatrixEx(Y, BasePKMN(X).Type2, (CurrentMode = nbRBYBattle))
        End If
        Select Case s
        Case 0.25
            cmbType.ListIndex = 4
        Case 0.5
            cmbType.ListIndex = 3
        Case 1
            cmbType.ListIndex = 2
        Case 2
            cmbType.ListIndex = 1
        Case 4
            cmbType.ListIndex = 0
        End Select
    End If
    Call UpdateDefense
    picDamage.Visible = True
End Sub

Private Sub cmbPoke_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmbPoke.SelStart = Len(cmbPoke.Text)
        Exit Sub
    End If
    Temp = FutureText(cmbPoke, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With cmbPoke
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call cmbPoke_Click
                .Text = .List(X)
                .ListIndex = X
                .SelStart = Y
                .SelLength = Len(.List(X)) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = .SelStart + 1
            .Text = Temp
            .SelStart = X
        End If
    End With
End Sub

Private Sub cmbMove_Click()
    Dim X As Integer
    X = GetMoveNum(cmbMove.List(cmbMove.ListIndex))
    txtDamageCalc(3).Text = Moves(X).Power
    Select Case Moves(X).Type
    Case 2 To 6, 11, 15, 16
        optDef(1).Value = True
    Case Else
        optDef(0).Value = True
    End Select
    Call DoCalc
End Sub

Private Sub cmbMove_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmbMove.SelStart = Len(cmbMove.Text)
        Exit Sub
    End If
    Temp = FutureText(cmbMove, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With cmbMove
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call cmbMove_Click
                .Text = .List(X)
                .ListIndex = X
                .SelStart = Y
                .SelLength = Len(.List(X)) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = .SelStart + 1
            .Text = Temp
            .SelStart = X
        End If
    End With
End Sub

Private Sub cmbType_Click()
    Call DoCalc
End Sub

Private Sub cmbWeather_Click()
    Call DoCalc
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub optDef_Click(Index As Integer)
    Call UpdateDefense
End Sub

Private Sub optNature_Click(Index As Integer)
    Call UpdateDefense
End Sub

Private Sub optReflect_Click(Index As Integer)
    Call DoCalc
End Sub

Private Sub Slider1_Click()
    Slider1.Value = (Slider1.Value \ 4) * 4
    Call UpdateDefense
    lblDamageCalc(9).Caption = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
    Call Slider1_Click
End Sub

Private Sub Slider2_Click()
    Slider2.Value = (Slider2.Value \ 4) * 4
    lblDamageCalc(8).Caption = Slider2.Value
    If CurrentMode = 2 Then
        DemoBar.Max = GetAdvHP(BasePKMN(imgDefender.Tag).BaseHP, 31, Slider2.Value, Val(txtDamageCalc(4).Text))
    Else
        DemoBar.Max = GetHP(Val(txtDamageCalc(4).Text), BasePKMN(imgDefender.Tag).BaseHP, 15)
    End If
    DemoBar.RefreshBar
    Call UpdateDefense
End Sub

Private Sub Slider2_Scroll()
    Call Slider2_Click
End Sub

Private Sub Timer1_Timer()
    Cycle = Not Cycle
    DemoBar.Value = DemoBar.Max - Cap(IIf(Cycle, Min, Max), DemoBar.Max)
End Sub

Private Sub txtDamageCalc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtDamageCalc_LostFocus(Index)
        KeyAscii = 0
    End If
End Sub

Private Sub txtDamageCalc_LostFocus(Index As Integer)
    Dim X As Long
    X = Val(txtDamageCalc(Index).Text)
    If X < 1 And Index <> 3 Then X = 1
    If X < 0 And Index = 3 Then X = 0
    If (Index = 0 Or Index = 4) And X > 100 Then X = 100
    txtDamageCalc(Index).Text = CStr(X)
    If Index <> 2 Then Call Slider2_Click
    Call DoCalc
End Sub
Private Sub UpdateDefense()
    Dim Y As Integer
    Dim Z As Integer
    lblDamageCalc(9).Caption = IIf(Slider1.Enabled, Slider1.Value, "---")
    lblDamageCalc(8).Caption = IIf(Slider2.Enabled, Slider2.Value, "---")
    If imgDefender.Tag = 0 Then Exit Sub
    If optDef(0).Value Then Y = BasePKMN(imgDefender.Tag).BaseDefense Else Y = BasePKMN(imgDefender.Tag).BaseSDefense
    If optNature(0).Value Then Z = 1
    If optNature(2).Value Then Z = -1
    If CurrentMode = 2 Then
        Y = GetAdvStat(Y, 31, Slider1.Value, Val(txtDamageCalc(4).Text), Z)
    Else
        Y = GetStat(Val(txtDamageCalc(4).Text), Y, 15)
    End If
    txtDamageCalc(2).Text = CStr(Int(Y * StatChange(HScroll1.Value)))
    Call DoCalc
End Sub
Private Sub DoCalc()
    Dim ATK As Long
    Dim DEF As Long
    Dim Lev As Long
    Dim POW As Single
    Dim TMATCH As Single
    Dim STAB As Single
    Dim RAND As Integer
    Dim DamageTemp As Long
    Dim DT2 As Long
    Dim X As Long
    Dim Temp As String
    If Val(txtDamageCalc(3).Text) = 0 Then Exit Sub
    For X = 1 To 2
        Lev = Val(txtDamageCalc(0).Text)
        ATK = Val(txtDamageCalc(1).Text)
        DEF = Val(txtDamageCalc(2).Text)
        POW = Val(txtDamageCalc(3).Text)
        STAB = IIf(chkDamageCalc(0).Value = 1, 1.5, 1)
        If chkDamageCalc(1).Value = 1 Then POW = POW * 1.1
        If chkDamageCalc(2).Value = 1 Then Lev = Lev * 2
        If cmbWeather.ListIndex = 0 Then POW = POW * 1.5
        If cmbWeather.ListIndex = 2 Then POW = POW \ 2
        Select Case cmbType.ListIndex
        Case 0: TMATCH = 4
        Case 1: TMATCH = 2
        Case 2: TMATCH = 1
        Case 3: TMATCH = 0.5
        Case 4: TMATCH = 0.25
        End Select
        RAND = IIf(X = 1, 217, 255)
        DamageTemp = Int(((((((((2 * Lev) \ 5 + 2) * ATK * POW) \ DEF) / 50) + 2) * STAB) * TMATCH) * RAND) \ 255
        DT2 = ((((((((2 * Lev / 5 + 2) * ATK * POW) / DEF) / 50) + 2) * STAB) * TMATCH) * RAND) / 255
        Debug.Assert Abs(DT2 - DamageTemp) > 2
        If chkDamageCalc(3).Value = 1 Then DamageTemp = DamageTemp * 1.5
        If optReflect(1).Value Then DamageTemp = DamageTemp \ 2
        If optReflect(2).Value Then DamageTemp = Int(DamageTemp / 3 * 2)
        If DamageTemp = 0 Then DamageTemp = 1
        If X = 1 Then Min = DamageTemp Else Max = DamageTemp
    Next X

    Label6.Caption = "Damage: " & Min & " ~ " & Max
    If imgDefender.Tag > 0 Then
        If Min >= DemoBar.Max Then
            lblDamageCalc(7).Caption = "100% Damage"
        Else
            X = Round((Min / DemoBar.Max) * 100)
            If X = 100 Then X = 99
            Temp = "Between " & X
            X = Round((Cap(Max, DemoBar.Max) / DemoBar.Max) * 100)
            If Max < DemoBar.Max And X = 100 Then X = 99
            lblDamageCalc(7).Caption = Temp & "% and " & X & "% Damage"
        End If
        Cycle = False
        Call Timer1_Timer
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
        
End Sub


Private Sub HScroll1_Change()
    lblBattleMod.Caption = IIf(HScroll1.Value > 0, "+", "") & CStr(HScroll1.Value)
    Call UpdateDefense
End Sub

