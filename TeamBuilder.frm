VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TeamBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Team Builder"
   ClientHeight    =   5970
   ClientLeft      =   3600
   ClientTop       =   3555
   ClientWidth     =   9195
   Icon            =   "TeamBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9195
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox TBContainer 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   9135
      TabIndex        =   238
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdPlaceholder 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   237
         Top             =   6240
         Width           =   615
      End
      Begin VB.CommandButton cmdPlaceholder 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   52
         Top             =   6240
         Width           =   615
      End
      Begin VB.PictureBox picMoves 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   5295
         TabIndex        =   236
         Top             =   4860
         Width           =   5295
         Begin VB.TextBox txtMove 
            Height          =   285
            Index           =   4
            Left            =   4050
            MaxLength       =   12
            TabIndex        =   31
            Text            =   "Move 4"
            Top             =   0
            WhatsThisHelpID =   10059
            Width           =   1245
         End
         Begin VB.TextBox txtMove 
            Height          =   285
            Index           =   3
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   30
            Text            =   "Move 3"
            Top             =   0
            WhatsThisHelpID =   10059
            Width           =   1245
         End
         Begin VB.TextBox txtMove 
            Height          =   285
            Index           =   2
            Left            =   1350
            MaxLength       =   12
            TabIndex        =   29
            Text            =   "Move 2"
            Top             =   0
            WhatsThisHelpID =   10059
            Width           =   1245
         End
         Begin VB.TextBox txtMove 
            Height          =   285
            Index           =   1
            Left            =   0
            MaxLength       =   12
            TabIndex        =   28
            Text            =   "Move 1"
            Top             =   0
            WhatsThisHelpID =   10059
            Width           =   1245
         End
      End
      Begin VB.ComboBox MasterItem 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Text            =   "MasterItem"
         Top             =   2280
         WhatsThisHelpID =   10040
         Width           =   1695
      End
      Begin VB.ComboBox MasterSpecies 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "MasterSpecies"
         Top             =   1080
         WhatsThisHelpID =   10038
         Width           =   1695
      End
      Begin NetBattle.CompressZIt CompressZIt1 
         Left            =   -120
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Arrange..."
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   5280
         WhatsThisHelpID =   10031
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Done"
         Height          =   375
         Left            =   4440
         TabIndex        =   0
         Top             =   5280
         WhatsThisHelpID =   10035
         Width           =   975
      End
      Begin VB.CommandButton ShowBox 
         Caption         =   "&Boxes >>"
         Height          =   375
         Left            =   3360
         TabIndex        =   51
         Top             =   5280
         WhatsThisHelpID =   10034
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save..."
         Height          =   375
         Left            =   2280
         TabIndex        =   50
         Top             =   5280
         WhatsThisHelpID =   10033
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Load..."
         Height          =   375
         Left            =   1200
         TabIndex        =   49
         Top             =   5280
         WhatsThisHelpID =   10032
         Width           =   975
      End
      Begin VB.Frame BoxFrame 
         Caption         =   "Pokémon Boxes"
         Height          =   5535
         Left            =   5640
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   5175
            Left            =   120
            ScaleHeight     =   5175
            ScaleWidth      =   3255
            TabIndex        =   55
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   0
               Picture         =   "TeamBuilder.frx":1272
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Move to another box"
               Top             =   1920
               WhatsThisHelpID =   10047
               Width           =   375
            End
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   0
               OLEDropMode     =   1  'Manual
               Picture         =   "TeamBuilder.frx":13BC
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Delete"
               Top             =   2400
               WhatsThisHelpID =   10048
               Width           =   375
            End
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   0
               Picture         =   "TeamBuilder.frx":1506
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Move down"
               Top             =   1440
               WhatsThisHelpID =   10046
               Width           =   375
            End
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   0
               Picture         =   "TeamBuilder.frx":1650
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Move up"
               Top             =   960
               WhatsThisHelpID =   10045
               Width           =   375
            End
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   0
               Picture         =   "TeamBuilder.frx":179A
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Copy Pokemon to box"
               Top             =   480
               WhatsThisHelpID =   10044
               Width           =   375
            End
            Begin VB.CommandButton BoxNav 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   0
               Picture         =   "TeamBuilder.frx":18E4
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Insert at current position"
               Top             =   0
               WhatsThisHelpID =   10043
               Width           =   375
            End
            Begin MSComctlLib.ListView PokeBox 
               Height          =   2775
               Left            =   480
               TabIndex        =   46
               Top             =   0
               WhatsThisHelpID =   10060
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   4895
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               OLEDragMode     =   1
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               OLEDragMode     =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Main"
                  Object.Width           =   3264
               EndProperty
            End
            Begin MSComctlLib.TabStrip BoxTabs 
               Height          =   5175
               Left            =   2520
               TabIndex        =   47
               Top             =   0
               WhatsThisHelpID =   10049
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   9128
               TabWidthStyle   =   2
               MultiRow        =   -1  'True
               Style           =   1
               TabFixedWidth   =   1759
               TabFixedHeight  =   582
               Placement       =   3
               TabMinWidth     =   883
               ImageList       =   "ImageList1"
               _Version        =   393216
               BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
                  NumTabs         =   10
                  BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 1"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 2"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 3"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 4"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 5"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 6"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 7"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 8"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 9"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     Caption         =   "Box 10"
                     ImageVarType    =   2
                     ImageIndex      =   1
                  EndProperty
               EndProperty
               OLEDropMode     =   1
            End
            Begin VB.Image lblMarker 
               Height          =   240
               Index           =   0
               Left            =   360
               Picture         =   "TeamBuilder.frx":1A2E
               Top             =   4920
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image lblMarker 
               Height          =   240
               Index           =   1
               Left            =   840
               Picture         =   "TeamBuilder.frx":1B00
               Top             =   4920
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image lblMarker 
               Height          =   240
               Index           =   2
               Left            =   1320
               Picture         =   "TeamBuilder.frx":1BD2
               Top             =   4920
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image lblMarker 
               Height          =   240
               Index           =   3
               Left            =   1800
               Picture         =   "TeamBuilder.frx":1CA4
               Top             =   4920
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image imgBoxVer 
               Height          =   240
               Left            =   2160
               Stretch         =   -1  'True
               Top             =   2880
               Width           =   240
            End
            Begin VB.Line InfoLine 
               Index           =   2
               X1              =   0
               X2              =   2400
               Y1              =   4560
               Y2              =   4560
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Held Item: Held Item"
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   209
               Top             =   4605
               Width           =   2415
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Move 4"
               Height          =   255
               Index           =   11
               Left            =   1260
               TabIndex        =   208
               Top             =   4305
               Width           =   1095
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Move 3"
               Height          =   255
               Index           =   10
               Left            =   45
               TabIndex        =   207
               Top             =   4305
               Width           =   1095
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Move 2"
               Height          =   255
               Index           =   9
               Left            =   1260
               TabIndex        =   206
               Top             =   4065
               Width           =   1095
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "Move 1"
               Height          =   255
               Index           =   8
               Left            =   45
               TabIndex        =   205
               Top             =   4065
               Width           =   1095
            End
            Begin VB.Line InfoLine 
               Index           =   1
               X1              =   0
               X2              =   2400
               Y1              =   4005
               Y2              =   4005
            End
            Begin VB.Line InfoLine 
               Index           =   0
               X1              =   0
               X2              =   2400
               Y1              =   3465
               Y2              =   3465
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SDef"
               Height          =   255
               Index           =   7
               Left            =   2025
               TabIndex        =   204
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SAtk"
               Height          =   255
               Index           =   6
               Left            =   1620
               TabIndex        =   203
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Spd"
               Height          =   255
               Index           =   5
               Left            =   1215
               TabIndex        =   202
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Def"
               Height          =   255
               Index           =   4
               Left            =   825
               TabIndex        =   201
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Atk"
               Height          =   255
               Index           =   3
               Left            =   420
               TabIndex        =   200
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "HP"
               Height          =   255
               Index           =   2
               Left            =   15
               TabIndex        =   199
               Top             =   3525
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   1
               Left            =   420
               TabIndex        =   198
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   2
               Left            =   825
               TabIndex        =   197
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   3
               Left            =   1215
               TabIndex        =   196
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   4
               Left            =   1620
               TabIndex        =   195
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   5
               Left            =   2025
               TabIndex        =   194
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblStat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "999"
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   193
               Top             =   3765
               Width           =   375
            End
            Begin VB.Label lblInfo 
               Caption         =   "Lv.100 Pokemon (M)"
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   192
               Top             =   3165
               Width           =   1815
            End
            Begin VB.Label lblInfo 
               Caption         =   "Nickname"
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   191
               Top             =   2925
               Width           =   1455
            End
            Begin VB.Image imgInfo 
               Height          =   480
               Left            =   75
               Top             =   2880
               Width           =   480
            End
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TeamBuilder.frx":1D76
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TeamBuilder.frx":2310
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TeamBuilder.frx":28AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TeamBuilder.frx":2E44
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TeamBuilder.frx":33DE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip MainTabs 
         Height          =   735
         Left            =   720
         TabIndex        =   1
         Tag             =   "1340"
         Top             =   120
         WhatsThisHelpID =   10036
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1296
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         Style           =   1
         TabFixedWidth   =   2002
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   0
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   8
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Trainer"
               Object.ToolTipText     =   "Set trainer information"
               ImageVarType    =   2
               ImageIndex      =   3
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &1"
               Object.ToolTipText     =   "Pokémon #1"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &2"
               Object.ToolTipText     =   "Pokémon #2"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &3"
               Object.ToolTipText     =   "Pokémon #3"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &4"
               Object.ToolTipText     =   "Pokémon #4"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &5"
               Object.ToolTipText     =   "Pokémon #5"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PKMN &6"
               Object.ToolTipText     =   "Pokémon #6"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Compat."
               ImageVarType    =   2
               ImageIndex      =   5
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.PictureBox picSwap 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   6
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   165
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   5
            Left            =   4200
            OLEDropMode     =   2  'Automatic
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   166
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   5
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   5
            Left            =   0
            MaxLength       =   15
            TabIndex        =   15
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   1800
            TabIndex        =   9
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   5
            Left            =   1800
            TabIndex        =   32
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   5
            Left            =   0
            TabIndex        =   22
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   188
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   189
            Top             =   4080
            Width           =   5295
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   187
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   186
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   185
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   184
            Top             =   120
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   183
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   182
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   181
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   180
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   179
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   178
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   177
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   176
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   175
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   174
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   173
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   172
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3710
            TabIndex        =   171
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   170
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   5
            Left            =   3710
            TabIndex        =   169
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   5
            Left            =   4200
            TabIndex        =   168
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   167
            Top             =   1800
            Width           =   1935
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   5
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   210
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   4
            Left            =   1800
            TabIndex        =   33
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   4
            Left            =   0
            MaxLength       =   15
            TabIndex        =   16
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   4
            Left            =   4200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   211
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   4
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   4
            Left            =   0
            TabIndex        =   23
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   234
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   4
            Left            =   4200
            TabIndex        =   233
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3710
            TabIndex        =   232
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   231
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3710
            TabIndex        =   230
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   229
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   228
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   227
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   226
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   4
            Left            =   2640
            TabIndex        =   225
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   224
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   223
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   222
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   221
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   220
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   219
            Top             =   360
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   218
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   217
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   216
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   215
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   214
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   213
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   212
            Top             =   4080
            Width           =   5295
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   4
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   140
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   3
            Left            =   4200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   141
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   3
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   3
            Left            =   0
            MaxLength       =   15
            TabIndex        =   17
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   34
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   3
            Left            =   0
            TabIndex        =   24
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   164
            Top             =   4080
            Width           =   5295
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   163
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   162
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   161
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   160
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   159
            Top             =   120
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   158
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   157
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   156
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   155
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   154
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   153
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   152
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   151
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   150
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   149
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   148
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   147
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3710
            TabIndex        =   146
            Top             =   1560
            Width           =   1545
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   145
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   3
            Left            =   3710
            TabIndex        =   144
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   143
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   142
            Top             =   1800
            Width           =   1935
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   3
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   115
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   2
            Left            =   4200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   116
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   2
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   2
            Left            =   0
            MaxLength       =   15
            TabIndex        =   18
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   35
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   139
            Top             =   4080
            Width           =   5295
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   138
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   137
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   136
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   135
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   134
            Top             =   120
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   133
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   132
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   131
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   130
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   129
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   128
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   127
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   126
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   125
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   124
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   123
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   122
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3710
            TabIndex        =   121
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   120
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   2
            Left            =   3710
            TabIndex        =   119
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   118
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   117
            Top             =   1800
            Width           =   1935
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   2
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   90
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   36
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   1
            Left            =   0
            MaxLength       =   15
            TabIndex        =   19
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   1
            Left            =   4200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   91
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   1
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   114
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   113
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3710
            TabIndex        =   112
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   111
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3710
            TabIndex        =   110
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   109
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   108
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   107
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   106
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   105
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   104
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   103
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   102
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   101
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   100
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   99
            Top             =   360
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   98
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   97
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   96
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   95
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   93
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   92
            Top             =   4080
            Width           =   5295
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   1
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin VB.PictureBox PKMNPicBox 
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   0
            Left            =   4200
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   1035
            ScaleWidth      =   1035
            TabIndex        =   66
            Top             =   120
            Width           =   1095
            Begin VB.Image PKMNPic 
               Height          =   825
               Index           =   0
               Left            =   120
               OLEDropMode     =   1  'Manual
               Top             =   120
               Width           =   840
            End
         End
         Begin VB.TextBox Nickname 
            Height          =   285
            Index           =   0
            Left            =   0
            MaxLength       =   15
            TabIndex        =   20
            Top             =   840
            WhatsThisHelpID =   10039
            Width           =   1695
         End
         Begin VB.CommandButton Rebuild 
            Caption         =   "S&witch"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            WhatsThisHelpID =   10041
            Width           =   975
         End
         Begin VB.CommandButton ExpertBuild 
            Caption         =   "&Expert..."
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   37
            Top             =   760
            WhatsThisHelpID =   10042
            Width           =   975
         End
         Begin MSComctlLib.ListView MovePick 
            Height          =   1935
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Tag             =   "0"
            Top             =   2040
            WhatsThisHelpID =   10050
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Move"
               Object.Width           =   3149
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Power"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Acc."
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "PP"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Learned By"
               Object.Width           =   2275
            EndProperty
         End
         Begin VB.Label CurrMoves 
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   89
            Top             =   4080
            Width           =   5295
         End
         Begin VB.Label Label5 
            Caption         =   "Pokémon"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Nickname"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   86
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   85
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   84
            Top             =   120
            Width           =   615
         End
         Begin VB.Label HP 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   83
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack:"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   82
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Attack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   81
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense:"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   80
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Defense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   79
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed:"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   78
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Speed 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   77
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Attack:"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   76
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label SpecialAttack 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   75
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Defense:"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   74
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label SpecialDefense 
            Alignment       =   1  'Right Justify
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   73
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Type1:"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   72
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Type1 
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3710
            TabIndex        =   71
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type2:"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   70
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Type2 
            Caption         =   "???"
            Height          =   255
            Index           =   0
            Left            =   3710
            TabIndex        =   69
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label GenderDisp 
            Alignment       =   2  'Center
            Caption         =   "???"
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   68
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Moves"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   67
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label StatWarn 
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   88
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   7
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   190
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin MSComctlLib.TreeView CompatTree 
            Height          =   3015
            Left            =   0
            TabIndex        =   38
            Top             =   120
            WhatsThisHelpID =   10051
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   5318
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   882
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   6
            SingleSel       =   -1  'True
            Appearance      =   1
         End
         Begin VB.TextBox MoveAdvanced 
            Height          =   975
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   3240
            WhatsThisHelpID =   10052
            Width           =   4575
         End
         Begin VB.Line Line3 
            X1              =   5280
            X2              =   4680
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Line Line2 
            X1              =   5280
            X2              =   4680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line1 
            X1              =   5280
            X2              =   5280
            Y1              =   120
            Y2              =   4200
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   6
            Left            =   4680
            ToolTipText     =   "GSC Only"
            Top             =   1560
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   5
            Left            =   4680
            ToolTipText     =   "Advance (Trade)"
            Top             =   360
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   0
            Left            =   4680
            ToolTipText     =   "R/B/Y"
            Top             =   960
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   1
            Left            =   4680
            ToolTipText     =   "G/S/C"
            Top             =   2160
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   2
            Left            =   4680
            ToolTipText     =   "Advance (Legit)"
            Top             =   2880
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image OnePKMN 
            Height          =   480
            Index           =   3
            Left            =   4680
            ToolTipText     =   "Advance (Full)"
            Top             =   3480
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.PictureBox TabHolder 
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   0
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5415
         TabIndex        =   56
         Top             =   840
         Width           =   5415
         Begin VB.Frame Frame3 
            Caption         =   "Auto Messages"
            Height          =   2175
            Left            =   1440
            TabIndex        =   63
            Top             =   2160
            Width           =   3975
            Begin VB.TextBox WinMSG 
               Height          =   495
               Left            =   120
               MaxLength       =   240
               TabIndex        =   6
               Top             =   480
               WhatsThisHelpID =   10057
               Width           =   3735
            End
            Begin VB.TextBox LoseMSG 
               Height          =   495
               Left            =   120
               MaxLength       =   240
               TabIndex        =   7
               Top             =   1440
               WhatsThisHelpID =   10058
               Width           =   3735
            End
            Begin VB.Label Label1 
               Caption         =   "Win"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Lose"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   1200
               Width           =   2175
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Info"
            Height          =   1695
            Left            =   1440
            TabIndex        =   59
            Top             =   0
            Width           =   3975
            Begin VB.ComboBox VersionSelect 
               Height          =   315
               ItemData        =   "TeamBuilder.frx":3CB8
               Left            =   1920
               List            =   "TeamBuilder.frx":3CBA
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   480
               WhatsThisHelpID =   10055
               Width           =   1935
            End
            Begin VB.TextBox UserName 
               Height          =   315
               Left            =   120
               MaxLength       =   20
               TabIndex        =   2
               Top             =   480
               WhatsThisHelpID =   10054
               Width           =   1695
            End
            Begin VB.TextBox ExtraInfo 
               Height          =   495
               Left            =   120
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   4
               Top             =   1080
               WhatsThisHelpID =   10056
               Width           =   3735
            End
            Begin VB.Label Label3 
               Caption         =   "User Name"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Extra Info"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label13 
               Caption         =   "Graphics"
               Height          =   255
               Left            =   1920
               TabIndex        =   60
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Image"
            Height          =   4335
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   1335
            Begin MSComctlLib.ListView TrainerPics 
               Height          =   3975
               Left            =   120
               TabIndex        =   5
               Top             =   240
               WhatsThisHelpID =   10053
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   7011
               View            =   3
               Arrange         =   2
               LabelEdit       =   1
               Sorted          =   -1  'True
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
                  Text            =   "Pictures"
                  Object.Width           =   1806
               EndProperty
            End
         End
      End
      Begin VB.Image imgTBMode 
         Height          =   495
         Left            =   120
         Top             =   240
         WhatsThisHelpID =   10037
         Width           =   495
      End
      Begin VB.Image TrueGSC 
         Height          =   495
         Left            =   -360
         ToolTipText     =   "True GSC"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image TrueRBY 
         Height          =   495
         Left            =   -360
         ToolTipText     =   "True RBY"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image ADVTrade 
         Height          =   480
         Left            =   -360
         ToolTipText     =   "Advance with Trades"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image AdvPlus 
         Height          =   480
         Left            =   -360
         ToolTipText     =   "Advance Plus"
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image ADV 
         Height          =   480
         Left            =   -360
         ToolTipText     =   "Advance"
         Top             =   3360
         Width           =   480
      End
      Begin VB.Image GS 
         Height          =   480
         Left            =   -360
         ToolTipText     =   "GSC with Trades"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image RBY 
         Height          =   480
         Left            =   -360
         ToolTipText     =   "RBY with Trades"
         Top             =   1560
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   5715
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13600
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As..."
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Revert"
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Export to Text "
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close Team Builder"
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 (No Recent File)"
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 (No Recent File)"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 (No Recent File)"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 (No Recent File)"
         Enabled         =   0   'False
         Index           =   12
      End
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "&Version"
      Begin VB.Menu mnuVersionItem 
         Caption         =   "True &RBY"
         Index           =   0
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "RBY &with Trades"
         Index           =   1
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "True &GSC"
         Index           =   3
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "GSC with &Trades"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Ru/Sa Only"
         Index           =   6
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Full Advance"
         Index           =   7
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Mod:"
         Index           =   9
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "View Current Mods..."
         Index           =   10
      End
      Begin VB.Menu mnuVersionItem 
         Caption         =   "Load New Mod File..."
         Index           =   11
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "&Sort"
      Begin VB.Menu mnuSortItem 
         Caption         =   "By &Pokédex Number"
         Index           =   0
      End
      Begin VB.Menu mnuSortItem 
         Caption         =   "By &GSC Pokédex"
         Index           =   1
      End
      Begin VB.Menu mnuSortItem 
         Caption         =   "By &Advance Pokédex"
         Index           =   2
      End
      Begin VB.Menu mnuSortItem 
         Caption         =   "By &Name"
         Index           =   3
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
   Begin VB.Menu mnuMarker 
      Caption         =   "Marker"
      Visible         =   0   'False
      Begin VB.Menu mnuMarkerItem 
         Caption         =   "Circle"
         Index           =   0
      End
      Begin VB.Menu mnuMarkerItem 
         Caption         =   "Square"
         Index           =   1
      End
      Begin VB.Menu mnuMarkerItem 
         Caption         =   "Triangle"
         Index           =   2
      End
      Begin VB.Menu mnuMarkerItem 
         Caption         =   "Heart"
         Index           =   3
      End
   End
End
Attribute VB_Name = "TeamBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SendBack As Boolean
Public TeamLoadOK As Boolean
Private SkipDialog As Boolean
Private BlankPKMN As Pokemon
Private SwapTrainer As Trainer
Private OrigFilename As String
Private VerPoke As Integer
Private mintCurFrame As Integer ' Current Frame visible
Private LoadingForm As Boolean
Private DontUpdate As Boolean
Private Changed(6) As Boolean
Private LastTBMode As Byte
Private CurrentDisplay As Integer
Private ShuttingDown As Boolean
Private IDiedOnceAlready As Boolean
Private RecentLoad As String
Private OrigTeam As String
Private SaveTeam As String
Private CheckTeam As String
Private ShowBoxes As Boolean
Private Const TCM_FIRST = &H1300
Private Const TCM_HITTEST = (TCM_FIRST + 13)
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type TCHITTESTINFO
    pt As POINTAPI
    Flags As Long
End Type
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Sub RefreshCurrMoves()
    Dim X As Integer
    Dim Y As Integer
    X = mintCurFrame - 1
    If X = 0 Or X = 7 Then
        picMoves.Visible = False
        Exit Sub
    End If
    picMoves.Visible = True
    For Y = 1 To 4
        If PKMN(X).Move(Y) = 0 Then
            txtMove(Y).Tag = ""
            txtMove(Y).Text = ""
        Else
            txtMove(Y).Tag = "#" & Format(PKMN(X).Move(Y), "000")
            txtMove(Y).Text = Moves(PKMN(X).Move(Y)).Name
        End If
    Next Y
End Sub

Private Sub BoxNav_Click(Index As Integer)
    Dim Answer As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Long
    Dim HasMoves As Boolean
    Dim TempPKMN As Pokemon
    Dim ThisPKMN As Integer
    Dim SecondPKMN As Integer
    Dim TempMove() As Integer
    Dim TempItem As ListItem
    Dim Temp As String
    On Error Resume Next
    ThisPKMN = -1
    ThisPKMN = Val(PokeBox.SelectedItem.Tag)
    If ThisPKMN = -1 And Index <> 1 Then Exit Sub
    On Error GoTo 0
    With PokeBox
        Select Case Index
        'Left
        Case 0
            If mintCurFrame >= 2 And mintCurFrame <= 7 Then
                If PKMN(mintCurFrame - 1).No > 0 Then
                    Answer = MsgBox("Overwrite the current Pokémon?", vbYesNo + vbQuestion, "Confirm Change")
                    If Answer = vbNo Then Exit Sub
                End If
                Select Case BoxPKMN(ThisPKMN).No
                    Case 386, 387, 388, 389
                        For X = 1 To 6
                            If (PKMN(X).No = 386 Or PKMN(X).No = 387 Or PKMN(X).No = 388 Or PKMN(X).No = 389) _
                            And X <> mintCurFrame - 1 _
                            And BoxPKMN(ThisPKMN).No <> BoxPKMN(X).No Then
                                MsgBox "You already have a " & BasePKMN(386).Name & " on your team!", vbCritical, "Duplicate Pokémon"
                                Exit Sub
                            End If
                        Next
                    Case Else
'                        For X = 1 To 6
'                            If PKMN(X).No = BoxPKMN(ThisPKMN).No _
'                            And X <> mintCurFrame - 1 Then
'                                MsgBox "You already have a " & PKMN(X).Name & " on your team!", vbCritical, "Duplicate Pokémon"
'                                Exit Sub
'                            End If
'                        Next
                End Select
                TempPKMN = BoxPKMN(ThisPKMN)
                Select Case TBMode
                Case 0, 5: If Not TempPKMN.ExistRBY Then X = 1
                Case 1: If TempPKMN.No > 251 Then X = 1
                Case 2: If Not TempPKMN.ExistAdv Then X = 1
                Case 6: If Not TempPKMN.ExistGSC Then X = 1
                End Select
                If X = 1 Then
                    MsgBox "This Pokémon is not compatible with the current mode.", , "Error"
                    Exit Sub
                End If
                FillInPokeData TempPKMN, TBMode
                'TempPKMN.GameVersion = TBMode
                If LegalMove(TempPKMN) <> "" Then
                    Answer = MsgBox("This Pokémon's moveset is not compatible with the current mode.  Some moves will be reset.  Continue?", vbYesNo + vbQuestion, "Confirm Change")
                    If Answer = vbNo Then Exit Sub
                    TempMove = TempPKMN.Move
                    For X = 1 To 4
                        TempPKMN.Move(X) = 0
                    Next X
                    Y = 1
                    For X = 1 To 4
                        TempPKMN.Move(Y) = TempMove(X)
                        If LegalMove(TempPKMN) = "" Then
                            Y = Y + 1
                        Else
                            TempPKMN.Move(Y) = 0
                        End If
                    Next X
                End If
                PKMN(mintCurFrame - 1) = TempPKMN
                PKMN(mintCurFrame - 1).Image = ChooseImage(PKMN(mintCurFrame - 1), You.Version)
                Changed(mintCurFrame - 1) = True
                Call LoadSettings
            End If
        'Right
        Case 1
            
            If mintCurFrame >= 2 And mintCurFrame <= 7 Then
                Y = mintCurFrame - 1
            ElseIf mintCurFrame = 8 Then
                If CompatTree.SelectedItem.Key = "USER" Then Exit Sub
                Y = Val(Mid(CompatTree.SelectedItem.Key, 5, 1))
            Else
                Y = 0
            End If
            If Y > 0 Then
                HasMoves = False
                For X = 1 To 4
                    If PKMN(Y).Move(X) > 0 Then HasMoves = True
                Next
                If PKMN(Y).No = 0 Or Not HasMoves Then Exit Sub
                ReDim Preserve BoxPKMN(UBound(BoxPKMN) + 1) As Pokemon
                X = UBound(BoxPKMN)
                BoxPKMN(X) = PKMN(Y)
                If BoxPKMN(X).Nickname = "" Then BoxPKMN(X).Nickname = BoxPKMN(X).Name
                BoxPKMN(X).InBox = BoxTabs.SelectedItem.Index
                Set TempItem = PokeBox.ListItems.Add(, , BoxPKMN(X).Nickname & " (" & BoxPKMN(X).Name & ")")
                TempItem.Tag = X
                .ListItems(.ListItems.count).Selected = True
            End If
            
        'Up
        Case 2
            If .SelectedItem.Index = 1 Then Exit Sub
            X = .SelectedItem.Index
            Y = X - 1
            SecondPKMN = Val(.ListItems(Y).Tag)
            If SecondPKMN = -1 Then Exit Sub
            TempPKMN = BoxPKMN(ThisPKMN)
            BoxPKMN(ThisPKMN) = BoxPKMN(SecondPKMN)
            BoxPKMN(SecondPKMN) = TempPKMN
            Temp = .ListItems(Y).Text
            .ListItems(Y).Text = .ListItems(X).Text
            .ListItems(X).Text = Temp
            .ListItems(Y).Selected = True
            CurrentDisplay = .ListItems(X).Tag
        'Down
        Case 3
            If .SelectedItem.Index = .ListItems.count Then Exit Sub
            X = .SelectedItem.Index
            Y = X + 1
            SecondPKMN = Val(.ListItems(Y).Tag)
            If SecondPKMN = -1 Then Exit Sub
            TempPKMN = BoxPKMN(ThisPKMN)
            BoxPKMN(ThisPKMN) = BoxPKMN(SecondPKMN)
            BoxPKMN(SecondPKMN) = TempPKMN
            Temp = .ListItems(Y).Text
            .ListItems(Y).Text = .ListItems(X).Text
            .ListItems(X).Text = Temp
            .ListItems(Y).Selected = True
            CurrentDisplay = .ListItems(Y).Tag
        'Del
        Case 4
            Answer = MsgBox("Are you sure you want to delete this Pokémon?", vbYesNo + vbQuestion, "Confirm Delete")
            If Answer = vbNo Then Exit Sub
            For X = ThisPKMN To UBound(BoxPKMN) - 1
                BoxPKMN(X) = BoxPKMN(X + 1)
            Next
            ReDim Preserve BoxPKMN(UBound(BoxPKMN) - 1) As Pokemon
            X = .SelectedItem.Index
            .ListItems.Remove X
            If .ListItems.count <> 0 Then .ListItems(IIf(X > .ListItems.count, X - 1, X)).Selected = True
            For X = X To .ListItems.count
                .ListItems(X).Tag = .ListItems(X).Tag - 1
            Next
        'Move
        Case 5
            MoveBoxNum = ThisPKMN
            FromBox = BoxTabs.SelectedItem.Index
            If ToBox = -1 Then
                CopyFlag = False
                MovePKMN.Show vbModal
                If ToBox = -1 Then Exit Sub
            End If
            If CopyFlag Then
                ReDim Preserve BoxPKMN(UBound(BoxPKMN) + 1)
                BoxPKMN(UBound(BoxPKMN)) = BoxPKMN(ThisPKMN)
                BoxPKMN(UBound(BoxPKMN)).InBox = ToBox
            Else
                TempPKMN = BoxPKMN(ThisPKMN)
                For X = ThisPKMN To UBound(BoxPKMN) - 1
                    BoxPKMN(X) = BoxPKMN(X + 1)
                Next
                BoxPKMN(X) = TempPKMN
                BoxPKMN(X).InBox = ToBox
                X = .SelectedItem.Index
                .ListItems.Remove X
                If .ListItems.count <> 0 Then .ListItems(IIf(X > .ListItems.count, X - 1, X)).Selected = True
                For X = X To .ListItems.count
                    .ListItems(X).Tag = .ListItems(X).Tag - 1
                Next
            End If
            ToBox = -1
            CopyFlag = False
        End Select
    End With
    Call RefreshBoxTabs
    Call RefreshBoxMinor
End Sub
Private Sub RefreshBoxMinor()
    Dim X As Integer
    Dim Y As Boolean
    Dim Z As Integer
    Dim A As Long
    On Error Resume Next
    With PokeBox
        X = IIf(.ListItems.count > 12, 1600, 1850)
        If .ColumnHeaders(1).Width <> X Then .ColumnHeaders(1).Width = X
        
        For Z = 1 To .ListItems.count
            X = Val(.ListItems(Z).Tag)
            Y = False
            Select Case TBMode
            Case 0, 5: If Not BoxPKMN(X).ExistRBY Then Y = True
            Case 1: If BoxPKMN(X).No > 251 Then Y = True
            Case 2: If Not BoxPKMN(X).ExistAdv Then Y = True
            Case 6: If Not BoxPKMN(X).ExistGSC Then Y = True
            End Select
            A = IIf(Y, vbRed, vbBlack)
            If .ListItems(Z).ForeColor <> A Then .ListItems(Z).ForeColor = A
        Next Z
    
        If .ListItems.count = 0 Then
            BoxNav(0).Enabled = False
            BoxNav(2).Enabled = False
            BoxNav(3).Enabled = False
            BoxNav(4).Enabled = False
        Else
            BoxNav(0).Enabled = True
            BoxNav(4).Enabled = True
            BoxNav(2).Enabled = (.SelectedItem.Index <> 1)
            BoxNav(3).Enabled = (.SelectedItem.Index <> .ListItems.count)
            .SetFocus
        End If
        
        .SelectedItem.EnsureVisible
        For Z = 0 To 3
            lblMarker(Z).Visible = ((BoxPKMN(.SelectedItem.Tag).MarkerNum And 2 ^ Z) > 0)
        Next Z
    End With
    Call PokeBox_ItemClick(PokeBox.SelectedItem)
End Sub
Private Sub RefreshBox(Optional ByVal ListIndex = 1)
    Dim X As Integer
    Dim Y As Integer
    Dim BoxCount As Integer
    Dim TempItem As ListItem
    'LockWindowUpdate PokeBox.hwnd
    PokeBox.ListItems.Clear
    If UBound(BoxPKMN) = 0 Then
        BoxNav(0).Enabled = False
        BoxNav(2).Enabled = False
        BoxNav(3).Enabled = False
        BoxNav(4).Enabled = False
        For X = 1 To 10
            BoxTabs.Tabs(X).Image = 1
        Next X
        Call PokeBox_ItemClick(PokeBox.SelectedItem)
        Exit Sub
    End If
    For X = 1 To UBound(BoxPKMN)
        If BoxPKMN(X).InBox = BoxTabs.SelectedItem.Index Then
            Set TempItem = PokeBox.ListItems.Add(, , BoxPKMN(X).Nickname & " (" & BoxPKMN(X).Name & ")")
            TempItem.Tag = X
            Y = 0
            Select Case TBMode
            Case 0, 5: If Not BoxPKMN(X).ExistRBY Then Y = 1
            Case 1: If BoxPKMN(X).No > 251 Then Y = 1
            Case 2: If Not BoxPKMN(X).ExistAdv Then Y = 1
            Case 6: If Not BoxPKMN(X).ExistGSC Then Y = 1
            End Select
            If Y = 1 Then TempItem.ForeColor = vbRed
        End If
    Next
    Call RefreshBoxTabs
    If PokeBox.ListItems.count = 0 Then
        BoxNav(0).Enabled = False
        BoxNav(2).Enabled = False
        BoxNav(3).Enabled = False
        BoxNav(4).Enabled = False
    Else
        If PokeBox.ListItems.count >= ListIndex Then PokeBox.SelectedItem = PokeBox.ListItems(ListIndex)
        BoxNav(0).Enabled = True
        BoxNav(4).Enabled = True
        BoxNav(2).Enabled = (PokeBox.SelectedItem.Index <> 1)
        BoxNav(3).Enabled = (PokeBox.SelectedItem.Index <> PokeBox.ListItems.count)
        On Error Resume Next
        PokeBox.SetFocus
    End If
    Call PokeBox_ItemClick(PokeBox.SelectedItem)
    'LockWindowUpdate 0
End Sub
Private Sub RefreshBoxTabs()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    For X = 1 To 10
        Z = 0
        For Y = 1 To UBound(BoxPKMN)
            If BoxPKMN(Y).InBox = X Then Z = Z + 1
        Next
        If Z = 0 Then
            BoxTabs.Tabs(X).Image = 1
        Else
            BoxTabs.Tabs(X).Image = 2
        End If
    Next
End Sub
Private Sub BoxTabs_Click()
    Call RefreshBox
End Sub

Private Sub BoxTabs_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim HitTestInfo As TCHITTESTINFO
    Dim R As Long
    Dim Temp As String
    Dim Dummy As String
    Call BoxTabs_OLEDragOver(Data, Effect, Button, Shift, X, Y, 0)
    'Debug.Print Effect
    If Effect <> vbDropEffectCopy And Effect <> vbDropEffectMove Then Exit Sub
    Temp = Data.GetData(1)
    Dummy = ChopString(Temp, 10)
    HitTestInfo.pt.X = X / Screen.TwipsPerPixelX
    HitTestInfo.pt.Y = Y / Screen.TwipsPerPixelY
    R = SendMessage(BoxTabs.hWnd, TCM_HITTEST, 0&, HitTestInfo)
    ToBox = R + 1
    CopyFlag = (Effect = vbDropEffectCopy)
    Call BoxNav_Click(5)
    Effect = 3
End Sub

Private Sub BoxTabs_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim HitTestInfo As TCHITTESTINFO
    Dim R As Long
    Dim Temp As String
    On Error GoTo NoSel
    Temp = Data.GetData(1)
    If ChopString(Temp, 10) <> "PNBBOXDRAG" Then
NoSel:
        Effect = vbDropEffectNone
        Exit Sub
    End If
    HitTestInfo.pt.X = X / Screen.TwipsPerPixelX
    HitTestInfo.pt.Y = Y / Screen.TwipsPerPixelY
    R = SendMessage(BoxTabs.hWnd, TCM_HITTEST, 0&, HitTestInfo)
    If R = -1 Or Val(Temp) <> PokeBox.SelectedItem.Tag Then
        Effect = vbDropEffectNone
        Exit Sub
    ElseIf R = BoxTabs.SelectedItem.Index - 1 Then
        Effect = vbDropEffectNone
        Exit Sub
    End If
    If Shift And vbCtrlMask = vbCtrlMask Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectMove
    End If
End Sub

Private Sub cmdPlaceholder_GotFocus(Index As Integer)
    If Index = 1 Then
        ShowBox.SetFocus
    Else
        Command1.SetFocus
    End If
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim X As Integer
    TeamLoadOK = False
    TeamChangeFromTB = True
    Me.Enabled = False
    If RecentLoad = "" Then
        If Not TeamLoader.OpenTheFile Then
            TeamChangeFromTB = False
            Me.Enabled = True
            Me.ZOrder
            Me.SetFocus
            Exit Sub
        End If
    Else
        Call TeamLoader.ReadFile(RecentLoad)
        RecentLoad = ""
    End If
    TeamChangeFromTB = False
    SaveTeam = CreateTeamString
    'If Not TeamLoadOK Then Exit Sub
    For X = 1 To 6
        Changed(X) = True
    Next
    Call LoadSettings
    MainTabs.Tabs(1).Selected = True
    mnuFileItem(2).Enabled = True
    Me.Enabled = True
    Me.ZOrder
    Me.SetFocus
End Sub

Private Sub Command3_Click()
    Dim FileToUse As String
    Dim ByteArray() As Byte
    Dim HeaderBytes() As Byte
    Dim Worked As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim FileNum As Integer
    Dim Temp As String
        
    FileNum = FreeFile
    Call ApplySettings
    If StoredFileName <> "" And SkipDialog Then
        FileToUse = StoredFileName
    Else
        With MainContainer.FileBox
            .DialogTitle = "Save Trainer/Team"
            .Flags = cdlOFNOverwritePrompt
            .Filter = "Pokémon NetBattle File (*.pnb)|*.pnb|All Files (*.*)|*.*"
            .CancelError = True
            .DefaultExt = ".pnb"
            .FileName = ""
            Temp = GetSetting("NetBattle", "Options", "InitDir", "")
            If Temp <> "" Then .InitDir = Temp
            On Error GoTo Cancelled
            .ShowSave
            FileToUse = .FileName
            SaveSetting "NetBattle", "Options", "InitDir", Left$(FileToUse, InStrRev(FileToUse, "\"))
            StoredFileName = FileToUse
            mnuFileItem(2).Enabled = True
        End With
    End If
    With You
        Temp = " PNB4.1"
        Temp = Temp & Chr$(Len(.Name)) & .Name
        Temp = Temp & Chr$(Len(.Extra)) & .Extra
        Temp = Temp & Chr$(Len(.WinMess)) & .WinMess
        Temp = Temp & Chr$(Len(.LoseMess)) & .LoseMess
        Temp = Temp & Chr$(TBMode)
        Temp = Temp & Chr$(.Picture)
        Temp = Temp & Chr$(.Version)
    End With
    For X = 1 To 6
        Temp = Temp & PKMN2Str(PKMN(X))
    Next X
    If PKMN(1).GameVersion = nbModAdv Then
        Temp = Temp & Pad(DBModName, 20)
        Temp = Temp & DBModStr
    End If
    If FileExists(FileToUse) Then Kill FileToUse
    Open FileToUse For Binary Access Write As #FileNum
    Put #FileNum, 1, Temp
    Close #FileNum
'    ReDim ByteArray(FileLen(SlashPath & "team.tmp") - 1)
'    Open SlashPath & "team.tmp" For Binary Access Read As #FileNum
'        Get #FileNum, , ByteArray()
'    Close #FileNum
'    Kill SlashPath & "team.tmp"
'    Open SlashPath & "team.tmp" For Output As #FileNum
'        Write #FileNum, "PNB4.0"
'        Worked = CompressZIt1.CompressData(ByteArray())
'        Write #FileNum, CompressZIt1.OriginalSize
'    Close #FileNum
'    ReDim HeaderBytes(FileLen(SlashPath & "team.tmp") - 1)
'    Open SlashPath & "team.tmp" For Binary Access Read As #FileNum
'        Get #FileNum, , HeaderBytes()
'    Close #FileNum
'    Kill SlashPath & "team.tmp"
'    Open FileToUse For Binary Access Write As #FileNum
'        Put #FileNum, , HeaderBytes
'        Put #FileNum, , ByteArray
'    Close #FileNum
    Call UpdateListings(FileToUse)
    mnuFileItem(2).Enabled = True
    SkipDialog = False
    SaveTeam = CreateTeamString
Cancelled:
End Sub

Private Sub Command4_Click()
    Dim X As Integer
    Dim N(6) As Integer
    Dim B As Boolean
    For X = 1 To 6
        N(X) = PKMN(X).No
    Next X
    Rearrange.Show vbModal
    B = False
    For X = 1 To 6
        If PKMN(X).No <> N(X) Then
            Changed(X) = True
            B = True
        Else
            Changed(X) = False
        End If
    Next X
    If B Then Call LoadSettings
    X = MainTabs.SelectedItem.Index
    RefreshCurrMoves
End Sub



Private Sub ExpertBuild_Click(Index As Integer)
    Dim Duplicate As Boolean
    Dim Y As Integer
    Dim X As Integer
    Dim AttackVar As Integer
    Dim DefenseVar As Integer
    Dim SpeedVar As Integer
    Dim SpecialVar As Integer
    Dim Unown As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim i As Integer
    Dim Temp As Integer
    Dim NewSpecies As Integer
    Dim ClearList As Boolean
    Dim MaxMoves As Integer
    
    For X = 1 To UBound(BasePKMN)
        If BasePKMN(X).Name = MasterSpecies.List(MasterSpecies.ListIndex) Then NewSpecies = X
    Next
    
    Duplicate = False
    Select Case NewSpecies
        Case 386, 387, 388, 389
            For X = 1 To 6
                If (PKMN(X).No = 386 Or PKMN(X).No = 387 Or PKMN(X).No = 388 Or PKMN(X).No = 389) And X <> Index + 1 And PKMN(X).No <> NewSpecies Then
                    MsgBox "You already have a " & BasePKMN(386).Name & " on your team!", vbCritical, "Duplicate Pokémon"
                    Exit Sub
                End If
            Next
'        Case Else
'            For X = 1 To 6
'                If PKMN(X).No = NewSpecies And X <> Index + 1 Then Duplicate = True
'            Next
    End Select
    
    Changed(Index + 1) = True
    If PKMN(Index + 1).No = 0 Then
        If MasterSpecies.ListIndex < 0 Then Exit Sub
        Call Rebuild_Click(Index)
    End If
    
    'If Duplicate = True Then Exit Sub
    
    'Generate DVs
    ExpertPKMN = PKMN(Index + 1)
    ExpertPKMN.GameVersion = TBMode
    Select Case TBMode
        Case 0, 1, 5, 6
            Expert.Show 1
            PKMN(Index + 1) = ExpertPKMN
        Case 2, 3, 4
            AdvExpert.Show 1
            PKMN(Index + 1) = ExpertPKMN
    End Select
    
    Call FillInPokeData(PKMN(Index + 1), TBMode)
    
    With PKMN(Index + 1)
    
        'Adjust on-screen values
        If TBMode = nbTrueRBY Or TBMode = nbRBYTrade Then
            GenderDisp(Index).Caption = "Lv." & .Level
        Else
            GenderDisp(Index).Caption = Gender(.Gender) & " Lv." & .Level
        End If
        HP(Index) = .MaxHP
        Attack(Index) = .Attack
        Defense(Index) = .Defense
        Speed(Index) = .Speed
        SpecialAttack(Index) = .SpecialAttack
        SpecialDefense(Index) = .SpecialDefense
        Type1(Index).Caption = Element(.Type1)
        If (TBMode = nbTrueRBY Or TBMode = nbRBYTrade) And (.No = 81 Or .No = 82) Then
            Type2(Index).Caption = ""
        Else
            Type2(Index).Caption = Element(.Type2)
        End If
        'Set nickname
        'Nickname(Index).Text = BasePKMN(NewSpecies).Nickname
        Call RefreshImage(Index + 1, True)
        Call RefreshCurrMoves
    End With
    'PKMN(Index + 1).Item = GetItemNum(ItemPick(Index).List(ItemPick(Index).ListIndex))
    MaxMoves = MovePick(Index).ListItems.count
    If MaxMoves > 4 Then MaxMoves = 4
    For X = 1 To MaxMoves
        If PKMN(Index + 1).Move(X) > 0 Then Exit For
    Next
    If X <> MaxMoves + 1 Then MainTabs.Tabs(Index + 2).Image = 2 Else MainTabs.Tabs(Index + 2).Image = 1
    Call DoRBY
    Call DoPower
End Sub

Private Sub ExtraInfo_Change()
    If LoadingForm Then Exit Sub
    ExtraInfo.Text = Replace(ExtraInfo.Text, Chr(1), " ")
    You.Extra = ExtraInfo.Text
End Sub


Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim ThisCaption As String
    Dim FullItem As String
    Dim BoxCount As Integer
    Call WriteDebugLog("TB Load")
    'LockWindowUpdate Me.hwnd
    If TeamChangeFromMS Then
        If Not MasterServer.mnuOptionsItem(7).Checked Then
            Call MasterServer.SendData("AWAY:")
            SendBack = True
        End If
    End If
    DontUpdate = False
    SkipDialog = False
    IDiedOnceAlready = False
    CurrentDisplay = -2
    ToBox = -1
    CopyFlag = False
    OrigTeam = CreateTeamString
    SaveTeam = OrigTeam
    Call ReadBoxPKMN
    Call RefreshBox
    Call WriteDebugLog("Box Loaded")
    ShowBoxes = GetSetting("NetBattle", "Team Builder", "Show Boxes", False)
    Call DoBoxChange
    If StoredFileName <> "" And UCase(Right(StoredFileName, 4)) = ".PNB" Then
        mnuFileItem(2).Enabled = True
    Else
        mnuFileItem(2).Enabled = False
    End If
    OrigFilename = StoredFileName
    LoadingForm = True
    SwapTrainer = You
    For X = 1 To 6
        'PKMN(X) = StoredPKMN(X)
        Changed(X) = True
    Next
    LastTBMode = 7
    VersionSelect.Clear
    VersionSelect.AddItem "Green"
    VersionSelect.AddItem "Red/Blue"
    VersionSelect.AddItem "Yellow"
    VersionSelect.AddItem "Gold"
    VersionSelect.AddItem "Silver"
    VersionSelect.AddItem "Ruby/Sapphire"
    VersionSelect.AddItem "Leaf/Fire"
    VersionSelect.AddItem "Emerald"
    If HasColGFX Then VersionSelect.AddItem "Colosseum"
    VersionSelect.ListIndex = 6
    mnuSortItem(TBSort).Checked = True
'    TrueRBY.Picture = LoadResPicture("RBY", vbResIcon)
'    RBY.Picture = LoadResPicture("RBYT", vbResIcon)
'    GS.Picture = LoadResPicture("GSCT", vbResIcon)
'    TrueGSC.Picture = LoadResPicture("GSC", vbResIcon)
'    ADV.Picture = LoadResPicture("ADV", vbResIcon)
'    AdvPlus.Picture = LoadResPicture("ADV+", vbResIcon)
'    ADVTrade.Picture = LoadResPicture("MOD", vbResIcon)
    OnePKMN(0).Picture = LoadResPicture("RBYT", vbResIcon)
    OnePKMN(1).Picture = LoadResPicture("GSCT", vbResIcon)
    OnePKMN(2).Picture = LoadResPicture("ADV", vbResIcon)
    OnePKMN(3).Picture = LoadResPicture("ADV+", vbResIcon)
    'OnePKMN(4).Picture = LoadResPicture("MOD", vbResIcon)
    OnePKMN(5).Picture = LoadResPicture("RBY", vbResIcon)
    OnePKMN(6).Picture = LoadResPicture("GSC", vbResIcon)
    For X = 0 To 6
        If X <> 4 Then
            OnePKMN(X).ToolTipText = ModeText(X)
        End If
    Next
    If PKMN(1).No = 0 Then
        VersionSelect.ListIndex = nbGFXRS
        You.Version = nbGFXRS
    End If
    MainTabs.TabIndex = 1
    For X = 1 To 6
        TabHolder(X).Visible = False
    Next
    mintCurFrame = 1
    TrainerPics.Icons = MainContainer.Trainers
    TrainerPics.SmallIcons = MainContainer.MiniTrainers
    For X = 0 To 5
        MovePick(X).Icons = MainContainer.Types
        MovePick(X).SmallIcons = MainContainer.Types
    Next
    CompatTree.ImageList = MainContainer.Types
    For X = 1 To MainContainer.Trainers.ListImages.count
        If X < 10 Then ThisCaption = "#0" & X Else ThisCaption = "#" & X
        TrainerPics.ListItems.Add X, , ThisCaption, X, X
    Next
    Call WriteDebugLog("Entering LoadSettings")
    Call LoadSettings
'    X = GetSetting("NetBattle", "Options", "Show Boxes", 0)
'    If X = 1 Then ShowBox_Click
    Call WriteDebugLog("Centering")
    CenterWindow Me
'    For X = 0 To 5
'        If PKMN(X + 1).No = 0 Then Species(X).ListIndex = 0
'    Next
    
    Call WriteDebugLog("Setting up Mod")
    With mnuVersionItem(9)
        If Len(DBModStr) = 0 Then
            .Caption = "No Mod Loaded"
            .Enabled = False
        Else
            .Enabled = True
            If Len(DBModName) = 0 Then
                .Caption = "Mod: [No Name]"
            Else
                .Caption = "Mod: " & DBModName
            End If
        End If
    End With
        
    
        


    LoadingForm = False
    Call DoRBY
    Call RefreshRecentFiles
    
    Me.SetFocus
    Call WriteDebugLog("Finished")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim CantQuit As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Answer As Integer
    Dim FileToUse As String
    Dim Temp As String
    Dim SaveChange As Boolean
    Call ApplySettings
    For X = 1 To 6
        If PKMN(X).Nickname = "" Then PKMN(X).Nickname = PKMN(X).Name
    Next X
    Call RefreshCurrMoves
    CheckTeam = CreateTeamString
    SaveChange = Not (StrComp(SaveTeam, CheckTeam, vbBinaryCompare) = 0)
    If UnloadMode > 1 Then
        If IDiedOnceAlready Then Exit Sub
        IDiedOnceAlready = True
        ShuttingDown = True
        Answer = MsgBox("Do you want to save your team before exiting?", vbYesNo + vbQuestion, "Save changes?")
        If Answer = vbYes Then Call Command3_Click
        Exit Sub
    End If
    If UnloadMode <= 1 And SaveChange Then
        If StoredFileName = "" Then
            Answer = MsgBox("Do you want to save the changes to your team?", vbDefaultButton3 + vbYesNoCancel + vbExclamation)
        Else
            Temp = Chr$(34) & Right$(StoredFileName, Len(StoredFileName) - InStrRev(StoredFileName, "\")) & Chr$(34)
            Answer = MsgBox("Do you want to save the changes to " & Temp & "?", vbDefaultButton3 + vbYesNoCancel + vbExclamation)
        End If
        Select Case Answer
        Case vbCancel
            Cancel = True
            Exit Sub
        Case vbYes
            SkipDialog = True
            Call Command3_Click
            SkipDialog = False
        End Select
    End If
    Call ApplySettings
    Temp = BattleOK
    If TeamChangeFromMS Then
        If Temp <> "" Then
            MsgBox Temp & "  While connected to a server, you cannot exit the Team Builder without a valid team.", vbCritical, "Invalid Team"
            Cancel = True
        End If
    Else
        If Temp <> "" Then
            If MsgBox(Temp & "  Exit anyway?", vbYesNo, "Not Ready") = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
'    For X = 1 To 6
'        PKMN(X).Nickname = Replace(PKMN(X).Nickname, ",", "")
'        PKMN(X).Nickname = Replace(PKMN(X).Nickname, ":", "")
'        PKMN(X).Nickname = Replace(PKMN(X).Nickname, "'", "")
'    Next X
'    You.Name = Replace(You.Name, ",", "")
'    You.Name = Replace(You.Name, ":", "")
'    'You.Name = Replace(You.Name, "'", "")
Cancelled:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim X As Integer
    Dim OrigChange As Boolean
    Dim UserChange As Boolean
    On Error Resume Next
    If ShuttingDown Then
        Call WriteBoxPKMN
        Exit Sub
    End If
    If DontUpdate Then
        If TeamChangeFromMS Then MasterServer.TeamChanged = False
    Else
        Call ApplySettings
        For X = 1 To 6
            If PKMN(X).Nickname = "" Then PKMN(X).Nickname = PKMN(X).Name
            'If StoredPKMN(X).Nickname <> PKMN(X).Nickname Then ChangeOK = True
            'If StoredPKMN(X).Item <> PKMN(X).Item Then ChangeOK = True
        Next X
        
        For X = 1 To 6
            StoredPKMN(X) = PKMN(X)
        Next X
        
        If TeamChangeFromMS Then
            MasterServer.TeamChanged = Not (StrComp(OrigTeam, CheckTeam, vbBinaryCompare) = 0)
            MasterServer.TBUserChange = Not (StrComp(Left(OrigTeam, 702), Left(CheckTeam, 702), vbBinaryCompare) = 0)
        Else
            Call Loader.RefreshBattleButtons
        End If
    End If
    Ranking = TeamRank
    Call WriteBoxPKMN
    SaveSetting "NetBattle", "Options", "LastTBMode", TBMode
    On Error Resume Next
    If TeamChangeFromMS Then
        'MasterServer.Enabled = True
        Call MasterServer.mnuTeamItem_Click(3)
    Else
        Call Loader.RefreshRecentFiles
        Loader.Visible = True
    End If
    If SendBack Then Call MasterServer.SendData("BACK:")
End Sub

Private Sub MasterItem_Click()
    If LoadingForm Then Exit Sub
    PKMN(mintCurFrame - 1).Item = GetItemNum(MasterItem.List(MasterItem.ListIndex))
    Call RefreshCompatTree
End Sub

'Private Sub MasterItem_KeyUp(KeyCode As Integer, Shift As Integer)
'    If LoadingForm Then Exit Sub
'    PKMN(mintCurFrame - 1).Item = GetItemNum(MasterItem.List(MasterItem.ListIndex))
'    Call RefreshCompatTree
'End Sub
'
'Private Sub MasterItem_LostFocus()
'    If LoadingForm Then Exit Sub
'    PKMN(mintCurFrame - 1).Item = GetItemNum(MasterItem.List(MasterItem.ListIndex))
'    Call RefreshCompatTree
'End Sub

Private Sub LoseMSG_Change()
    If LoadingForm Then Exit Sub
    LoseMSG.Text = CorrectText(LoseMSG.Text)
    You.LoseMess = LoseMSG.Text
End Sub


Private Sub mnuDatadexItem_Click(Index As Integer)

End Sub

Private Sub MasterItem_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    F = mintCurFrame - 2
    If KeyAscii = 13 Then
        If MasterItem.ListIndex = -1 Then Exit Sub
        MovePick(F).SetFocus
        Exit Sub
    End If
    Temp = FutureText(MasterItem, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With MasterItem
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call MasterItem_Click
                .Text = .List(X)
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

Private Sub MasterItem_LostFocus()
    With MasterItem
        If .ListIndex = -1 Then .ListIndex = 0
        .Text = .List(.ListIndex)
    End With
    Call MasterItem_Click
End Sub

Private Sub MasterSpecies_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    F = mintCurFrame - 2
    If KeyAscii = 13 Then
        If MasterSpecies.ListIndex = -1 Then Exit Sub
        If MasterSpecies.List(MasterSpecies.ListIndex) <> MasterSpecies.Text Then
            For X = 0 To MasterSpecies.ListCount - 1
                If MasterSpecies.List(X) = MasterSpecies.Text Then
                    MasterSpecies.ListIndex = X
                    Exit For
                End If
            Next X
            If X = MasterSpecies.ListCount Then Exit Sub
        End If
        If Rebuild(F).Enabled Then Call Rebuild_Click(F)
        Nickname(F).SetFocus
        Nickname(F).SelLength = Len(Nickname(F).Text)
        Exit Sub
'    ElseIf KeyAscii = 22 Then
'        Temp = MasterSpecies.Text
    ElseIf KeyAscii < 32 Then
        Exit Sub
    Else
        Temp = FutureText(MasterSpecies, KeyAscii)
        If Temp = "" Then Exit Sub
        KeyAscii = 0
    End If
    B = False
    With MasterSpecies
        Y = Len(Temp)
        For X = 0 To .ListCount - 1
            If LCase(Left(.List(X), Y)) = LCase(Temp) Then
                .ListIndex = X
                Call MasterSpecies_Click
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

Private Sub MasterSpecies_LostFocus()
    With MasterSpecies
        If .ListIndex = -1 Then .Text = "" Else .Text = .List(.ListIndex)
    End With
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileToUse As String
    Dim BlankTrainer As Trainer
    Dim Answer As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Byte
    Dim H1 As Long
    Dim H2 As Long
    Dim Temp As String
    Select Case Index
        Case 0
            For X = 1 To 6
                PKMN(X) = BlankPKMN
                Changed(X) = True
            Next
            You = BlankTrainer
            If BetaRel <> "" Then
                You.ProgVersion = App.Major & "." & App.Minor & "." & BetaRel
            Else
                You.ProgVersion = App.Major & "." & App.Minor & "." & App.Revision
            End If
            Call LoadSettings
            StoredFileName = ""
            mnuFileItem(2).Enabled = False
            SaveTeam = CreateTeamString
        Case 1
            Call Command2_Click
        Case 2
            SkipDialog = True
            Call Command3_Click
            SkipDialog = False
        Case 3
            Call Command3_Click
        Case 4
            If MsgBox("Revert your team to the way it was when the Team Builder opened?", vbQuestion + vbYesNo, "Team Revert") = vbYes Then
                You = SwapTrainer
                For X = 1 To 6
                    PKMN(X) = StoredPKMN(X)
                    Changed(X) = True
                Next X
                OrigTeam = CreateTeamString
                SaveTeam = OrigTeam
                Call LoadSettings
                StoredFileName = OrigFilename
                If StoredFileName <> "" And UCase(Right(StoredFileName, 4)) = ".PNB" Then
                    mnuFileItem(2).Enabled = True
                Else
                    mnuFileItem(2).Enabled = False
                End If
            End If
        Case 5
            On Error GoTo ETrap
            Temp = MakeTeamText(PKMN)
            H1 = GetWinHandle(Shell("Notepad.exe", vbNormalFocus))
            H2 = FindWindowEx(H1, 0&, "Edit", vbNullString)
            SendMessageString H1, WM_SETTEXT, 256, "NetBattle Team: " & You.Name
            SendMessageString H2, WM_SETTEXT, 256, Temp
        Case 7
            Unload Me
        Case 9 To 12
            RecentLoad = RecentFiles(Index - 8)
            Call Command2_Click
    End Select
ETrap:
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute 0, vbNullString, "http://www.netbattle.net", vbNullString, vbNullString, 0
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Private Sub mnuMarkerItem_Click(Index As Integer)
    Dim X As Integer
    Dim Y As Integer
    X = Val(PokeBox.SelectedItem.Tag)
    If mnuMarkerItem(Index).Checked Then
        BoxPKMN(X).MarkerNum = BoxPKMN(X).MarkerNum - 2 ^ Index
    Else
        BoxPKMN(X).MarkerNum = BoxPKMN(X).MarkerNum + 2 ^ Index
    End If
    mnuMarkerItem(Index).Checked = Not mnuMarkerItem(Index).Checked
    Call RefreshBoxMinor
End Sub

Private Sub mnuSortItem_Click(Index As Integer)
    TBSort = Index
    SaveSetting "NetBattle", "Options", "Team Builder Sort", TBSort
    mnuSortItem(0).Checked = False
    mnuSortItem(1).Checked = False
    mnuSortItem(2).Checked = False
    mnuSortItem(3).Checked = False
    mnuSortItem(TBSort).Checked = True
    Call FillSpeciesList
End Sub

Private Sub mnuVersionItem_Click(Index As Integer)
    Dim Answer As Long
    Dim X As Long
    Dim Y As Long
    Dim Temp As String
    Dim Build() As String
    Dim FinalBuild As String
    If Index = 11 Then
        With MainContainer.FileBox
            .DialogTitle = "Load Database Mod"
            .Flags = cdlOFNHideReadOnly
            .CancelError = True
            .Filter = "NetBattle Database Mod (*.mod)|*.mod"
            .DefaultExt = ".mod"
            .FileName = ""
            If Len(Dir(SlashPath & "Database Mods\", vbDirectory)) > 0 Then
                .InitDir = SlashPath & "Database Mods\"
            Else
                .InitDir = SlashPath
            End If
            On Error GoTo ETrap
            .ShowOpen
            LoadDBMod .FileName
            mnuVersionItem(9).Caption = "DB Mod: " & DBModName
            mnuVersionItem(9).Enabled = True
            Exit Sub
        End With
    ElseIf Index = 10 Then
        If Len(DBModStr) = 0 Then Exit Sub
        ReDim Build(1 To UBound(BasePKMN), 1 To 3)
        SetSourceStringASM StrPtr(DBModStr)
        Do While GetBitsLeftASM > 2
            Select Case StreamOutASM(2)
            Case 0
                Exit Do
            Case 1
                X = StreamOutASM(9)
                Build(X, 1) = Build(X, 1) & "  Added move " & Moves(StreamOutASM(9)).Name & "." & vbNewLine
            Case 2
                X = StreamOutASM(9)
                Y = StreamOutASM(7)
                Build(X, 2) = Build(X, 2) & "  Changed Trait in slot " & CStr(StreamOutASM(1) + 1) & " to " & AttributeText(Y) & "." & vbNewLine
            Case 3
                X = StreamOutASM(9)
                Y = StreamOutASM(2)
                If Y = 0 Then
                    Build(X, 3) = Build(X, 3) & "  Cannot have move " & Moves(StreamOutASM(9)).Name & "." & vbNewLine
                Else
                    Build(X, 3) = Build(X, 3) & "  Cannot have moves " & Moves(StreamOutASM(9)).Name
                    For Y = 1 To Y - 1
                        Build(Y, 3) = Build(X, 3) & ", " & Moves(StreamOutASM(9)).Name
                    Next Y
                    If Y > 1 Then Build(X, 3) = Build(X, 3) & ","
                    Build(X, 3) = Build(X, 3) & " and " & Moves(StreamOutASM(9)).Name & " at the same time." & vbNewLine
                End If
            End Select
        Loop
        
        FinalBuild = vbNullString
        For X = 1 To UBound(BasePKMN)
            If Len(Build(X, 1)) > 0 Or Len(Build(X, 2)) > 0 Or Len(Build(X, 3)) > 0 Then
                FinalBuild = FinalBuild & BasePKMN(X).Name & ":" & vbNewLine & Build(X, 1) & Build(X, 2) & Build(X, 3)
            End If
        Next X
        
        X = GetWinHandle(Shell("Notepad.exe", vbNormalFocus))
        Y = FindWindowEx(X, 0&, "Edit", vbNullString)
        SendMessageString X, WM_SETTEXT, 256, "Currently Loaded Database Mods"
        SendMessageString Y, WM_SETTEXT, 256, FinalBuild
    End If

    
    If mnuVersionItem(Index).Checked = True Then Exit Sub
    'If TBMode = nbModAdv Then RestoreDB
    Y = TBMode
    Select Case Index
        Case 0
            TBMode = nbTrueRBY
        Case 1
            TBMode = nbRBYTrade
        Case 3
            TBMode = nbTrueGSC
        Case 4
            TBMode = nbGSCTrade
        Case 6
            TBMode = nbTrueRuSa
        Case 7
            TBMode = nbFullAdvance
        Case 9
            TBMode = nbModAdv
            'ApplyDBMod
    End Select
    If Not Compatibility(TBMode) Then
        Answer = MsgBox("Your current team is not compatible with the selected mode.  Some Pokemon, moves, or items will be reset.  Would you like to continue?", vbYesNo + vbQuestion, "Version Change")
        If Answer = vbNo Then
            TBMode = Y
            Exit Sub
        End If
    End If
    For X = 0 To 9
        mnuVersionItem(X).Checked = (X = Index)
    Next X
    For X = 1 To 6
        Changed(X) = True
    Next X
    Call DoVerChange(TBMode, Y)
    Call LoadSettings
ETrap:
End Sub

'Private Sub MoveList_Click()
'    Dim X As Integer
'    Dim MoveID As Integer
'    On Error GoTo Failed
'    MoveID = PKMN(TeamList.ListIndex + 1).Move(MoveList.ListIndex + 1)
'    With Moves(MoveID)
'        If .RBYMove Then
'            MoveAdvanced.Caption = .Name & " was added in Red/Blue/Yellow."
'        ElseIf .GSCMove Then
'            MoveAdvanced.Caption = .Name & " was added in Gold/Silver/Crystal."
'        Else
'            MoveAdvanced.Caption = .Name & " was added in Ruby/Sapphire."
'        End If
'    End With
'    With PKMN(TeamList.ListIndex + 1)
'        For X = 1 To UBound(.RBYMoves)
'            If Abs(.RBYMoves(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by R/B/Y Level"
'            End If
'        Next
'        For X = 1 To UBound(.RBYTM)
'            If Abs(.RBYTM(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by R/B/Y " & Moves(MoveID).OldTM
'            End If
'        Next
'        For X = 1 To UBound(.BaseMoves)
'            If Abs(.BaseMoves(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by G/S/C Level"
'            End If
'        Next
'        For X = 1 To UBound(.MachineMoves)
'            If Abs(.MachineMoves(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by G/S/C " & Moves(MoveID).NewTM
'            End If
'        Next
'        For X = 1 To UBound(.BreedingMoves)
'            If Abs(.BreedingMoves(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by G/S/C Breeding"
'            End If
'        Next
'        For X = 1 To UBound(.SpecialMoves)
'            If Abs(.SpecialMoves(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by Special (Stadium, Pokecenter)"
'            End If
'        Next
'        For X = 1 To UBound(.MoveTutor)
'            If Abs(.MoveTutor(X)) = MoveID Then
'             MoveAdvanced.Caption = MoveAdvanced.Caption & vbCrLf & "This is learned by G/S/C Move Tutor"
'            End If
'        Next
'    End With
'Failed:
'End Sub

Private Sub MovePick_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim Numbers As Boolean
    Select Case ColumnHeader.Index
    Case 2, 3, 4
        Numbers = True
    End Select
    If MovePick(Index).SortKey = ColumnHeader.Index - 1 Then
        If MovePick(Index).SortOrder = lvwAscending Then MovePick(Index).SortOrder = lvwDescending Else MovePick(Index).SortOrder = lvwAscending
    Else
        MovePick(Index).SortKey = ColumnHeader.Index - 1
        MovePick(Index).SortOrder = IIf(Numbers, lvwDescending, lvwAscending)
    End If
    If Numbers Then
        Call ListViewNumberSort(MovePick(Index), ColumnHeader.Index)
    Else
        MovePick(Index).Sorted = True
        MovePick(Index).Sorted = False
    End If
End Sub

Private Sub MovePick_DblClick(Index As Integer)
    On Error Resume Next
    Call MovePick_ItemClick(Index, MovePick(Index).SelectedItem)
End Sub

Private Sub MovePick_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    Dim M As Integer
    Dim StringToUse As String

    For X = 1 To UBound(Moves)
        If Moves(X).Name = Item.Text Then M = X
    Next
    StatusBar1.Panels(1).Text = Moves(M).Text
End Sub

Private Sub Nickname_Change(Index As Integer)
    If LoadingForm Then Exit Sub
    Nickname(Index).Text = CorrectText(Nickname(Index).Text)
    PKMN(Index + 1).Nickname = Nickname(Index).Text
End Sub

Private Sub Nickname_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
         KeyAscii = 0
         If MasterItem.Enabled Then
             MasterItem.SetFocus
             MasterItem.SelLength = Len(MasterItem.Text)
         Else
             MovePick(mintCurFrame - 2).SetFocus
         End If
     End If
End Sub

Private Sub Nickname_LostFocus(Index As Integer)
    If Nickname(Index).Text = "" And PKMN(Index + 1).No > 0 Then
        Nickname(Index).Text = PKMN(Index + 1).Name
    End If
End Sub

Private Sub PKMNPic_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PKMNPicBox_OLEDragDrop(Index, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PKMNPic_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Call PKMNPicBox_OLEDragOver(Index, Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub PKMNPicBox_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PKMNPicBox_OLEDragOver(Index, Data, Effect, Button, Shift, X, Y, 0)
    If Effect <> vbDropEffectCopy Then Exit Sub
    Call BoxNav_Click(0)
End Sub

Private Sub PKMNPicBox_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim Temp As String
    Temp = Data.GetData(1)
    If ChopString(Temp, 10) <> "PNBBOXDRAG" Then
        Effect = vbDropEffectNone
        Exit Sub
    End If
    If Val(Temp) <> PokeBox.SelectedItem.Tag Then
        Effect = vbDropEffectNone
        Exit Sub
    End If
    Effect = vbDropEffectCopy
End Sub

Private Sub PokeBox_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call PokeBoxClick(Item)
End Sub
Private Sub PokeBoxClick(PokeItem As ListItem)
    Dim X As Byte
    Dim Z As Byte
    Dim ThisPKMN As Integer
    Dim Vis As Boolean
    ThisPKMN = -1
    On Error Resume Next
    ThisPKMN = Val(PokeItem.Tag)
    If ThisPKMN = CurrentDisplay Then Exit Sub
    For X = 0 To 4
        lblMarker(X).Visible = False
    Next
    Vis = (ThisPKMN <> -1)
    If Vis Then
        With BoxPKMN(ThisPKMN)
            Call MainContainer.DoPicture(ChooseImage(BoxPKMN(ThisPKMN), nbGFXSml))
            imgInfo.Picture = MainContainer.SwapSpace.Picture
'            GenderImage(1).Visible = (.Gender = 1)
'            GenderImage(2).Visible = (.Gender = 2)
            lblInfo(0).Caption = .Nickname
            lblInfo(1).Caption = "Lv." & .Level & " " & .Name
            If .Gender = 1 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (M)"
            ElseIf .Gender = 2 Then
                lblInfo(1).Caption = lblInfo(1).Caption & " (F)"
            End If
            lblStat(0).Caption = .MaxHP
            lblStat(1).Caption = .Attack
            lblStat(2).Caption = .Defense
            lblStat(3).Caption = .Speed
            lblStat(4).Caption = .SpecialAttack
            lblStat(5).Caption = .SpecialDefense
            For X = 8 To 11
                lblInfo(X).Caption = Moves(.Move(X - 7)).Name
            Next X
            lblInfo(12).Caption = "Held Item: " & Item(.Item)
            
            Select Case .GameVersion
            Case 0: imgBoxVer.Picture = LoadResPicture("RBYT", vbResIcon)
            Case 1: imgBoxVer.Picture = LoadResPicture("GSCT", vbResIcon)
            Case 2: imgBoxVer.Picture = LoadResPicture("ADV", vbResIcon)
            Case 3: imgBoxVer.Picture = LoadResPicture("ADV+", vbResIcon)
            Case 4: imgBoxVer.Picture = LoadResPicture("MOD", vbResIcon)
            Case 5: imgBoxVer.Picture = LoadResPicture("RBY", vbResIcon)
            Case 6: imgBoxVer.Picture = LoadResPicture("GSC", vbResIcon)
            End Select
        
            For Z = 0 To 3
                lblMarker(Z).Visible = ((BoxPKMN(ThisPKMN).MarkerNum And 2 ^ Z) > 0)
                mnuMarkerItem(Z).Checked = lblMarker(Z).Visible
            Next Z

        End With
        BoxNav(2).Enabled = (PokeBox.SelectedItem.Index <> 1)
        BoxNav(3).Enabled = (PokeBox.SelectedItem.Index <> PokeBox.ListItems.count)

    End If
    imgInfo.Visible = Vis
    For X = 0 To 12
        lblInfo(X).Visible = Vis
    Next X
    For X = 0 To 5
        lblStat(X).Visible = Vis
    Next X
    For X = 0 To 2
        InfoLine(X).Visible = Vis
    Next X
    CurrentDisplay = ThisPKMN
End Sub

Private Sub PokeBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ETrap
    If Button = vbRightButton Then
        PokeBox.SelectedItem.EnsureVisible
        Me.PopupMenu mnuMarker
    End If
ETrap:
End Sub


Private Sub PokeBox_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    On Error GoTo NoSel
    With PokeBox.SelectedItem
        Data.SetData "PNBBOXDRAG" & .Tag, 1
    End With
NoSel:
End Sub

Private Sub Rebuild_Click(Index As Integer)
    Dim Duplicate As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim AttackVar As Integer
    Dim DefenseVar As Integer
    Dim SpeedVar As Integer
    Dim SpecialVar As Integer
    Dim Unown As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim i As Integer
    Dim Temp As Integer
    Dim NewSpecies As Integer
    
    If MasterSpecies.ListIndex = -1 Then Exit Sub
    
    For X = 1 To UBound(BasePKMN)
        If BasePKMN(X).Name = MasterSpecies.List(MasterSpecies.ListIndex) Then NewSpecies = X
    Next
    
    Duplicate = False
    Select Case NewSpecies
        Case 386, 387, 388, 389
            For X = 1 To 6
                If (PKMN(X).No = 386 Or PKMN(X).No = 387 Or PKMN(X).No = 388 Or PKMN(X).No = 389) And X <> Index + 1 And PKMN(X).No <> NewSpecies Then
                    MsgBox "You already have a " & BasePKMN(386).Name & " on your team!", vbCritical, "Duplicate Pokémon"
                    Exit Sub
                End If
            Next
'        Case Else
'            For X = 1 To 6
'                If PKMN(X).No = NewSpecies And X <> Index + 1 Then Duplicate = True
'            Next
    End Select
    If Duplicate Then
        MsgBox "You already have a " & BasePKMN(NewSpecies).Name & " on your team!", vbCritical, "Duplicate Pokémon"
        Exit Sub
    End If
    
    Changed(Index + 1) = True
    'MovePick(Index).ListItems.Clear
    StatusBar1.Panels(1).Text = ""
    
    'Get move list ready
    'Call PrepareMode(TBMode, NewSpecies)
    
    'Copy base values to player Pokemon
    PKMN(Index + 1) = BasePKMN(NewSpecies)
    
    'Set nickname
    Nickname(Index).Text = BasePKMN(NewSpecies).Name
    PKMN(Index + 1).Nickname = BasePKMN(NewSpecies).Name
    
    With PKMN(Index + 1)
        .GameVersion = TBMode
        
        Select Case TBMode
        Case 0, 5, 6, 1
            'Generate DVs & level
            .DV_Atk = 15
            .DV_Def = 15
            .DV_Spd = 15
            .DV_SAtk = 15
            .DV_SDef = 0
            .EV_Atk = 0
            .EV_Def = 0
            .EV_Spd = 0
            .EV_SAtk = 0
            .EV_SDef = 0
            .EV_HP = 0
            .Level = 100
            
            'Adjust stats for DVs
            .Attack = GetStat(.Level, .BaseAttack, .DV_Atk)
            .Defense = GetStat(.Level, .BaseDefense, .DV_Def)
            .Speed = GetStat(.Level, .BaseSpeed, .DV_Spd)
            Select Case TBMode
            Case 0, 5
                .SpecialAttack = GetStat(.Level, .BaseSpecial, .DV_SAtk)
                .SpecialDefense = GetStat(.Level, .BaseSpecial, .DV_SAtk)
            Case 1, 6
                .SpecialAttack = GetStat(.Level, .BaseSAttack, .DV_SAtk)
                .SpecialDefense = GetStat(.Level, .BaseSDefense, .DV_SAtk)
            End Select
            
            'Generate HP DV based on the others
            .DV_HP = 0
            If .DV_Atk Mod 2 = 1 Then .DV_HP = .DV_HP + 8
            If .DV_Def Mod 2 = 1 Then .DV_HP = .DV_HP + 4
            If .DV_Spd Mod 2 = 1 Then .DV_HP = .DV_HP + 2
            If .DV_SAtk Mod 2 = 1 Then .DV_HP = .DV_HP + 1
            .MaxHP = GetHP(.Level, .BaseHP, .DV_HP)
            
            'Set gender
            If .PercentFemale = -1 Then
                .Gender = 0
            Else
                If .DV_Atk <= .PercentFemale - 1 Then .Gender = 2 Else .Gender = 1
            End If
            
        Case Else
            .DV_HP = 31
            .DV_Atk = 31
            .DV_Def = 31
            .DV_Spd = 31
            .DV_SAtk = 31
            .DV_SDef = 31
            .EV_Atk = 85
            .EV_Def = 85
            .EV_Spd = 85
            .EV_SAtk = 85
            .EV_SDef = 85
            .EV_HP = 85
            .Level = 100
            .NatureNum = 0
            .AttNum = 0
            
            .MaxHP = GetAdvHP(.BaseHP, 31, 85, 100)
            .Attack = GetAdvStat(.BaseAttack, 31, 85, 100, 0)
            .Defense = GetAdvStat(.BaseDefense, 31, 85, 100, 0)
            .Speed = GetAdvStat(.BaseSpeed, 31, 85, 100, 0)
            .SpecialAttack = GetAdvStat(.BaseSAttack, 31, 85, 100, 0)
            .SpecialDefense = GetAdvStat(.BaseSDefense, 31, 85, 100, 0)
            
            Select Case .PercentFemale
            Case -1: .Gender = 0
            Case 16: .Gender = 2
            Case Else: .Gender = 1
            End Select
        End Select
    
        Call RefreshImage(Index + 1, True)
        
        'Adjust on-screen values
        If TBMode = nbRBYTrade Or TBMode = nbTrueRBY Then
            GenderDisp(Index).Caption = "Lv." & .Level
        Else
            GenderDisp(Index).Caption = Gender(.Gender) & " Lv." & .Level
        End If
        HP(Index) = .MaxHP
        Attack(Index) = .Attack
        Defense(Index) = .Defense
        Speed(Index) = .Speed
        SpecialAttack(Index) = .SpecialAttack
        SpecialDefense(Index) = .SpecialDefense
        Type1(Index).Caption = Element(.Type1)
        If (TBMode = nbTrueRBY Or TBMode = nbRBYTrade) And (.No = 81 Or .No = 82) Then
            Type2(Index).Caption = ""
        Else
            Type2(Index).Caption = Element(.Type2)
        End If
'        If .Legendary Then
'            StatWarn(Index).Caption = "Legendary"
'        ElseIf .Uber Then
'            StatWarn(Index).Caption = "Uber"
'        Else
'            StatWarn(Index).Caption = ""
'        End If
        Call RefreshCurrMoves

        Call FillInMoveList(Index)
        .Item = GetItemNum(MasterItem.List(MasterItem.ListIndex))
    End With
    Rebuild(Index).Enabled = False
    MainTabs.Tabs(Index + 2).Image = 1
    Call DoRBY
    Call DoPower
    'ExpertBuild(Index).Enabled = True
End Sub

Sub ApplySettings()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim M As Integer
    Dim Temp As Integer
    
    You.Name = UserName.Text
    For X = 1 To TrainerPics.ListItems.count
        If TrainerPics.ListItems(X).Selected Then You.Picture = X
    Next X
    You.Version = VersionSelect.ListIndex
    You.Extra = ExtraInfo.Text
    You.WinMess = WinMSG.Text
    You.LoseMess = LoseMSG.Text
    For X = 0 To 5
        Call SetPokeMoves(X)
    Next
End Sub

Public Sub LoadSettings()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim i As Integer
    Dim Temp As Integer
    Dim Temp2 As String
    Dim HasMoves As Boolean
    Dim MaxMoves As Integer
    
    'If You.Name = "" Then Exit Sub
    MainContainer.MousePointer = vbHourglass
    SetRedraw Me.hWnd, False
    For X = 0 To 5
        MainTabs.Tabs(X + 2).Image = 1
        CurrMoves(X).Caption = ""
        'MovePick(x).Enabled = False
        BoxNav(X).Enabled = False
        Nickname(X).Enabled = False
    Next X
    MasterSpecies.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    UserName.Text = You.Name
    RefreshCurrMoves
    If You.Name <> "" Then MainTabs.Tabs(1).Image = 4 Else MainTabs.Tabs(1).Image = 3
    If You.Picture > 0 Then Set TrainerPics.SelectedItem = TrainerPics.ListItems(You.Picture)
    VersionSelect.ListIndex = You.Version
    ExtraInfo.Text = You.Extra
    WinMSG.Text = You.WinMess
    LoseMSG.Text = You.LoseMess
    CompatTree.Nodes.Clear
    Call RefreshRecentFiles
    Call RedrawTB
    Call WriteDebugLog("Loading Blank Slots")
    For X = 1 To 6
        If Not Changed(X) Then
            MaxMoves = MovePick(X - 1).ListItems.count
            If MaxMoves > 4 Then MaxMoves = 4
            For Y = 1 To MaxMoves
                If PKMN(X).Move(Y) = 0 Then Exit For
            Next
            If Y = MaxMoves + 1 Then MainTabs.Tabs(X + 1).Image = 2 Else MainTabs.Tabs(X + 1).Image = 1
        End If
    Next X
    Call RefreshCompatTree
    For X = 1 To 6
        If PKMN(X).No = 0 And Changed(X) Then
            Call ResetSlot(X)
        ElseIf PKMN(X).No > 0 And Changed(X) Then
            With PKMN(X)
                Select Case TBMode
                Case 2, 3
                    If Not (.GameVersion = 2 Or .GameVersion = 3) Then
                        .DV_HP = 31
                        .DV_Atk = 31
                        .DV_Def = 31
                        .DV_Spd = 31
                        .DV_SAtk = 31
                        .DV_SDef = 31
                        .EV_Atk = 85
                        .EV_Def = 85
                        .EV_Spd = 85
                        .EV_SAtk = 85
                        .EV_SDef = 85
                        .EV_HP = 85
                        .NatureNum = 0
                        .AttNum = 0
                    End If
                Case Else
                    If .GameVersion = 2 Or .GameVersion = 3 Then
                        .DV_Atk = 15
                        .DV_Def = 15
                        .DV_Spd = 15
                        .DV_SAtk = 15
                        .DV_SDef = 0
                        .EV_Atk = 0
                        .EV_Def = 0
                        .EV_Spd = 0
                        .EV_SAtk = 0
                        .EV_SDef = 0
                        .EV_HP = 0
                    End If
                End Select
                Call FillInPokeData(PKMN(X), TBMode)
                Call RefreshImage(X, False)
                If TBMode = nbTrueRBY Or TBMode = nbRBYTrade Then
                    GenderDisp(X - 1).Caption = "Lv." & .Level
                Else
                    GenderDisp(X - 1).Caption = Gender(.Gender) & " Lv." & .Level
                End If
                'Species(X - 1).Text = .Name
                Nickname(X - 1).Text = .Nickname
                HP(X - 1) = .MaxHP
                Attack(X - 1) = .Attack
                Defense(X - 1) = .Defense
                Speed(X - 1) = .Speed
                SpecialAttack(X - 1) = .SpecialAttack
                SpecialDefense(X - 1) = .SpecialDefense
                Type1(X - 1).Caption = Element(.Type1)
                If (TBMode = nbTrueRBY Or TBMode = nbRBYTrade) And (.No = 81 Or .No = 82) Then
                    Type2(X - 1).Caption = ""
                Else
                    Type2(X - 1).Caption = Element(.Type2)
                End If
                Call FillInMoveList(X - 1)
                SetPokeMoves (X - 1)
                HasMoves = False
                For Y = 1 To 4
                    If .Move(Y) = 0 Then Exit For
                Next Y
                HasMoves = (Y = 5)
                If HasMoves Then MainTabs.Tabs(X + 1).Image = 2 Else MainTabs.Tabs(X + 1).Image = 1
            End With
            Changed(X) = False
        End If
    Next X
    Call WriteDebugLog("Doing power")
    Call DoRBY
    Call DoPower
    For X = 0 To 5
        BoxNav(X).Enabled = True
        Nickname(X).Enabled = True
    Next X
    MasterSpecies.Enabled = True
    If mintCurFrame > 1 And mintCurFrame < 8 Then
        If PKMN(mintCurFrame - 1).Name <> "" Then
            For Y = 0 To MasterSpecies.ListCount - 1
                If MasterSpecies.List(Y) = PKMN(mintCurFrame - 1).Name Then MasterSpecies.ListIndex = Y
            Next Y
        Else
            MasterSpecies.ListIndex = -1
        End If
        For Y = 0 To MasterItem.ListCount - 1
            If MasterItem.List(Y) = Item(PKMN(mintCurFrame - 1).Item) Then Exit For
        Next Y
        MasterItem.ListIndex = IIf(Y = MasterItem.ListCount, 0, Y)
    End If
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Call WriteDebugLog("Exiting Load Settings")
    SetRedraw Me.hWnd, True
    MainContainer.MousePointer = vbNormal
End Sub

Private Sub MovePick_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim M As Integer
    Dim N As Integer
    Dim MaxMoves As Integer
    Dim MoveArrayTemp() As Boolean
    Dim Temp As String
    Dim Temp2 As Integer
    Dim Temp3 As String
    Dim FromBox As Byte
    FromBox = MovePick(Index).Tag
    MovePick(Index).Tag = 0
    
    'Figure out how many are checked
    N = 0
    For X = 1 To MovePick(Index).ListItems.count
        If MovePick(Index).ListItems(X).Checked Then N = N + 1
    Next X

    'Error if you check 5, uncheck the 5th
    If N = 5 Then
        MsgBox "Pokémon can only use four moves at a time!", , "Error!"
        Item.Checked = False
        MovePick(Index).SetFocus
        Exit Sub
    End If

    'If the move is not working, pop up a message
    If Item.Ghosted = True And Item.Checked Then
        MsgBox "That move doesn't work right yet - you can pick it, but it's special ability might not work.", vbInformation, "Warning"
    End If

'    'Adjust the moves
'    Call SetPokeMoves(Index)
    If Item.Checked Then
        For X = 1 To 4
            If PKMN(Index + 1).Move(X) = 0 Then Exit For
        Next X
        If FromBox > 0 Then X = FromBox
        PKMN(Index + 1).Move(X) = Val(Right(Item.Key, 3))
    Else
        For X = 1 To 4
            If PKMN(Index + 1).Move(X) = Val(Right(Item.Key, 3)) Then Exit For
        Next X
        PKMN(Index + 1).Move(X) = 0
    End If
    
    'Legal Move Check
    X = Index + 1
    Temp = LegalMove(PKMN(X))
    If Temp <> "" Then
        MsgBox Temp, vbCritical, "Illegal Move"
        Item.Checked = False
        Call MovePick_ItemCheck(Index, Item)
        If FromBox = 0 Then
            MovePick(Index).SetFocus
        Else
            txtMove(FromBox).SetFocus
        End If
        Exit Sub
    End If
    
    'Hey, there's a faster way.  Don't need this part then...
    'MaxMoves = GetMoveCount(PKMN(Index + 1), TBMode)
    MaxMoves = MovePick(Index).ListItems.count
    If MaxMoves > 4 Then MaxMoves = 4
    For X = 1 To MaxMoves
        If PKMN(Index + 1).Move(X) = 0 Then Exit For
    Next
    If X = MaxMoves + 1 Then MainTabs.Tabs(Index + 2).Image = 2 Else MainTabs.Tabs(Index + 2).Image = 1
    RefreshCurrMoves
    Call DoRBY
End Sub

Sub SetPokeMoves(ByVal Index As Integer)
    Dim M As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Temp As String
    Dim Temp2 As String
    For Y = 1 To 4
        PKMN(Index + 1).Move(Y) = 0
    Next
    M = 1
    For Y = 1 To MovePick(Index).ListItems.count
        If MovePick(Index).ListItems(Y).Checked Then
            Temp = MovePick(Index).ListItems(Y).Text
            If PKMN(Index + 1).Move(4) <> 0 Then Exit Sub
            For X = 1 To 4
                M = PKMN(Index + 1).Move(X)
                If Temp < Moves(M).Name Or M = 0 Then Exit For
            Next X
            If M <> 0 Then
                For X = 4 To X + 1 Step -1
                    PKMN(Index + 1).Move(X) = PKMN(Index + 1).Move(X - 1)
                Next X
            End If
            PKMN(Index + 1).Move(X) = Val(Right(MovePick(Index).ListItems(Y).Key, 3))
        End If
    Next Y
    Call RefreshCurrMoves
End Sub

Sub FillInMoveList(ByVal Index As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim B As Boolean
    Dim Temp As String
    Dim TempMove() As Integer
    Dim TempSource() As String
    Dim TempItem As ListItem
    Dim ThisMove As Move
    On Error Resume Next
    
    If PKMN(Index + 1).No = 0 Then
        MovePick(Index).ListItems.Clear
        Exit Sub
    End If
    Call MakeMoveArray(PKMN(Index + 1).No, TBMode, TempMove, TempSource)
        
    MovePick(Index).ListItems.Clear
    For X = 1 To UBound(TempMove)
        Z = TempMove(X)
        If Z > 0 Then
            ThisMove = ConvertMove(Moves(Z), CompatVersion(TBMode))
            Set TempItem = MovePick(Index).ListItems.Add(, "#" & Format(Z, "000"), ThisMove.Name, ThisMove.Type, ThisMove.Type)
            With TempItem
                If ThisMove.Power > 0 Then .SubItems(1) = ThisMove.Power Else .SubItems(1) = "-"
                Y = ThisMove.Accuracy
                If Y = 0 Then Y = 100
                .SubItems(2) = CStr(Y) & "%"
                .SubItems(3) = ThisMove.PP
                .SubItems(4) = TempSource(X)
                .ToolTipText = ThisMove.Text
                .Ghosted = Not ThisMove.WorksRight
            End With
        End If
    Next X
    For X = 1 To 4
        If PKMN(Index + 1).Move(X) > 0 Then
            MovePick(Index).ListItems("#" & Format(PKMN(Index + 1).Move(X), "000")).Checked = True
        End If
    Next
    MovePick(Index).Sorted = True
    MovePick(Index).SortKey = 0
    MovePick(Index).SortOrder = lvwAscending
End Sub

Private Sub ShowBox_Click()
    ShowBoxes = Not ShowBoxes
    SaveSetting "NetBattle", "Team Builder", "Show Boxes", ShowBoxes
    Call DoBoxChange
    ShowBox.Refresh
    CenterWindow Me
End Sub

Private Sub MasterSpecies_Click()
    Dim X As Integer
    X = MasterSpecies.ListIndex
    If X >= 0 Then
        If MasterSpecies.List(X) <> PKMN(mintCurFrame - 1).Name Then
            Rebuild(mintCurFrame - 2).Enabled = True
        Else
            Rebuild(mintCurFrame - 2).Enabled = False
        End If
    End If
End Sub


'Private Sub MasterSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim X As Integer
'    X = MasterSpecies.ListIndex
'    If X >= 0 Then
'        If MasterSpecies.List(X) <> PKMN(mintCurFrame - 1).Name Then
'            Rebuild(mintCurFrame - 2).Enabled = True
'        Else
'            Rebuild(mintCurFrame - 2).Enabled = False
'        End If
'    End If
'End Sub
'
'Private Sub MasterSpecies_LostFocus()
'    Dim X As Integer
'    X = MasterSpecies.ListIndex
'    If X >= 0 Then
'        If MasterSpecies.List(X) <> PKMN(mintCurFrame - 1).Name Then
'            Rebuild(mintCurFrame - 2).Enabled = True
'        Else
'            Rebuild(mintCurFrame - 2).Enabled = False
'        End If
'    End If
'End Sub

Private Sub MainTabs_Click()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    X = MainTabs.SelectedItem.Index
    If X = mintCurFrame Then Exit Sub ' No need to change frame.
    'Otherwise, hide old frame, show new.
    If X = 8 Then
'        TeamList.ListIndex = 0
'        Call TeamList_Click
    End If
    StatusBar1.Panels(1).Text = ""
    Z = mintCurFrame
    mintCurFrame = MainTabs.SelectedItem.Index
    RefreshCurrMoves
    If X <> 1 And X <> 8 Then
        MasterSpecies.Visible = True
        MasterItem.Visible = True
        If PKMN(mintCurFrame - 1).Name <> "" Then
            For Y = 0 To MasterSpecies.ListCount - 1
                If MasterSpecies.List(Y) = PKMN(X - 1).Name Then MasterSpecies.ListIndex = Y
            Next Y
        Else
            MasterSpecies.ListIndex = -1
        End If
        For Y = 0 To MasterItem.ListCount - 1
            If MasterItem.List(Y) = Item(PKMN(X - 1).Item) Then Exit For
        Next Y
        MasterItem.ListIndex = IIf(Y = MasterItem.ListCount, 0, Y)
    Else
        MasterSpecies.Visible = False
        MasterItem.Visible = False
    End If
    TabHolder(X - 1).Visible = True
    TabHolder(Z - 1).Visible = False
    'Set mintCurFrame to new value.
End Sub

Private Sub MainTabs_KeyUp(KeyCode As Integer, Shift As Integer)
    If MainTabs.SelectedItem.Index = mintCurFrame Then Exit Sub ' No need to change frame.
   ' Otherwise, hide old frame, show new.
   TabHolder(MainTabs.SelectedItem.Index - 1).Visible = True
   TabHolder(mintCurFrame - 1).Visible = False
   ' Set mintCurFrame to new value.
   mintCurFrame = MainTabs.SelectedItem.Index
End Sub

Private Sub MainTabs_LostFocus()
    If MainTabs.SelectedItem.Index = mintCurFrame Then Exit Sub ' No need to change frame.
   ' Otherwise, hide old frame, show new.
   TabHolder(MainTabs.SelectedItem.Index - 1).Visible = True
   TabHolder(mintCurFrame - 1).Visible = False
   ' Set mintCurFrame to new value.
   mintCurFrame = MainTabs.SelectedItem.Index
End Sub

Private Sub MainTabs_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim HitTestInfo As TCHITTESTINFO
    Dim R As Long
    Dim Temp As String
    Dim Dummy As String
    Call MainTabs_OLEDragOver(Data, Effect, Button, Shift, X, Y, 0)
    'Debug.Print Effect
    If Effect <> vbDropEffectCopy Then Exit Sub
    Temp = Data.GetData(1)
    Dummy = ChopString(Temp, 10)
    HitTestInfo.pt.X = X / Screen.TwipsPerPixelX
    HitTestInfo.pt.Y = Y / Screen.TwipsPerPixelY
    R = SendMessage(MainTabs.hWnd, TCM_HITTEST, 0&, HitTestInfo)
    MainTabs.Tabs(R + 1).Selected = True
    Call BoxNav_Click(0)
    Effect = 3
End Sub

Private Sub MainTabs_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim HitTestInfo As TCHITTESTINFO
    Dim R As Long
    Dim Temp As String
    On Error GoTo NoSel
    Temp = Data.GetData(1)
    If ChopString(Temp, 10) <> "PNBBOXDRAG" Then
NoSel:
        Effect = vbDropEffectNone
        Exit Sub
    End If
    HitTestInfo.pt.X = X / Screen.TwipsPerPixelX
    HitTestInfo.pt.Y = Y / Screen.TwipsPerPixelY
    R = SendMessage(MainTabs.hWnd, TCM_HITTEST, 0&, HitTestInfo)
    If R < 1 Or R > 6 Or Val(Temp) <> PokeBox.SelectedItem.Tag Then
        Effect = vbDropEffectNone
        Exit Sub
    End If
    Effect = vbDropEffectCopy
End Sub

Private Sub CompatTree_Collapse(ByVal Node As MSComctlLib.Node)
    Call CompatTree_NodeClick(Node)
End Sub

Private Sub CompatTree_Expand(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
End Sub

Private Sub CompatTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim X As Integer
    Dim MoveID As Integer
    Dim Temp As String
    Dim NL As String
    Dim OneCompat(0 To 6) As Boolean
    NL = vbNewLine
    On Error GoTo Failed
    MoveAdvanced.Text = ""
    If Node.Key = "Name" Then
        Call ReadBinArray(CompatCheck(PKMN), OneCompat)
        For X = 0 To 6
            If X <> 4 Then OnePKMN(X).Visible = OneCompat(X)
        Next X
    Else
        X = Val(Mid(Node.Key, 5, 1))
        With PKMN(X)
            Call ReadBinArray(CompatCheck(PKMN, X), OneCompat)
            For X = 0 To 6
                If X <> 4 Then OnePKMN(X).Visible = OneCompat(X)
            Next
        End With
    End If
    
    Select Case Len(Node.Key)
    Case 5
        With PKMN(Val(Mid(Node.Key, 5, 1)))
            If .No > 0 Then
                If .ExistRBY Then
                    Temp = .Name & " was added in Red/Blue/Yellow."
                    If .ExistGSC Then
                        Temp = Temp & NL & "It is naturally obtainable in Gold/Silver/Crystal."
                    Else
                        Temp = Temp & NL & "It is obtainable in Gold/Silver/Crystal through Trade."
                    End If
                    If .ExistAdv Then
                        Temp = Temp & NL & "It is naturally obtainable in Ruby/Sapphire."
                    Else
                        Temp = Temp & NL & "It is obtainable in Ruby/Sapphire through GBA Trade."
                    End If
                ElseIf .ExistGSC Then
                    Temp = .Name & " was added in Gold/Silver/Crystal."
                    If .ExistAdv Then
                        Temp = Temp & NL & "It is naturally obtainable in Ruby/Sapphire."
                    Else
                        Temp = Temp & NL & "It is obtainable in Ruby/Sapphire through GBA Trade."
                    End If
                ElseIf .ExistAdv Then
                    Temp = .Name & " was added in Ruby/Sapphire."
                End If
            End If
        End With
    Case 10
        X = Val(Mid(Node.Key, 5, 1))
        MoveID = PKMN(X).Move(Val(Mid(Node.Key, 10, 1)))
        With Moves(MoveID)
            If .Name = "" Then Exit Sub
            If .RBYMove Then
                Temp = .Name & " was added in Red/Blue/Yellow."
            ElseIf .GSCMove Then
                Temp = .Name & " was added in Gold/Silver/Crystal."
            Else
                Temp = .Name & " was added in Ruby/Sapphire."
            End If
        End With
        With PKMN(X)
            For X = 1 To UBound(.RBYMoves)
                If .RBYMoves(X) = MoveID Then
                    Temp = Temp & NL & "This is learned by R/B/Y Level"
                End If
            Next
            For X = 1 To UBound(.RBYTM)
                If .RBYTM(X) = MoveID Then
                    Temp = Temp & NL & "This is learned by R/B/Y " & Moves(MoveID).OldTM
                End If
            Next
            For X = 1 To UBound(.BaseMoves)
                If Abs(.BaseMoves(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by G/S/C Level"
                End If
            Next
            For X = 1 To UBound(.MachineMoves)
                If Abs(.MachineMoves(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by G/S/C " & Moves(MoveID).NewTM
                End If
            Next
            For X = 1 To UBound(.BreedingMoves)
                If Abs(.BreedingMoves(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by G/S/C Breeding"
                End If
            Next
            For X = 1 To UBound(.SpecialMoves)
                If Abs(.SpecialMoves(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by G/S/C Special (Stadium, Pokecenter)"
                End If
            Next
            For X = 1 To UBound(.MoveTutor)
                If Abs(.MoveTutor(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by G/S/C Move Tutor"
                End If
            Next
            For X = 1 To UBound(.AdvMoves)
                If Abs(.AdvMoves(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Ruby/Sapp Level"
                End If
            Next
            For X = 1 To UBound(.ADVTM)
                If Abs(.ADVTM(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Ruby/Sapp " & Moves(MoveID).ADVTM
                End If
            Next
            For X = 1 To UBound(.AdvBreeding)
                If Abs(.AdvBreeding(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Ruby/Sapp Breeding"
                End If
            Next
            For X = 1 To UBound(.AdvSpecial)
                If Abs(.AdvSpecial(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Ruby/Sapp Special (Box, Pokecenter)"
                End If
            Next
            For X = 1 To UBound(.LFOnly)
                If Abs(.LFOnly(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Leaf/Fire Level or Relearner"
                End If
            Next
            For X = 1 To UBound(.AdvTutor)
                If Abs(.AdvTutor(X)) = MoveID Then
                    Temp = Temp & NL & "This is learned by Leaf/Fire Tutor"
                End If
            Next
        End With
    Case 9
        X = Val(Mid(Node.Key, 5, 1))
        If PKMN(X).Item = nbNoItem Then Exit Sub
        If PKMN(X).Item <= 41 Then
            Temp = Item(PKMN(X).Item) & " was added in Gold/Silver/Crystal."
            If AdvItem(PKMN(X).Item) Then
                Temp = Temp & NL & "It is also available in Ruby/Sapphire."
            Else
                Temp = Temp & NL & "It is not available in Ruby/Sapphire."
            End If
        Else
            Temp = Item(PKMN(X).Item) & " was added in Ruby/Sapphire."
        End If
    End Select
    MoveAdvanced.Text = Temp
    Exit Sub
        
Failed:

End Sub

Private Sub txtMove_Change(Index As Integer)
    With MovePick(mintCurFrame - 2)
        If txtMove(Index).Text = "" And txtMove(Index).Tag <> "" Then
            If .ListItems(txtMove(Index).Tag).Checked Then
                .ListItems(txtMove(Index).Tag).Checked = False
                Call MovePick_ItemCheck(.Index, .ListItems(txtMove(Index).Tag))
            End If
            txtMove(Index).Tag = ""
        End If
    End With
End Sub

Private Sub txtMove_GotFocus(Index As Integer)
    If txtMove(Index).SelStart = 0 Then
        txtMove(Index).SelLength = Len(txtMove(Index).Text)
    End If
    If txtMove(Index).Tag <> "" Then
        With MovePick(mintCurFrame - 2).ListItems(txtMove(Index).Tag)
            .Selected = True
            .EnsureVisible
        End With
        Call MovePick_ItemClick(mintCurFrame - 2, MovePick(mintCurFrame - 2).SelectedItem)
    End If
End Sub

Private Sub txtMove_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo ETrap
    Dim X As Long
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        With MovePick(mintCurFrame - 2)
            X = .SelectedItem.Index
            Do
                X = X + IIf(KeyCode = vbKeyUp, -1, 1)
            Loop Until Not .ListItems(X).Checked
            Set .SelectedItem = .ListItems(X)
            KeyCode = 0
            .SelectedItem.EnsureVisible
            txtMove(Index).Text = .SelectedItem.Text
            Call txtMove_KeyPress(Index, 0)
        End With
        Call MovePick_ItemClick(mintCurFrame - 2, MovePick(mintCurFrame - 2).SelectedItem)
    End If
ETrap:
End Sub

Private Sub txtMove_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim F As Integer
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    F = mintCurFrame - 2
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If KeyAscii = 0 Then
        Temp = txtMove(Index).Text
    Else
        Temp = FutureText(txtMove(Index), KeyAscii)
    End If
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With MovePick(F)
        Y = Len(Temp)
        For X = 1 To .ListItems.count
            If LCase(Left(.ListItems(X).Text, Y)) = LCase(Temp) And Not .ListItems(X).Checked Then
                Temp = txtMove(Index).Tag
                If Temp <> "" Then
                    If MovePick(F).ListItems(Temp).Checked And Temp <> .ListItems(X).Key Then
                        MovePick(F).ListItems(Temp).Checked = False
                        Call MovePick_ItemCheck(F, MovePick(F).ListItems(Temp))
                    End If
                End If
                txtMove(Index).Tag = .ListItems(X).Key
                .ListItems(X).Selected = True
                .ListItems(X).EnsureVisible
                txtMove(Index).Text = .ListItems(X).Text
                txtMove(Index).SelStart = Y
                txtMove(Index).SelLength = Len(.ListItems(X).Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtMove(Index).SelStart + 1
            txtMove(Index).Text = Temp
            txtMove(Index).SelStart = X
            If txtMove(Index).Tag <> "" Then
                If .ListItems(txtMove(Index).Tag).Checked Then
                    .ListItems(txtMove(Index).Tag).Checked = False
                    Temp = txtMove(Index).Text
                    Call MovePick_ItemCheck(F, .ListItems(txtMove(Index).Tag))
                End If
                txtMove(Index).Tag = ""
                txtMove(Index).Text = Temp
                txtMove(Index).SelStart = Len(Temp)
            End If
        End If
    End With
End Sub

Private Sub txtMove_LostFocus(Index As Integer)
    Dim F As Integer
    On Error GoTo ETrap
    txtMove(Index).SelStart = 0
    txtMove(Index).SelLength = 0
    F = mintCurFrame - 2
    If txtMove(Index).Tag <> "" Then
        If Not MovePick(F).ListItems(txtMove(Index).Tag).Checked Then
            MovePick(F).Tag = Index
            MovePick(F).ListItems(txtMove(Index).Tag).Checked = True
            Call MovePick_ItemCheck(F, MovePick(F).ListItems(txtMove(Index).Tag))
        End If
    End If
ETrap:
End Sub

'Private Sub TeamList_Click()
'    Dim X As Integer
'    Dim OneCompat(0 To 6) As Boolean
'
'    With PKMN(TeamList.ListIndex + 1)
'        MoveList.Clear
'        For X = 1 To 4
'            If .Move(X) > 0 Then
'                MoveList.AddItem Moves(.Move(X)).Name
'            End If
'        Next
'        Call ReadBinArray(CompatCheck(TeamList.ListIndex + 1), OneCompat)
'        For X = 0 To 6
'            OnePKMN(X).Visible = OneCompat(X)
'        Next
'        MoveAdvanced.Caption = ""
'    End With
'End Sub

Private Sub UserName_Change()
    Dim X As Long
    If UserName.SelStart = Len(UserName.Text) Then X = 1
    UserName.Text = CorrectText(UserName.Text, True)
    If X Then UserName.SelStart = Len(UserName.Text)
    X = IIf(UserName.Text = "", 3, 4)
    If MainTabs.Tabs(1).Image <> X Then MainTabs.Tabs(1).Image = X
    You.Name = UserName.Text
End Sub

Private Sub CheckDupeMoves(ByVal Pokemon As Integer)
    Dim X As Integer
    Dim Y As Integer
    
    For X = 1 To 4
        For Y = 1 To 4
            If X <> Y And PKMN(Pokemon).Move(X) = PKMN(Pokemon).Move(Y) Then
                PKMN(Pokemon).Move(Y) = 0
            End If
        Next
    Next
End Sub

Private Sub VersionSelect_Click()
    Dim X As Byte
    Dim Y As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    If Not LoadingForm Then
        Select Case VersionSelect.ListIndex
            Case 0
                You.Version = nbGFXGrn
            Case 1
                You.Version = nbGFXRB
            Case 2
                You.Version = nbGFXYlo
            Case 3
                You.Version = nbGFXGld
            Case 4
                You.Version = nbGFXSil
            Case 5
                You.Version = nbGFXRS
            Case 6
                You.Version = nbGFXLF
            Case 7
                You.Version = nbGFXEme
            Case 8
                You.Version = nbGFXCol
        End Select
        For X = 1 To 6
            Call RefreshImage(X, False)
        Next X
    End If
End Sub
Private Sub RefreshImage(ByVal PokeNum As Long, CheckForChange As Boolean)
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    X = PokeNum - 1
    With PKMN(PokeNum)
        If .No > 0 Then
            Temp = .Image
            .Image = ChooseImage(PKMN(PokeNum), You.Version)
            If Temp <> .Image Or Not CheckForChange Then
                Call MainContainer.DoPicture(.Image)
                picSwap = MainContainer.SwapSpace.Picture
                Y = GetYOffset(picSwap) * Screen.TwipsPerPixelY
                PKMNPic(X).Picture = MainContainer.SwapSpace.Picture
                Y = (PKMNPicBox(X).Height - PKMNPic(X).Height - Y) / 2
                Z = (PKMNPicBox(X).Width - PKMNPic(X).Width) / 2
                PKMNPic(X).Top = Y
                PKMNPic(X).Left = Z
            End If
        End If
    End With
End Sub

Private Sub WinMSG_Change()
    If LoadingForm Then Exit Sub
    WinMSG.Text = CorrectText(WinMSG.Text)
    You.WinMess = WinMSG.Text
End Sub

Sub DoRBY()
    Dim X As Integer
    
    Call ReadBinArray(CompatCheck(PKMN), Compatibility)
    RBY.Visible = Compatibility(0)
    GS.Visible = Compatibility(1)
    ADV.Visible = Compatibility(2)
    AdvPlus.Visible = Compatibility(3)
    ADVTrade.Visible = Compatibility(4)
    TrueRBY.Visible = Compatibility(5)
    TrueGSC.Visible = Compatibility(6)
    Call RefreshCompatTree
End Sub
Sub RefreshCompatTree()
    Dim MainNode As Node
    Dim PokeNode As Node
    Dim TempNode As Node
    Dim X As Byte
    Dim Y As Byte
    
    With CompatTree
        .Nodes.Clear
        Set MainNode = .Nodes.Add(, , "Name", IIf(You.Name <> "", You.Name, "(Your Name)"))
        For X = 1 To 6
            If PKMN(X).No > 0 Then
                If PKMN(X).Nickname <> PKMN(X).Name Then
                    Set TempNode = .Nodes.Add(MainNode, tvwChild, "Poke" & X, PKMN(X).Nickname & " (" & PKMN(X).Name & ")")
                Else
                    Set TempNode = .Nodes.Add(MainNode, tvwChild, "Poke" & X, PKMN(X).Name)
                End If
                If PKMN(X).Item <> nbNoItem Then
                    .Nodes.Add TempNode, tvwChild, "Poke" & X & "Item", "@ " & Item(PKMN(X).Item)
                End If
                For Y = 1 To 4
                    If PKMN(X).Move(Y) > 0 Then
                        .Nodes.Add TempNode, tvwChild, "Poke" & X & "Move" & Y, Moves(PKMN(X).Move(Y)).Name, Moves(PKMN(X).Move(Y)).Type, Moves(PKMN(X).Move(Y)).Type
                    Else
                        .Nodes.Add TempNode, tvwChild, "Poke" & X & "Move" & Y, "(No Move)"
                    End If
                Next Y
            Else
                Set TempNode = .Nodes.Add(MainNode, tvwChild, "Poke" & X, "(No Pokémon)")
            End If
        Next X
        MainNode.Expanded = True
        MainNode.Selected = True
        Call CompatTree_NodeClick(MainNode)
    End With
End Sub
Sub DoPower()
    If PKMN(1).No > 0 And PKMN(2).No > 0 And PKMN(3).No > 0 And PKMN(4).No > 0 And PKMN(5).No > 0 And PKMN(6).No > 0 Then
        TeamBuilder.Caption = "Team Builder -" & TeamRank & "%"
    Else
        TeamBuilder.Caption = "Team Builder"
    End If
End Sub

Sub FillSpeciesList()
    Dim X As Integer
    Dim Y As Integer
    Dim PDSort() As Integer
    Dim TempSort() As Integer
    Dim StrSort() As String
    Dim UBase As Integer
    'MainContainer.MousePointer = vbArrowHourglass
    Call WriteDebugLog("Checkpoint")
    UBase = UBound(BasePKMN)
    ReDim StrSort(1 To UBase)
    ReDim PDSort(1 To UBase)
    Select Case TBSort
    Case 0
        For X = 1 To UBase
            PDSort(X) = X
        Next X
    Case 1
        For X = 1 To UBase
            PDSort(BasePKMN(X).GSNo) = X
        Next X
    Case 2
        For X = 1 To UBase
            PDSort(BasePKMN(X).AdvNo) = X
        Next X
    Case 3
        For X = 1 To UBase
            StrSort(X) = BasePKMN(X).Name
        Next X
        Call SortStringArray(StrSort)
        For X = 1 To UBase
            PDSort(X) = GetPokeNum(StrSort(X))
        Next X
    End Select
    
    TempSort = PDSort
    Select Case TBMode
    Case 0, 5
        Y = 1
        For X = 1 To UBase
            If BasePKMN(PDSort(X)).ExistRBY Then
                TempSort(Y) = PDSort(X)
                Y = Y + 1
            End If
        Next X
        ReDim Preserve TempSort(1 To Y - 1)
    Case 1
        Y = 1
        For X = 1 To UBase
            If BasePKMN(PDSort(X)).ExistRBY Or BasePKMN(PDSort(X)).ExistGSC Then
                TempSort(Y) = PDSort(X)
                Y = Y + 1
            End If
        Next X
        ReDim Preserve TempSort(1 To Y - 1)
    Case 2
        Y = 1
        For X = 1 To UBase
            If BasePKMN(PDSort(X)).ExistAdv Then
                TempSort(Y) = PDSort(X)
                Y = Y + 1
            End If
        Next X
        ReDim Preserve TempSort(1 To Y - 1)
    Case 6
        Y = 1
        For X = 1 To UBase
            If BasePKMN(PDSort(X)).ExistGSC Then
                TempSort(Y) = PDSort(X)
                Y = Y + 1
            End If
        Next X
        ReDim Preserve TempSort(1 To Y - 1)
    End Select
    ReDim StrSort(1 To UBound(TempSort))
    For X = 1 To UBound(TempSort)
        StrSort(X) = BasePKMN(TempSort(X)).Name
    Next X
    Call WriteDebugLog("Checkpoint")
    'LockWindowUpdate MasterSpecies.hwnd
    MasterSpecies.Visible = False
    MasterSpecies.Clear
    Call WriteDebugLog("Checkpoint")
    With MasterSpecies
        For X = 1 To UBound(TempSort)
            .AddItem StrSort(X), X - 1
        Next X
    End With
    Call WriteDebugLog("Checkpoint")
    On Error GoTo NoPokeSelected
    If PKMN(mintCurFrame - 1).Name <> "" Then
        For Y = 0 To MasterSpecies.ListCount - 1
            If MasterSpecies.List(Y) = PKMN(mintCurFrame - 1).Name Then MasterSpecies.ListIndex = Y
        Next Y
    Else
        MasterSpecies.ListIndex = -1
    End If
NoPokeSelected:
    MasterSpecies.Visible = (mintCurFrame <> 1 And mintCurFrame <> 8)
    'LockWindowUpdate 0
    'MainContainer.MousePointer = vbNormal
End Sub

Private Sub RedrawTB()
    Static LastRedraw As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Byte
    Dim TMove() As Integer
    
    If LastRedraw = TBMode And Not LoadingForm Then Exit Sub
    LastRedraw = TBMode
    
    Call WriteDebugLog("Redraw.  TBMode is " & TBMode)
    Select Case TBMode
    'RBY
    Case 0, 5
        For X = 0 To 5
            Call WriteDebugLog("Setting Visibility")
            'ItemPick(X).Visible = False
            Label8(X).Visible = False
            Label16(X).Caption = "Special:"
            Label18(X).Visible = False
            'GenderDisp(X).Visible = False
            SpecialDefense(X).Visible = False
        Next
        For X = 1 To 6
            If Not PKMN(X).ExistRBY Then PKMN(X) = BlankPKMN: Call ResetSlot(X)
        Next
    Case 1, 6
        For X = 0 To 5
            'ItemPick(X).Visible = True
            Label8(X).Visible = True
            Label16(X).Caption = "Sp.Attack:"
            Label18(X).Visible = True
            'GenderDisp(X).Visible = True
            SpecialDefense(X).Visible = True
        Next
        For X = 1 To 6
            If Not (PKMN(X).ExistRBY Or PKMN(X).ExistGSC) Then PKMN(X) = BlankPKMN: Call ResetSlot(X)
        Next
    Case Else
        For X = 0 To 5
            'ItemPick(X).Visible = True
            Label8(X).Visible = True
            Label16(X).Caption = "Sp.Attack:"
            Label18(X).Visible = True
            'GenderDisp(X).Visible = True
            SpecialDefense(X).Visible = True
        Next
        For X = 1 To 6
            If TBMode = 2 And Not PKMN(X).ExistAdv Then PKMN(X) = BlankPKMN: Call ResetSlot(X)
        Next
    End Select

    Call WriteDebugLog("Settting Box Colors")
    Call RefreshBoxMinor
    
    Call WriteDebugLog("Checking Illegalities")
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            TMove = PKMN(X).Move
            For Y = 1 To 4
                PKMN(X).Move(Y) = 0
            Next Y
            Z = 1
            For Y = 1 To 4
                PKMN(X).Move(Z) = TMove(Y)
                If LegalMove(PKMN(X)) <> "" Then
                    PKMN(X).Move(Z) = 0
                End If
                If PKMN(X).Move(Z) > 0 Then Z = Z + 1
            Next
        End If
    Next X
    
    Call WriteDebugLog("Filling Items")
    Call FillItems
    Call WriteDebugLog("Filling Species List")
    Call FillSpeciesList
    picMoves.Visible = Not (mintCurFrame = 1 Or mintCurFrame = 8)
    DoEvents
    Call WriteDebugLog("Filling Moves")
    For X = 0 To 5
        Call FillInMoveList(X)
    Next X
    
    Call WriteDebugLog("Setting Mode Settings")
    For X = 0 To 9
        mnuVersionItem(X).Checked = False
    Next X
    Select Case TBMode
    Case 0
        imgTBMode.Picture = LoadResPicture("RBYT", vbResIcon)
        mnuVersionItem(1).Checked = True
    Case 1
        imgTBMode.Picture = LoadResPicture("GSCT", vbResIcon)
        mnuVersionItem(4).Checked = True
    Case 2
        imgTBMode.Picture = LoadResPicture("ADV", vbResIcon)
        mnuVersionItem(6).Checked = True
    Case 3
        imgTBMode.Picture = LoadResPicture("ADV+", vbResIcon)
        mnuVersionItem(7).Checked = True
    Case 4
        imgTBMode.Picture = LoadResPicture("MOD", vbResIcon)
        mnuVersionItem(9).Checked = True
    Case 5
        imgTBMode.Picture = LoadResPicture("RBY", vbResIcon)
        mnuVersionItem(0).Checked = True
    Case 6
        imgTBMode.Picture = LoadResPicture("GSC", vbResIcon)
        mnuVersionItem(3).Checked = True
    End Select
    StatusBar1.Panels(2).Text = ModeText(TBMode)
    imgTBMode.ToolTipText = ModeText(TBMode)
End Sub
Private Sub FillItems()
    Dim X As Long
    Dim Y As Byte
    Dim Temp As String
    Dim TItem() As String
    Select Case CompatVersion(TBMode)
    Case nbRBYBattle
        ReDim TItem(0)
        TItem(0) = "N/A"
    Case nbGSCBattle
        TItem = Item
        ReDim Preserve TItem(41)
        TItem(0) = ""
        Call SortStringArray(TItem)
        TItem(0) = Item(0)
    Case nbAdvBattle
        Y = 1
        For X = 1 To UBound(Item)
            If Item(X) <> "" Then
                If AdvItem(X) Then
                    ReDim Preserve TItem(Y)
                    TItem(Y) = Item(X)
                    Y = Y + 1
                End If
            End If
        Next X
        TItem(0) = ""
        Call SortStringArray(TItem)
        TItem(0) = Item(0)
    End Select
    
    Temp = MasterItem.List(MasterItem.ListIndex)
    MasterItem.Visible = False
    MasterItem.Clear
    For Y = 0 To UBound(TItem)
        MasterItem.AddItem TItem(Y)
    Next Y
    If mintCurFrame > 1 And mintCurFrame < 8 Then
        If Temp = Item(PKMN(mintCurFrame - 1).Item) Then
            For Y = 0 To UBound(TItem)
                If MasterItem.List(Y) = Temp Then Exit For
            Next Y
            MasterItem.ListIndex = IIf(Y > UBound(TItem), 0, Y)
        End If
    End If
    MasterItem.Enabled = (CompatVersion(TBMode) <> nbRBYBattle)
    MasterItem.Visible = (mintCurFrame <> 1 And mintCurFrame <> 8)
End Sub
Private Sub ResetSlot(ByVal Slot As Integer)
    On Error Resume Next
    Slot = Slot - 1
    PKMNPic(Slot).Picture = LoadPicture
    Nickname(Slot).Text = ""
    HP(Slot).Caption = "???"
    Attack(Slot).Caption = "???"
    Defense(Slot).Caption = "???"
    Speed(Slot).Caption = "???"
    SpecialAttack(Slot).Caption = "???"
    SpecialDefense(Slot).Caption = "???"
    GenderDisp(Slot).Caption = "???"
    Type1(Slot).Caption = "???"
    Type2(Slot).Caption = "???"
    'ItemPick(Slot).ListIndex = 0
    'Species(Slot).ListIndex = 0
    MovePick(Slot).ListItems.Clear
    Rebuild(Slot).Enabled = True
End Sub

Private Sub mnuPokedexItem_Click(Index As Integer)
    Dim DexMode As Byte
    Dim DexPoke As Integer
        
    Select Case TBMode
        Case 0, 5
            DexMode = 0
        Case 1, 6
            DexMode = 1
        Case Else
            DexMode = 2
    End Select
    If mintCurFrame >= 2 And mintCurFrame <= 7 Then DexPoke = PKMN(mintCurFrame - 1).No
    If DexPoke = 0 Then DexPoke = 1
    MasterDex.Show
    Call MasterDex.SetMode(Index)
    Call MasterDex.SetVer(DexMode)
    Call MasterDex.SetPoke(DexPoke)
End Sub

Sub RefreshRecentFiles()
    'Refresh the Recent File listing
    Dim X As Integer
    For X = 1 To 4
        If RecentFiles(X) <> "" Then
            mnuFileItem(X + 8).Caption = "&" & X & " " & RecentFiles(X)
            mnuFileItem(X + 8).Enabled = True
        Else
            mnuFileItem(X + 8).Caption = "&" & X & " (No Recent File)"
            mnuFileItem(X + 8).Enabled = False
        End If
    Next X
End Sub

Function CreateTeamString() As String
    Dim X As Byte
    Dim Build As String
    With You
        Build = Pad(.Name, 20) & Hex(.Picture) & CStr(.Version) & Pad(Left(.Extra, 200), 200) & Pad(Left(.WinMess, 240), 240) & Pad(Left(.LoseMess, 240), 240)
    End With
    'Debug.Print Len(Build)
    For X = 1 To 6
        Build = Build & PKMN2Str(PKMN(X))
    Next X
    'Debug.Print Len(Build)
    CreateTeamString = Build
End Function

Sub DoBoxChange()
    If ShowBoxes Then
        BoxFrame.Visible = True
        TeamBuilder.Width = 9255 '9945
        ShowBox.Caption = "<< &Box"
        Call RefreshBox
    Else
        BoxFrame.Visible = False
        TeamBuilder.Width = 5745
        ShowBox.Caption = "&Box >>"
    End If
    ShowBox.Refresh
End Sub

Sub DoVerChange(ByVal NewVer As CompatModes, ByVal OldVer As CompatModes)
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Integer
    Dim BlankPKMN As Pokemon
    Dim ChangeString As String
    Dim IsMove As Boolean
    
    'First, wipe out any bad species & set versions
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            Select Case NewVer
                Case nbTrueRBY, nbRBYTrade
                    If Not PKMN(X).ExistRBY Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Removed" & vbCrLf
                        PKMN(X) = BlankPKMN
                    End If
                Case nbTrueGSC
                    If Not PKMN(X).ExistGSC Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Removed" & vbCrLf
                        PKMN(X) = BlankPKMN
                    End If
                Case nbGSCTrade
                    If Not PKMN(X).ExistGSC And Not PKMN(X).ExistRBY Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Removed" & vbCrLf
                        PKMN(X) = BlankPKMN
                    End If
                Case nbTrueRuSa
                    If Not PKMN(X).ExistAdv Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Removed" & vbCrLf
                        PKMN(X) = BlankPKMN
                    End If
            End Select
            PKMN(X).GameVersion = NewVer
        End If
    Next
    
    'Here we'll adjust DVs, EVs, Natures, etc. if going between versions
    Select Case NewVer
        Case nbTrueRBY, nbRBYTrade, nbTrueGSC, nbGSCTrade
            Select Case OldVer
                Case nbTrueRuSa, nbFullAdvance, nbModAdv
                    For X = 1 To 6
                        If PKMN(X).No > 0 Then
                            PKMN(X).DV_Atk = 15
                            PKMN(X).DV_Def = 15
                            PKMN(X).DV_HP = 15
                            PKMN(X).DV_SAtk = 15
                            PKMN(X).DV_SDef = 15
                            PKMN(X).DV_Spd = 15
                        End If
                    Next
                Case Else
                    'Values are Ok
            End Select
        Case nbTrueRuSa, nbFullAdvance, nbModAdv
            Select Case OldVer
                Case nbTrueRBY, nbRBYTrade, nbTrueGSC, nbGSCTrade
                    For X = 1 To 6
                        With PKMN(X)
                            If .No > 0 Then
                                .DV_Atk = 31
                                .DV_Def = 31
                                .DV_HP = 31
                                .DV_SAtk = 31
                                .DV_SDef = 31
                                .DV_Spd = 31
                                .NatureNum = 0
                                If NewVer = nbModAdv Then
                                    .Attribute = BasePKMN(.No).ModAttr(0)
                                Else
                                    .Attribute = BasePKMN(.No).PAtt(0)
                                End If
                                .AttNum = 1
                                .EV_Atk = 85
                                .EV_Def = 85
                                .EV_HP = 85
                                .EV_SAtk = 85
                                .EV_SDef = 85
                                .EV_Spd = 85
                            End If
                        End With
                    Next
                Case Else
                    'Values are Ok
            End Select
    End Select
    
    'Item Check
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            Select Case NewVer
                Case nbTrueRBY, nbRBYTrade
                    If PKMN(X).Item <> nbNoItem Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Item Removed" & vbCrLf
                        PKMN(X).Item = nbNoItem
                    End If
                Case nbTrueGSC, nbGSCTrade
                    If PKMN(X).Item > 41 Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Advance-Only Item Removed" & vbCrLf
                        PKMN(X).Item = nbNoItem
                    End If
                Case nbTrueRuSa, nbFullAdvance, nbModAdv
                    If Not AdvItem(PKMN(X).Item) Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - GSC-Only Item Removed" & vbCrLf
                        PKMN(X).Item = nbNoItem
                    End If
            End Select
        End If
    Next
    
    'Check for moves
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            For Y = 1 To 4
                If PKMN(X).Move(Y) > 0 Then
                    Select Case NewVer
                        Case nbTrueRBY, nbRBYTrade
                            If Not Moves(PKMN(X).Move(Y)).RBYMove Then
                                ChangeString = ChangeString & PKMN(X).Nickname & " - Removed " & Moves(PKMN(X).Move(Y)).Name & "(Newer Move)" & vbCrLf
                                PKMN(X).Move(Y) = 0
                            End If
                        Case nbTrueGSC, nbGSCTrade
                            If Not Moves(PKMN(X).Move(Y)).GSCMove Then
                                ChangeString = ChangeString & PKMN(X).Nickname & " - Removed " & Moves(PKMN(X).Move(Y)).Name & "(Newer Move)" & vbCrLf
                                PKMN(X).Move(Y) = 0
                            End If
                    End Select
                End If
            Next
        End If
    Next
    
    'Another move check - can't learn
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            For Y = 1 To 4
                If PKMN(X).Move(Y) > 0 Then
                    IsMove = False
                    Select Case NewVer
                        Case nbTrueRBY
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).RBYMoves)
                                If BasePKMN(PKMN(X).No).RBYMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).RBYTM)
                                If BasePKMN(PKMN(X).No).RBYTM(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                        Case nbTrueGSC
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).BaseMoves)
                                If BasePKMN(PKMN(X).No).BaseMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).MachineMoves)
                                If BasePKMN(PKMN(X).No).MachineMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).BreedingMoves)
                                If BasePKMN(PKMN(X).No).BreedingMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).SpecialMoves)
                                If BasePKMN(PKMN(X).No).SpecialMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).MoveTutor)
                                If BasePKMN(PKMN(X).No).MoveTutor(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                        Case nbRBYTrade, nbGSCTrade
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).RBYMoves)
                                If BasePKMN(PKMN(X).No).RBYMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).RBYTM)
                                If BasePKMN(PKMN(X).No).RBYTM(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).BaseMoves)
                                If BasePKMN(PKMN(X).No).BaseMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).MachineMoves)
                                If BasePKMN(PKMN(X).No).MachineMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).BreedingMoves)
                                If BasePKMN(PKMN(X).No).BreedingMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).SpecialMoves)
                                If BasePKMN(PKMN(X).No).SpecialMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).MoveTutor)
                                If BasePKMN(PKMN(X).No).MoveTutor(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                        Case nbTrueRuSa
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvMoves)
                                If BasePKMN(PKMN(X).No).AdvMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).ADVTM)
                                If BasePKMN(PKMN(X).No).ADVTM(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvBreeding)
                                If BasePKMN(PKMN(X).No).AdvBreeding(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvSpecial)
                                If BasePKMN(PKMN(X).No).AdvSpecial(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                        Case nbFullAdvance, nbModAdv
                            If NewVer = nbModAdv Then Z = UBound(BasePKMN(PKMN(X).No).AdvMoves) Else Z = PKMN(X).TotalAdvMoves
                            For Z = 1 To Z
                                If BasePKMN(PKMN(X).No).AdvMoves(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).ADVTM)
                                If BasePKMN(PKMN(X).No).ADVTM(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvBreeding)
                                If BasePKMN(PKMN(X).No).AdvBreeding(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvSpecial)
                                If BasePKMN(PKMN(X).No).AdvSpecial(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).AdvTutor)
                                If BasePKMN(PKMN(X).No).AdvTutor(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                            For Z = 1 To UBound(BasePKMN(PKMN(X).No).LFOnly)
                                If BasePKMN(PKMN(X).No).LFOnly(Z) = PKMN(X).Move(Y) Then IsMove = True
                            Next
                    End Select
                    If IsMove = False Then
                        ChangeString = ChangeString & PKMN(X).Nickname & " - Removed " & Moves(PKMN(X).Move(Y)).Name & "(Can't Learn)" & vbCrLf
                        PKMN(X).Move(Y) = 0
                    End If
                End If
            Next
        End If
    Next
    
    'Now, we'll clear out the moves for anything with illegal combos
    For X = 1 To 6
        If PKMN(X).No > 0 Then
            If LegalMove(PKMN(X)) <> "" Then
                ChangeString = ChangeString & PKMN(X).Nickname & " - Moves cleared (Illegal Combo)" & vbCrLf
                For Y = 1 To 4
                    PKMN(X).Move(Y) = 0
                Next Y
            End If
        End If
    Next
    
    If ChangeString <> "" Then MsgBox ChangeString, vbInformation, "Team Changes"
End Sub
