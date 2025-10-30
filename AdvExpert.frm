VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AdvExpert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expert Mode: [Pokemon Name]"
   ClientHeight    =   6975
   ClientLeft      =   10515
   ClientTop       =   3780
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPlaceholder 
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   37
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton cmdPlaceholder 
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   90
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Final Stats"
      Height          =   2655
      Left            =   120
      TabIndex        =   64
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4800
         ScaleHeight     =   615
         ScaleWidth      =   855
         TabIndex        =   85
         Top             =   160
         Width           =   855
         Begin VB.TextBox txtLevel 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            MaxLength       =   3
            TabIndex        =   86
            TabStop         =   0   'False
            Text            =   "100"
            Top             =   240
            Width           =   375
         End
         Begin VB.HScrollBar hscLevel 
            Height          =   255
            Left            =   0
            Max             =   100
            Min             =   1
            TabIndex        =   14
            Top             =   250
            Value           =   100
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Level:"
            Height          =   735
            Left            =   120
            TabIndex        =   87
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5655
         TabIndex        =   65
         Top             =   240
         Width           =   5655
         Begin VB.PictureBox picUnown 
            BorderStyle     =   0  'None
            Height          =   1215
            Left            =   4680
            ScaleHeight     =   1215
            ScaleWidth      =   975
            TabIndex        =   88
            Top             =   135
            Width           =   975
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "AdvExpert.frx":0000
               Left            =   120
               List            =   "AdvExpert.frx":0058
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   "Unown Letter:"
               Height          =   615
               Left            =   120
               TabIndex        =   89
               Top             =   420
               Width           =   735
            End
         End
         Begin VB.ComboBox cmbHiddenPower 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1920
            Width           =   1335
         End
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   3120
            ScaleHeight     =   735
            ScaleWidth      =   2535
            TabIndex        =   81
            Top             =   560
            Width           =   2535
            Begin VB.OptionButton optTrait 
               Caption         =   "Levitate"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton optTrait 
               Caption         =   "Trait 2"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   480
               Width           =   2235
            End
            Begin VB.Label Label2 
               Caption         =   "Trait:"
               Height          =   255
               Left            =   0
               TabIndex        =   82
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.OptionButton optGender 
            Caption         =   "Female"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   12
            Top             =   1800
            Width           =   855
         End
         Begin VB.OptionButton optGender 
            Caption         =   "Male"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   11
            Top             =   1560
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.ComboBox cmbNature 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
         Begin VB.CheckBox chkShiny 
            Caption         =   "Shiny"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   0
            ItemData        =   "AdvExpert.frx":00CC
            Left            =   1980
            List            =   "AdvExpert.frx":0130
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   1
            ItemData        =   "AdvExpert.frx":01B4
            Left            =   1980
            List            =   "AdvExpert.frx":0218
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   2
            ItemData        =   "AdvExpert.frx":029C
            Left            =   1980
            List            =   "AdvExpert.frx":0300
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   3
            ItemData        =   "AdvExpert.frx":0384
            Left            =   1980
            List            =   "AdvExpert.frx":03E8
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   855
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   4
            ItemData        =   "AdvExpert.frx":046C
            Left            =   1980
            List            =   "AdvExpert.frx":04D0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1560
            Width           =   855
         End
         Begin VB.ComboBox cmbIV 
            Height          =   315
            Index           =   5
            ItemData        =   "AdvExpert.frx":0554
            Left            =   1980
            List            =   "AdvExpert.frx":05B8
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Gender:"
            Height          =   255
            Left            =   3120
            TabIndex        =   83
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblHiddenPower 
            Caption         =   "Hidden Power: "
            Height          =   255
            Left            =   4320
            TabIndex        =   80
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000012&
            X1              =   1020
            X2              =   1020
            Y1              =   60
            Y2              =   2280
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Def"
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
            Index           =   5
            Left            =   0
            TabIndex        =   79
            Top             =   1920
            Width           =   900
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            Caption         =   "Sp. Atk"
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
            Index           =   4
            Left            =   0
            TabIndex        =   78
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed"
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
            Index           =   3
            Left            =   0
            TabIndex        =   77
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense"
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
            Index           =   2
            Left            =   0
            TabIndex        =   76
            Top             =   840
            Width           =   900
         End
         Begin VB.Label lbStat 
            Alignment       =   1  'Right Justify
            Caption         =   "Attack"
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
            Index           =   1
            Left            =   0
            TabIndex        =   75
            Top             =   480
            Width           =   900
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            Caption         =   "HP"
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
            Index           =   0
            Left            =   0
            TabIndex        =   74
            Top             =   120
            Width           =   900
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "22"
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
            Index           =   5
            Left            =   1020
            TabIndex        =   73
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "245"
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
            Index           =   4
            Left            =   1020
            TabIndex        =   72
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "742"
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
            Index           =   3
            Left            =   1020
            TabIndex        =   71
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
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
            Index           =   2
            Left            =   1020
            TabIndex        =   70
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "222"
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
            Index           =   1
            Left            =   1020
            TabIndex        =   69
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblValue 
            Alignment       =   1  'Right Justify
            Caption         =   "703"
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
            Index           =   0
            Left            =   1020
            TabIndex        =   68
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblPlus 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   1620
            TabIndex        =   67
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblMinus 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1620
            TabIndex        =   66
            Top             =   1920
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Line Line2 
            X1              =   3000
            X2              =   3000
            Y1              =   60
            Y2              =   2400
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EV Distribution"
      Height          =   3615
      Left            =   120
      TabIndex        =   38
      Top             =   2760
      Width           =   5895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3315
         Left            =   50
         ScaleHeight     =   3315
         ScaleWidth      =   5820
         TabIndex        =   39
         Top             =   240
         Width           =   5820
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   315
            Left            =   4800
            TabIndex        =   36
            Top             =   3000
            Width           =   900
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   315
            Left            =   3840
            TabIndex        =   35
            Top             =   3000
            Width           =   900
         End
         Begin VB.CommandButton Preset 
            Caption         =   "Apply"
            Height          =   315
            Left            =   2880
            TabIndex        =   34
            Top             =   3000
            Width           =   900
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   -240
            ScaleHeight     =   2415
            ScaleWidth      =   1035
            TabIndex        =   56
            Top             =   -150
            Width           =   1030
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Sp. Def"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   62
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Sp. Atk"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   61
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Speed"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   60
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Defense"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   59
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Attack"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   58
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Alignment       =   1  'Right Justify
               Caption         =   "HP"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   57
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   4560
            ScaleHeight     =   2415
            ScaleWidth      =   1575
            TabIndex        =   49
            Top             =   -150
            Width           =   1575
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   23
               Top             =   220
               Width           =   735
            End
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   24
               Top             =   580
               Width           =   735
            End
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   25
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   26
               Top             =   1300
               Width           =   735
            End
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   4
               Left            =   480
               TabIndex        =   27
               Top             =   1660
               Width           =   735
            End
            Begin VB.CheckBox EVLock 
               Caption         =   "Lock"
               Height          =   255
               Index           =   5
               Left            =   480
               TabIndex        =   28
               Top             =   2020
               Width           =   735
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   55
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   54
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   53
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   52
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   51
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label lblEVStat 
               Caption         =   "255"
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   50
               Top             =   2040
               Width           =   975
            End
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   1
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   48
            Top             =   720
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   2
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   47
            Top             =   1080
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   3
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   46
            Top             =   1440
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   4
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   45
            Top             =   1800
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   5
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   44
            Top             =   2160
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   6
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   43
            Top             =   0
            Width           =   4200
         End
         Begin VB.PictureBox picBlocker 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   30
            Index           =   0
            Left            =   500
            ScaleHeight     =   30
            ScaleWidth      =   4200
            TabIndex        =   42
            Top             =   360
            Width           =   4200
         End
         Begin VB.PictureBox EVBar 
            AutoRedraw      =   -1  'True
            Height          =   255
            Left            =   60
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   372
            TabIndex        =   40
            Top             =   2640
            Width           =   5640
            Begin VB.PictureBox EVBarCover 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               ScaleHeight     =   255
               ScaleWidth      =   5655
               TabIndex        =   41
               Top             =   0
               Width           =   5655
            End
         End
         Begin VB.CommandButton cmdEVClear 
            Caption         =   "Clear EVs"
            Height          =   315
            Left            =   2880
            TabIndex        =   30
            Top             =   2280
            Width           =   900
         End
         Begin VB.CommandButton cmdUnlock 
            Caption         =   "Unlock All"
            Height          =   315
            Left            =   3840
            TabIndex        =   31
            Top             =   2280
            Width           =   900
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Lock All"
            Height          =   315
            Left            =   4800
            TabIndex        =   32
            Top             =   2280
            Width           =   900
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   0
            Left            =   720
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   -120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   1
            Left            =   720
            TabIndex        =   18
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   2
            Left            =   720
            TabIndex        =   19
            Top             =   600
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   3
            Left            =   720
            TabIndex        =   20
            Top             =   960
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   4
            Left            =   720
            TabIndex        =   21
            Top             =   1320
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin MSComctlLib.Slider EVSlider 
            Height          =   495
            Index           =   5
            Left            =   720
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   873
            _Version        =   393216
            LargeChange     =   4
            Max             =   255
            TickStyle       =   2
         End
         Begin VB.TextBox txtCoverup 
            Height          =   315
            Left            =   960
            MaxLength       =   20
            TabIndex        =   84
            Text            =   "Text1"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1810
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "AdvExpert.frx":063C
            Left            =   960
            List            =   "AdvExpert.frx":063E
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   3000
            Width           =   1815
         End
         Begin VB.CheckBox chkSnap 
            Caption         =   "Snap"
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblEVStat 
            Caption         =   "Remaining EP: 510"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   63
            Top             =   2400
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "AdvExpert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastChange As Byte
Dim Processing As Boolean
Dim Locked(0 To 5) As Boolean
Dim EVBalance(0 To 5) As Single
Dim DoneLoading As Boolean
Dim OKClick As Boolean
Dim OrigPKMN As Pokemon
Dim EVSaves() As Byte

Private Sub chkShiny_Click()
    ExpertPKMN.Shiny = (chkShiny.Value = 1)
End Sub

Private Sub chkSnap_Click()
    Dim X As Long
    Dim Y As Long
    Y = IIf(chkSnap.Value = 1, 4, 1)
    For X = 0 To 5
        EVSlider(X).SmallChange = Y
    Next X
End Sub

Private Sub cmbHiddenPower_Change()
    Dim T(0 To 5) As Integer
    Dim X As Byte
    If Processing Then Exit Sub
    Select Case cmbHiddenPower.ListIndex
    Case 0: T(2) = 1: T(3) = 1: T(4) = 1
    Case 1: T(3) = 1: T(4) = 1
    Case 2: T(4) = 1
    Case 3: T(2) = 1: T(4) = 1
    Case 4: T(3) = 1
    Case 5: T(2) = 1: T(3) = 1: T(4) = 1: T(5) = 1
    Case 6: T(2) = 1: T(4) = 1: T(5) = 1
    Case 7: T(4) = 1: T(5) = 1
    Case 8: T(3) = 1: T(4) = 1: T(5) = 1
    Case 9: T(2) = 1: T(3) = 1
    Case 10: T(3) = 1: T(5) = 1
    Case 11: T(2) = 1: T(3) = 1: T(5) = 1
    Case 12: T(2) = 1: T(5) = 1
    Case 13: T(2) = 1
    Case 14
    Case 15: T(5) = 1
    End Select
    For X = 0 To 5
        cmbIV(X).ListIndex = T(X)
    Next X
End Sub

Private Sub cmbHiddenPower_Click()
    Call cmbHiddenPower_Change
End Sub

Private Sub cmbHiddenPower_KeyUp(KeyCode As Integer, Shift As Integer)
    Call cmbHiddenPower_Change
End Sub

Private Sub cmbNature_Change()
    Dim X As Integer
    
    If cmbNature.ListIndex Mod 5 = cmbNature.ListIndex \ 5 Then
        lblPlus.Visible = False
        lblMinus.Visible = False
    Else
        lblPlus.Visible = True
        lblMinus.Visible = True
        For X = 1 To 5
            If Nature(cmbNature.ListIndex).StatChg(X) = 1 Then lblPlus.Top = lblValue(X).Top - 75
            If Nature(cmbNature.ListIndex).StatChg(X) = -1 Then lblMinus.Top = lblValue(X).Top - 75
        Next
    End If
    Call RefreshStats
End Sub

Private Sub cmbNature_Click()
    Call cmbNature_Change
End Sub

Private Sub cmbNature_KeyUp(KeyCode As Integer, Shift As Integer)
    Call cmbNature_Change
End Sub

Private Sub cmdDelete_Click()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Byte
    X = Combo2.ListIndex
    For Y = X - 2 To UBound(EVSaves, 2) - 1
        For Z = 0 To 5
            EVSaves(Z, Y) = EVSaves(Z, Y + 1)
        Next Z
    Next Y
    ReDim Preserve EVSaves(0 To 5, Y - 1)
    Combo2.RemoveItem X
    If X = Combo2.ListCount Then X = X - 1
    Combo2.ListIndex = X
End Sub

Private Sub cmdEVClear_Click()
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Integer
    Dim A() As Integer
    Dim B As Byte
    Dim Temp As String
    Processing = True
    For X = 0 To 5
        If Not Locked(X) Then EVSlider(X).Value = 0
    Next X
    Processing = False
    RefreshEVBar
    RefreshEVLabels
    RefreshStats
End Sub
Private Sub cmdLock_Click()
    Dim X As Byte
    For X = 0 To 5
        EVLock(X).Value = 1
        EVLock(X).Refresh
    Next X
End Sub

Private Sub cmdPlaceholder_Click(Index As Integer)
    If Index = 1 Then
        cmdDelete.SetFocus
    Else
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPlaceholder_GotFocus(Index As Integer)
    If Index = 1 Then
        If cmdDelete.Enabled Then
            cmdDelete.SetFocus
        Else
            cmdSave.SetFocus
        End If
    Else
        Command1.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
    Dim X As Integer
    Dim Y As Integer
    With txtCoverup
        If .Visible = True Then
            If txtCoverup.Text = "" Then Exit Sub
            ReDim Preserve EVSaves(0 To 5, UBound(EVSaves, 2) + 1)
            For X = 0 To 5
                EVSaves(X, UBound(EVSaves, 2)) = EVSlider(X).Value
            Next X
            Combo2.AddItem .Text
            Combo2.ListIndex = Combo2.ListCount - 1
            txtCoverup.Visible = False
            Combo2.SetFocus
            .Text = ""
            Preset.Enabled = True
            Call Combo2_Change
        Else
            X = 0
            Do
                X = X + 1
                For Y = 3 To Combo2.ListCount - 1
                    If Combo2.List(Y) = "Custom" & X Then Exit For
                Next Y
            Loop Until Y = Combo2.ListCount
            .Text = "Custom" & X
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
            Preset.Enabled = False
            cmdDelete.Enabled = False
        End If
    End With
End Sub

Private Sub cmdSave_LostFocus()
    If txtCoverup.Visible And Me.ActiveControl.Name <> "txtCoverup" Then
        txtCoverup.Visible = False
        txtCoverup.Text = ""
        Preset.Enabled = True
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub cmdUnlock_Click()
    Dim X As Byte
    For X = 0 To 5
        EVLock(X).Value = 0
        EVLock(X).Refresh
    Next X
End Sub

Private Sub cmbIV_Change(Index As Integer)
    Call RefreshStats
End Sub

Private Sub cmbIV_Click(Index As Integer)
    Call cmbIV_Change(Index)
End Sub

Private Sub cmbIV_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call cmbIV_Change(Index)
End Sub


Private Sub Combo1_Change()
    ExpertPKMN.UnownLetter = Combo1.ListIndex
End Sub

Private Sub Combo1_Click()
    Call Combo1_Change
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Combo1_Change
End Sub

Private Sub Combo2_Change()
    cmdDelete.Enabled = (Combo2.ListIndex > 2)
End Sub

Private Sub Combo2_Click()
    Call Combo2_Change
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Combo2_Change
End Sub

Private Sub Command1_Click()
    OKClick = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub EVLock_Click(Index As Integer)
    Locked(Index) = (EVLock(Index).Value = 1)
End Sub

Private Sub RefreshEVBar()
    Dim X As Integer
    X = EVTotal
    lblEVStat(6).Caption = "Remaining EP: " & CStr(510 - X)
    X = EVBar.ScaleWidth - ((X / 510) * EVBar.ScaleWidth) + 1
    If X = EVBarCover.Left Then Exit Sub
    EVBarCover.Left = X
End Sub

Private Sub RefreshEVLabels()
    Dim X As Integer
    For X = 0 To 5
        lblEVStat(X + 7).Caption = CStr(EVSlider(X).Value)
    Next X
'    lblEVStat(0).Caption = "HP EV: " & CStr(EVSlider(0).Value)
'    lblEVStat(1).Caption = "Atk EV: " & CStr(EVSlider(1).Value)
'    lblEVStat(2).Caption = "Def EV: " & CStr(EVSlider(2).Value)
'    lblEVStat(3).Caption = "Spd EV: " & CStr(EVSlider(3).Value)
'    lblEVStat(4).Caption = "SAtk EV: " & CStr(EVSlider(4).Value)
'    lblEVStat(5).Caption = "SDef EV: " & CStr(EVSlider(5).Value)
End Sub

Private Sub EVSlider_Change(Index As Integer)
    Dim B As Boolean
    Dim T As Integer
    Dim X As Byte
    Dim Y As Byte
    Dim Z As Integer
    Dim TempEV(0 To 5) As Integer
    
    If Processing Then Exit Sub
    Processing = True
    If chkSnap.Value = 1 Then EVSlider(Index).Value = (EVSlider(Index).Value \ 4) * 4
    Y = LastChange
    If Y = 6 Then Y = 0
    For X = 0 To 5
        TempEV(X) = EVSlider(X).Value
        T = T + TempEV(X)
    Next X
    While T > 510
        Y = Y + 1
        If Y = 6 Then Y = 0
        If TempEV(Y) <> 0 And Y <> Index And Not Locked(Y) Then
            TempEV(Y) = TempEV(Y) - 1
            T = T - 1
            B = True
        End If
        If Y = LastChange Then
            If B Then
                B = False
            Else 'Nothing can be reduced, the bar goes nowhere.
                Z = 0
                For X = 0 To 5
                    If X <> Index Then Z = Z + TempEV(X)
                Next X
                TempEV(Index) = 510 - Z
                T = 510
            End If
        End If
    Wend
    For X = 0 To 5
        EVSlider(X).Value = TempEV(X)
    Next X
    LastChange = Y
    Call RefreshEVBar
    Call RefreshEVLabels
    Processing = False
    Call RefreshStats
End Sub
Private Sub EVSlider_Scroll(Index As Integer)
    Call EVSlider_Change(Index)
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    On Error Resume Next
    DoneLoading = False
    OrigPKMN = ExpertPKMN
    OKClick = False
    With EVBar
        For X = 1 To .ScaleWidth
            EVBar.Line (X, 0)-(X, .ScaleHeight), RGB(255, 255 - Int((X / .ScaleWidth) * 255), 0)
        Next X
    End With
    
    picUnown.Visible = (ExpertPKMN.No = 201)
    Combo1.ListIndex = ExpertPKMN.UnownLetter
    
    Combo2.AddItem "Balanced", 0
    Combo2.AddItem "Strengths", 1
    Combo2.AddItem "Weaknesses", 2
    Combo2.ListIndex = 0
    X = 0
    ReDim EVSaves(0 To 5, 0)
    Do
        X = X + 1
        Temp = GetSetting("NetBattle", "EV Saves", CStr(X), "")
        If Len(Temp) <> 32 Then Exit Do
        Combo2.AddItem Trim(ChopString(Temp, 20)), X + 2
        ReDim Preserve EVSaves(0 To 5, UBound(EVSaves, 2) + 1)
        For Y = 0 To 5
            EVSaves(Y, UBound(EVSaves, 2)) = Dec(ChopString(Temp, 2))
        Next Y
    Loop
    If ExpertPKMN.GameVersion = nbModAdv Then
        optTrait(0).Caption = AttributeText(ExpertPKMN.ModAttr(0))
        optTrait(1).Caption = AttributeText(ExpertPKMN.ModAttr(1))
    Else
        optTrait(0).Caption = AttributeText(ExpertPKMN.PAtt(0))
        optTrait(1).Caption = AttributeText(ExpertPKMN.PAtt(1))
    End If
    If optTrait(1).Caption = "No Trait" Then optTrait(1).Enabled = False
    optTrait(ExpertPKMN.AttNum).Value = True
    hscLevel.Value = ExpertPKMN.Level
    
    If ExpertPKMN.PercentFemale = -1 Then
        optGender(0).Enabled = False
        optGender(1).Enabled = False
        optGender(0).Value = False
        optGender(1).Value = False
    Else
        optGender(ExpertPKMN.Gender - 1).Value = True
        optGender(0).Enabled = (ExpertPKMN.PercentFemale < 16)
        optGender(1).Enabled = (ExpertPKMN.PercentFemale > 0)
    End If
    
    For X = 0 To 5
        Locked(X) = False
    Next
    
    With ExpertPKMN
        cmbIV(0).ListIndex = 31 - .DV_HP
        cmbIV(1).ListIndex = 31 - .DV_Atk
        cmbIV(2).ListIndex = 31 - .DV_Def
        cmbIV(3).ListIndex = 31 - .DV_Spd
        cmbIV(4).ListIndex = 31 - .DV_SAtk
        cmbIV(5).ListIndex = 31 - .DV_SDef
        EVSlider(0).Value = .EV_HP
        EVSlider(1).Value = .EV_Atk
        EVSlider(2).Value = .EV_Def
        EVSlider(3).Value = .EV_Spd
        EVSlider(4).Value = .EV_SAtk
        EVSlider(5).Value = .EV_SDef
        chkShiny.Value = IIf(.Shiny, 1, 0)
    End With
    
    For X = 0 To 24
        cmbNature.AddItem Nature(X).Name, X
        cmbNature.ListIndex = ExpertPKMN.NatureNum
    Next
    For X = 2 To 17
        cmbHiddenPower.AddItem Element(X), X - 2
    Next X
    AdvExpert.Caption = "Expert Mode: " & ExpertPKMN.Name
    DoneLoading = True
    Call RefreshEVBar
    Call RefreshEVLabels
    Call RefreshStats
    chkSnap.Value = GetSetting("NetBattle", "Options", "Snap", 1)
End Sub
Public Function EVTotal() As Integer
    Dim X As Byte
    Dim Y As Integer
    For X = 0 To 5
        Y = Y + EVSlider(X).Value
    Next X
    EVTotal = Y
End Function

Private Sub GetBalances(Optional ByVal Reversed As Boolean = False)
    Dim BSTotal As Integer
    Dim NAdjust(1 To 5) As Single
    Dim X As Byte
    Dim Y As Byte
    Dim Lowest As Byte
    
    'Note that I divide BaseHP by 2 since HP is so far out of whack with the rest of the stats.
    'This isn't always true, but it can be adjusted if needed.
    
    'Masamune's note: Actually, there's no need to divide by 2 since we're going by BASE stats
    'instead of actual stats.  The Base HP is on the same level as the other stats.
    
    With ExpertPKMN
        'Get adjustments for personality
        For X = 1 To 5
            Select Case Nature(.NatureNum).StatChg(X)
                Case -1
                    NAdjust(X) = 0.9
                Case 0
                    NAdjust(X) = 1
                Case 1
                    NAdjust(X) = 1.1
            End Select
        Next
            
        'Total up base stats
        BSTotal = .BaseHP + (.BaseAttack * NAdjust(1)) + (.BaseDefense * NAdjust(2)) + (.BaseSpeed * NAdjust(3)) + (.BaseSAttack * NAdjust(4)) + (.BaseSDefense * NAdjust(5))
    
        'Figure out the percentage each individual stat contributes
        'Masamune's note: If you leave the percentages as Singles, the
        'EV values will be exact and there will never be any overage.
        EVBalance(0) = (.BaseHP * 100) / BSTotal
        EVBalance(1) = ((.BaseAttack * NAdjust(1)) * 100) / BSTotal
        EVBalance(2) = ((.BaseDefense * NAdjust(2)) * 100) / BSTotal
        EVBalance(3) = ((.BaseSpeed * NAdjust(3)) * 100) / BSTotal
        EVBalance(4) = ((.BaseSAttack * NAdjust(4)) * 100) / BSTotal
        EVBalance(5) = ((.BaseSDefense * NAdjust(5)) * 100) / BSTotal
        Debug.Print vbCrLf
        Debug.Print "HP", .BaseHP, EVBalance(0)
        Debug.Print "ATK", .BaseAttack, EVBalance(1)
        Debug.Print "DEF", .BaseDefense, EVBalance(2)
        Debug.Print "SPD", .BaseSpeed, EVBalance(3)
        Debug.Print "SATK", .BaseSAttack, EVBalance(4)
        Debug.Print "SDEF", .BaseSDefense, EVBalance(5)
        Debug.Print "Totals", BSTotal, EVBalance(0) + EVBalance(1) + EVBalance(2) + EVBalance(3) + EVBalance(4) + EVBalance(5)
        If Not Reversed Then Exit Sub
        BSTotal = 0
        For X = 0 To 5
            EVBalance(X) = 100 - EVBalance(X)
        Next
        Lowest = 255
        For X = 0 To 5
            If EVBalance(X) < Lowest Then Lowest = EVBalance(X)
        Next
        For X = 0 To 5
            EVBalance(X) = EVBalance(X) - Lowest + 1
            BSTotal = BSTotal + EVBalance(X)
        Next
        For X = 0 To 5
            EVBalance(X) = (EVBalance(X) * 100) / BSTotal
        Next
        Debug.Print "HP", .BaseHP, EVBalance(0)
        Debug.Print "ATK", .BaseAttack, EVBalance(1)
        Debug.Print "DEF", .BaseDefense, EVBalance(2)
        Debug.Print "SPD", .BaseSpeed, EVBalance(3)
        Debug.Print "SATK", .BaseSAttack, EVBalance(4)
        Debug.Print "SDEF", .BaseSDefense, EVBalance(5)
        Debug.Print "Totals", BSTotal, EVBalance(0) + EVBalance(1) + EVBalance(2) + EVBalance(3) + EVBalance(4) + EVBalance(5)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Temp As String
    Dim X As Integer
    Dim Y As Byte
    If Not OKClick Then ExpertPKMN = OrigPKMN
    On Error Resume Next
    DeleteSetting "NetBattle", "EV Saves"
    For X = 1 To UBound(EVSaves, 2)
        Temp = Pad(Combo2.List(X + 2), 20)
        For Y = 0 To 5
            Temp = Temp & FixedHex(EVSaves(Y, X), 2)
        Next Y
        SaveSetting "NetBattle", "EV Saves", CStr(X), Temp
    Next X
    SaveSetting "NetBattle", "Options", "Snap", chkSnap.Value
End Sub
Private Sub hscLevel_Change()
    txtLevel.Text = hscLevel.Value
    Call RefreshStats
End Sub

Private Sub hscLevel_Scroll()
    Call hscLevel_Change
End Sub

Private Sub lblMinus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Picture4_MouseDown(Button, Shift, lblMinus.Left + X, lblMinus.Top + Y)
End Sub

Private Sub lblPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Picture4_MouseDown(Button, Shift, lblPlus.Left + X, lblPlus.Top + Y)
End Sub

Private Sub optGender_Click(Index As Integer)
    ExpertPKMN.Gender = IIf(optGender(0).Value, 1, 2)
End Sub

Private Sub optTrait_Click(Index As Integer)
    ExpertPKMN.AttNum = IIf(optTrait(0).Value, 0, 1)
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Integer
    Dim A As Integer
    If X < 1550 Or X > 1950 Then Exit Sub
    If Y < 55 Or Y > 2200 Then Exit Sub
    Z = ((Y - 55) \ 359) - 1
    If Z < 0 Or Z > 4 Then Exit Sub
    If Button = vbLeftButton Then
        A = ExpertPKMN.NatureNum Mod 5
        cmbNature.ListIndex = (Z * 5) + A
    ElseIf Button = vbRightButton Then
        A = ExpertPKMN.NatureNum \ 5
        cmbNature.ListIndex = Z + (A * 5)
    End If
End Sub

Private Sub Preset_Click()
    Dim EVtoUse As Integer  'How much we have to work with
    Dim EVLeft As Integer   'How much we have left
    Dim AdjustQty As Byte   'How many bars we're moving
    Dim EVTemp(6) As Byte   'Temporary for the fancy adjustments
    Dim InOrder(5) As Byte  'Sort array.  Whee.
    Dim GreaterThan As Byte 'More sort stuff
    Dim X As Integer        'Loop stuff
    Dim Y As Integer
    Dim Z As Byte
    
    Processing = True
    
    'Start with the max
    EVtoUse = 510
    AdjustQty = 6
    'Subtract locked bars from the total
    For X = 0 To 5
        If Locked(X) Then
            EVtoUse = EVtoUse - EVSlider(X).Value
            AdjustQty = AdjustQty - 1
        End If
    Next
    EVLeft = EVtoUse
    'All bars are locked - why did you click this button, anyway?
    If AdjustQty = 0 Then Exit Sub
    '5 are locked, more than 255 EV left
    If AdjustQty = 1 And EVTotal > 255 Then Exit Sub
    'This is a bit of a kludge - in case two adjustments are equal, how do we know
    'when it's been set already (since 0 is a valid value)?
    'Preset to 6, since it's outside the range.
    For X = 0 To 5
        InOrder(X) = 6
    Next
    EVTemp(6) = 255
    
    'Clear out the bars, since the Change event might trigger...
'    For X = 0 To 5
'        If Not Locked(X) Then EVSlider(X).Value = 0
'    Next
    'Okay, let's do this.
    Select Case Combo2.ListIndex
        'Balanced
        Case 0
            'Give all unlocked bars and equal share
            For X = 0 To 5
                If Not Locked(X) Then
                    EVSlider(X).Value = EVtoUse \ AdjustQty
                    EVLeft = EVLeft - EVSlider(X).Value
                End If
            Next
            'Okay, if we've got anything left, fill in from the top down.
            X = 0
            While EVLeft > 0
                If Not Locked(X) And EVSlider(X) < 255 Then
                    EVSlider(X).Value = EVSlider(X).Value + 1
                    EVLeft = EVLeft - 1
                End If
                X = X + 1
                If X = 6 Then X = 0
            Wend
        'Strongest/Weakest first
        Case 1, 2
            'Let's calculate the values ahead of time here.
            Call GetBalances(Combo2.ListIndex = 2)
            For X = 0 To 5
                If Not Locked(X) Then
                    EVTemp(X) = CByte(Cap(Round((EVtoUse * EVBalance(X)) / 100), 255))
                    EVLeft = EVLeft - EVTemp(X)
                End If
            Next
            
            'Sort 'em out
            For X = 0 To 5
                For Y = 0 To 5
                    If EVTemp(InOrder(Y)) >= EVTemp(X) Then Exit For
                Next
                For Y = 4 To Y Step -1
                    InOrder(Y + 1) = InOrder(Y)
                Next Y
                InOrder(Y + 1) = X
            Next
            
            'Okay, let's make sure we didn't go too far either way...
            Y = Sgn(EVLeft)
            X = IIf(Y = 1, 0, 5)
            While EVLeft <> 0
                Z = InOrder(X)
                If Not Locked(Z) And Not ((EVTemp(Z) = 255 And Y = 1) Or (EVTemp(Z) = 0 And Y = -1)) Then
                    EVTemp(Z) = EVTemp(Z) + Y
                    EVLeft = EVLeft - Y
                End If
                X = X + Y
                If X = 6 Then X = 0
                If X = -1 Then X = 5
            Wend
            
            'Since we're Ok, do the change.
            For X = 0 To 5
                If Not Locked(X) Then
                    EVSlider(X).Value = EVTemp(X)
                End If
            Next
            
        'Custom
        Case Else
            X = Combo2.ListIndex - 2
            EVtoUse = 0
            For Y = 0 To 5
                If Locked(Y) Then
                    EVTemp(Y) = EVSlider(Y).Value
                Else
                    EVTemp(Y) = EVSaves(Y, X)
                End If
                EVtoUse = EVtoUse + EVTemp(Y)
            Next Y
            
            X = 0
            While EVtoUse > 510
                If Not Locked(X) And EVTemp(X) > 0 Then
                    EVTemp(X) = EVTemp(X) - 1
                    EVtoUse = EVtoUse - 1
                End If
                X = X + 1
                If X = 6 Then X = 0
            Wend
            
            For X = 0 To 5
                If Not Locked(X) Then
                    EVSlider(X).Value = EVTemp(X)
                End If
            Next

    End Select
    Processing = False
    Call RefreshEVBar
    Call RefreshEVLabels
    Call RefreshStats
End Sub

Private Sub RefreshStats()
    If Not DoneLoading Then Exit Sub
    With ExpertPKMN
        .DV_HP = 31 - cmbIV(0).ListIndex
        .DV_Atk = 31 - cmbIV(1).ListIndex
        .DV_Def = 31 - cmbIV(2).ListIndex
        .DV_Spd = 31 - cmbIV(3).ListIndex
        .DV_SAtk = 31 - cmbIV(4).ListIndex
        .DV_SDef = 31 - cmbIV(5).ListIndex
        .EV_HP = EVSlider(0).Value
        .EV_Atk = EVSlider(1).Value
        .EV_Def = EVSlider(2).Value
        .EV_Spd = EVSlider(3).Value
        .EV_SAtk = EVSlider(4).Value
        .EV_SDef = EVSlider(5).Value
        .NatureNum = cmbNature.ListIndex
        .Level = hscLevel.Value
        .MaxHP = GetAdvHP(.BaseHP, .DV_HP, .EV_HP, .Level)
        .HP = .MaxHP
        .Attack = GetAdvStat(.BaseAttack, .DV_Atk, .EV_Atk, .Level, Nature(.NatureNum).StatChg(1))
        .Defense = GetAdvStat(.BaseDefense, .DV_Def, .EV_Def, .Level, Nature(.NatureNum).StatChg(2))
        .Speed = GetAdvStat(.BaseSpeed, .DV_Spd, .EV_Spd, .Level, Nature(.NatureNum).StatChg(3))
        .SpecialAttack = GetAdvStat(.BaseSAttack, .DV_SAtk, .EV_SAtk, .Level, Nature(.NatureNum).StatChg(4))
        .SpecialDefense = GetAdvStat(.BaseSDefense, .DV_SDef, .EV_SDef, .Level, Nature(.NatureNum).StatChg(5))
        lblValue(0).Caption = .HP
        lblValue(1).Caption = .Attack
        lblValue(2).Caption = .Defense
        lblValue(3).Caption = .Speed
        lblValue(4).Caption = .SpecialAttack
        lblValue(5).Caption = .SpecialDefense
        lblHiddenPower.Caption = "Hidden Power: " & HiddenPowerStrengthAdv(ExpertPKMN)
        Processing = True
        cmbHiddenPower.ListIndex = HiddenPowerTypeAdv(ExpertPKMN) - 2
        Processing = False
    End With
End Sub

Private Sub txtCoverup_KeyPress(KeyAscii As Integer)
    Dim Temp As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdSave_Click
    End If
End Sub

Private Sub txtCoverup_LostFocus()
    If txtCoverup.Visible And Me.ActiveControl.Name <> "cmdSave" Then
        txtCoverup.Visible = False
        txtCoverup.Text = ""
        Preset.Enabled = True
        Call Combo2_Change
    End If
End Sub

Private Sub txtLevel_Change()
    If (txtLevel.Text <> "" And txtLevel.Text <> "0" And txtLevel.Text <> "00" And Val(txtLevel.Text) = 0) Or Val(txtLevel.Text) > 100 Then txtLevel.Text = "100"
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLevel_LostFocus()
    If Val(txtLevel.Text) = 0 Or Val(txtLevel.Text) > 100 Then txtLevel.Text = "100"
    hscLevel.Value = Val(txtLevel)
End Sub
