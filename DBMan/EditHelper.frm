VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EditHelp 
   Caption         =   "Database Editor"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "EditHelper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   7815
   Begin MSAdodcLib.Adodc PokeDataControl 
      Height          =   375
      Left            =   120
      Top             =   7440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text11 
      DataField       =   "Offset"
      DataSource      =   "PokeDataControl"
      Height          =   285
      Left            =   120
      TabIndex        =   129
      Text            =   "0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "$"
      Height          =   375
      Left            =   120
      TabIndex        =   127
      Top             =   7920
      Width           =   255
   End
   Begin MSComctlLib.ImageCombo Col2 
      Height          =   330
      Left            =   2880
      TabIndex        =   75
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo Col1 
      Height          =   330
      Left            =   1080
      TabIndex        =   74
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo Att1 
      Height          =   330
      Left            =   1080
      TabIndex        =   73
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Frame Frame2 
      Caption         =   "Special"
      Height          =   1095
      Left            =   4800
      TabIndex        =   56
      Top             =   720
      Width           =   2775
      Begin VB.CheckBox Check5 
         Caption         =   "ADV"
         DataField       =   "ADV"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Left            =   1320
         TabIndex        =   67
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "GSC"
         DataField       =   "GSC"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Left            =   1320
         TabIndex        =   66
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "RBY"
         DataField       =   "RBY"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Left            =   1320
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Legendary"
         DataField       =   "Legendary"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Uber"
         DataField       =   "Uber"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      DataField       =   "AdvNumber"
      DataSource      =   "PokeDataControl"
      Height          =   315
      Left            =   120
      TabIndex        =   48
      Text            =   "1"
      Top             =   1560
      Width           =   570
   End
   Begin VB.TextBox Text2 
      DataField       =   "GSCNumber"
      DataSource      =   "PokeDataControl"
      Height          =   315
      Left            =   120
      TabIndex        =   46
      Text            =   "1"
      Top             =   960
      Width           =   570
   End
   Begin VB.CommandButton Command5 
      Caption         =   "#"
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   7920
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Common TMs"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   38
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save Current"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base Stats"
      Height          =   975
      Left            =   1080
      TabIndex        =   21
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox Text1 
         DataField       =   "SpecialRBY"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   4440
         TabIndex        =   44
         Text            =   "1"
         Top             =   480
         Width           =   645
      End
      Begin VB.TextBox SDEF 
         DataField       =   "SpecialDefense"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Text            =   "1"
         Top             =   480
         Width           =   600
      End
      Begin VB.TextBox SATK 
         DataField       =   "SpecialAttack"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Text            =   "1"
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox SPD 
         DataField       =   "Speed"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Text            =   "1"
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox DEF 
         DataField       =   "Defense"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Text            =   "1"
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox ATK 
         DataField       =   "Attack"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox MaxHP 
         DataField       =   "HP"
         DataSource      =   "PokeDataControl"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "1"
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "RBY Sp."
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Sp.Def"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Sp.Atk"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Speed"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Defense"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Attack"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "HP"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox L5Moves 
      DataField       =   "BornWith"
      DataSource      =   "PokeDataControl"
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Text            =   "1"
      Top             =   360
      Width           =   585
   End
   Begin MSComctlLib.Slider PercentFemale 
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   7680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      Min             =   -1
      Max             =   16
      TickStyle       =   1
   End
   Begin VB.TextBox Name 
      DataField       =   "Name"
      DataSource      =   "PokeDataControl"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox PokeNum 
      DataField       =   "Number"
      DataSource      =   "PokeDataControl"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   16
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "RBY Move"
      TabPicture(0)   =   "EditHelper.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "RBY TMs"
      TabPicture(1)   =   "EditHelper.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1(5)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "GSC Move"
      TabPicture(2)   =   "EditHelper.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1(0)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "GSC  TMs"
      TabPicture(3)   =   "EditHelper.frx":27F6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView1(1)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "GSC Egg"
      TabPicture(4)   =   "EditHelper.frx":2812
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView1(2)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "GSC Tut."
      TabPicture(5)   =   "EditHelper.frx":282E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "MTutor(2)"
      Tab(5).Control(1)=   "MTutor(1)"
      Tab(5).Control(2)=   "MTutor(0)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "GSC Spec."
      TabPicture(6)   =   "EditHelper.frx":284A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListView1(4)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Adv Move"
      TabPicture(7)   =   "EditHelper.frx":2866
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "ListView1(6)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Adv TMs"
      TabPicture(8)   =   "EditHelper.frx":2882
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "ListView1(7)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Adv Egg"
      TabPicture(9)   =   "EditHelper.frx":289E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "ListView1(8)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Adv Tutor"
      TabPicture(10)  =   "EditHelper.frx":28BA
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "ListView1(10)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Adv Spec."
      TabPicture(11)  =   "EditHelper.frx":28D6
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "ListView1(9)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Evolution"
      TabPicture(12)  =   "EditHelper.frx":28F2
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Evolution(0)"
      Tab(12).Control(1)=   "Evolution(1)"
      Tab(12).Control(2)=   "Evolution(2)"
      Tab(12).Control(3)=   "Evolution(3)"
      Tab(12).Control(4)=   "Evolution(4)"
      Tab(12).Control(5)=   "EvoStage(0)"
      Tab(12).Control(6)=   "EvoStage(1)"
      Tab(12).Control(7)=   "EvoStage(2)"
      Tab(12).Control(8)=   "EvoStage(3)"
      Tab(12).Control(9)=   "EvoStage(4)"
      Tab(12).Control(10)=   "EvoStage(5)"
      Tab(12).Control(11)=   "Command6"
      Tab(12).Control(12)=   "EvoMethPick(0)"
      Tab(12).Control(13)=   "EvoMethPick(1)"
      Tab(12).Control(14)=   "EvoMethPick(2)"
      Tab(12).Control(15)=   "EvoMethPick(3)"
      Tab(12).Control(16)=   "EvoMethPick(4)"
      Tab(12).Control(17)=   "EvoMethPick(5)"
      Tab(12).Control(18)=   "Label20(0)"
      Tab(12).Control(19)=   "Label21(0)"
      Tab(12).Control(20)=   "Label20(1)"
      Tab(12).Control(21)=   "Label21(1)"
      Tab(12).Control(22)=   "Label20(2)"
      Tab(12).Control(23)=   "Label21(2)"
      Tab(12).Control(24)=   "Label20(3)"
      Tab(12).Control(25)=   "Label21(3)"
      Tab(12).Control(26)=   "Label20(4)"
      Tab(12).Control(27)=   "Label21(4)"
      Tab(12).Control(28)=   "Label21(5)"
      Tab(12).ControlCount=   29
      TabCaption(13)  =   "Pokédex"
      TabPicture(13)  =   "EditHelper.frx":290E
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Text10"
      Tab(13).Control(1)=   "Text9"
      Tab(13).Control(2)=   "Text8"
      Tab(13).Control(3)=   "Text7"
      Tab(13).Control(4)=   "Text6"
      Tab(13).Control(5)=   "Text5"
      Tab(13).Control(6)=   "Text4"
      Tab(13).Control(7)=   "Label28"
      Tab(13).Control(8)=   "Label27"
      Tab(13).Control(9)=   "Label26"
      Tab(13).Control(10)=   "Label25"
      Tab(13).Control(11)=   "Label24"
      Tab(13).Control(12)=   "Label23"
      Tab(13).Control(13)=   "Label22"
      Tab(13).ControlCount=   14
      TabCaption(14)  =   "LF Moves"
      TabPicture(14)  =   "EditHelper.frx":292A
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "ListView1(11)"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "Levels"
      TabPicture(15)  =   "EditHelper.frx":2946
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "SSTab2"
      Tab(15).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   133
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "RBY"
         TabPicture(0)   =   "EditHelper.frx":2962
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LListView(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LevelList(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "GSC"
         TabPicture(1)   =   "EditHelper.frx":297E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LListView(1)"
         Tab(1).Control(1)=   "LevelList(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Advance"
         TabPicture(2)   =   "EditHelper.frx":299A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LListView(2)"
         Tab(2).Control(1)=   "LevelList(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "L/F Only"
         TabPicture(3)   =   "EditHelper.frx":29B6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "LListView(3)"
         Tab(3).Control(1)=   "LevelList(3)"
         Tab(3).ControlCount=   2
         Begin VB.ListBox LevelList 
            Height          =   2790
            Index           =   3
            Left            =   -74880
            TabIndex        =   143
            Top             =   480
            Width           =   735
         End
         Begin VB.ListBox LevelList 
            Height          =   2790
            Index           =   2
            Left            =   -74880
            TabIndex        =   138
            Top             =   480
            Width           =   735
         End
         Begin VB.ListBox LevelList 
            Height          =   2790
            Index           =   1
            Left            =   -74880
            TabIndex        =   136
            Top             =   480
            Width           =   735
         End
         Begin VB.ListBox LevelList 
            Height          =   2790
            Index           =   0
            Left            =   120
            TabIndex        =   134
            Top             =   480
            Width           =   735
         End
         Begin MSComctlLib.ListView LListView 
            Height          =   2895
            Index           =   0
            Left            =   960
            TabIndex        =   135
            Top             =   480
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5106
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView LListView 
            Height          =   2895
            Index           =   1
            Left            =   -74040
            TabIndex        =   137
            Top             =   480
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5106
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView LListView 
            Height          =   2895
            Index           =   2
            Left            =   -74040
            TabIndex        =   139
            Top             =   480
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5106
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView LListView 
            Height          =   2895
            Index           =   3
            Left            =   -74040
            TabIndex        =   144
            Top             =   480
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5106
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "Types"
            SmallIcons      =   "Types"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.TextBox Evolution 
         DataField       =   "Evo1"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   0
         Left            =   -74880
         MaxLength       =   3
         TabIndex        =   110
         Top             =   1020
         Width           =   465
      End
      Begin VB.TextBox Evolution 
         DataField       =   "Evo2"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   1
         Left            =   -71040
         MaxLength       =   3
         TabIndex        =   109
         Top             =   1020
         Width           =   480
      End
      Begin VB.TextBox Evolution 
         DataField       =   "Evo3"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   2
         Left            =   -74880
         MaxLength       =   3
         TabIndex        =   108
         Top             =   1740
         Width           =   480
      End
      Begin VB.TextBox Evolution 
         DataField       =   "Evo4"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   3
         Left            =   -71040
         MaxLength       =   3
         TabIndex        =   107
         Top             =   1740
         Width           =   480
      End
      Begin VB.TextBox Evolution 
         DataField       =   "Evo5"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   4
         Left            =   -74880
         MaxLength       =   3
         TabIndex        =   106
         Top             =   2460
         Width           =   480
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "Stage1"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   0
         Left            =   -72120
         TabIndex        =   105
         Top             =   1020
         Width           =   360
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "Stage2"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   1
         Left            =   -68280
         TabIndex        =   104
         Top             =   1020
         Width           =   360
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "Stage3"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   2
         Left            =   -72120
         TabIndex        =   103
         Top             =   1740
         Width           =   360
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "Stage4"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   3
         Left            =   -68280
         TabIndex        =   102
         Top             =   1740
         Width           =   360
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "Stage5"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   4
         Left            =   -72120
         TabIndex        =   101
         Top             =   2460
         Width           =   360
      End
      Begin VB.TextBox EvoStage 
         DataField       =   "MyStage"
         DataSource      =   "PokeDataControl"
         Height          =   285
         Index           =   5
         Left            =   -68280
         TabIndex        =   100
         Top             =   2460
         Width           =   360
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Scan &Evolutions"
         Height          =   375
         Left            =   -69360
         TabIndex        =   99
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         DataField       =   "PokedexSapphire"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   3420
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         DataField       =   "PokedexRuby"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -71160
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   2580
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         DataField       =   "PokedexCrystal"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   2580
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         DataField       =   "PokedexSilver"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -71160
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   1740
         Width           =   3615
      End
      Begin VB.TextBox Text6 
         DataField       =   "PokedexGold"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   1740
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         DataField       =   "PokedexYellow"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -71160
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   900
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         DataField       =   "PokedexRB"
         DataSource      =   "PokeDataControl"
         Height          =   525
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   900
         Width           =   3615
      End
      Begin VB.CheckBox MTutor 
         Caption         =   "Thunderbolt"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   42
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CheckBox MTutor 
         Caption         =   "Ice Beam"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   41
         Top             =   1140
         Width           =   1575
      End
      Begin VB.CheckBox MTutor 
         Caption         =   "Flamethrower"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   40
         Top             =   780
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   6
         Left            =   -74880
         TabIndex        =   61
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   7
         Left            =   -74880
         TabIndex        =   62
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   8
         Left            =   -74880
         TabIndex        =   63
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   3
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   5
         Left            =   -74880
         TabIndex        =   92
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   0
         Left            =   -74880
         TabIndex        =   93
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   1
         Left            =   -74880
         TabIndex        =   94
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   2
         Left            =   -74880
         TabIndex        =   95
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   4
         Left            =   -74880
         TabIndex        =   96
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   9
         Left            =   -74880
         TabIndex        =   97
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   0
         Left            =   -74040
         TabIndex        =   98
         Top             =   1020
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   1
         Left            =   -70200
         TabIndex        =   111
         Top             =   1020
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   2
         Left            =   -74040
         TabIndex        =   112
         Top             =   1740
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   3
         Left            =   -70200
         TabIndex        =   113
         Top             =   1740
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   4
         Left            =   -74040
         TabIndex        =   114
         Top             =   2460
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo EvoMethPick 
         Height          =   330
         Index           =   5
         Left            =   -70200
         TabIndex        =   115
         Top             =   2460
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   10
         Left            =   -74880
         TabIndex        =   130
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Index           =   11
         Left            =   -74880
         TabIndex        =   131
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "Types"
         SmallIcons      =   "Types"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Form 1"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   126
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "EvoM1"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   0
         Left            =   -74040
         TabIndex        =   125
         Top             =   780
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Form 2"
         Height          =   255
         Index           =   1
         Left            =   -71040
         TabIndex        =   124
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "EvoM2"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   1
         Left            =   -70200
         TabIndex        =   123
         Top             =   780
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Form 3"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   122
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "EvoM3"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   2
         Left            =   -74040
         TabIndex        =   121
         Top             =   1500
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Form 4"
         Height          =   255
         Index           =   3
         Left            =   -71040
         TabIndex        =   120
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "EvoM4"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   3
         Left            =   -70200
         TabIndex        =   119
         Top             =   1500
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Form 5"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   118
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "EvoM5"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   4
         Left            =   -74040
         TabIndex        =   117
         Top             =   2220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Evolves By"
         DataField       =   "MyMethod"
         DataSource      =   "PokeDataControl"
         Height          =   255
         Index           =   5
         Left            =   -70200
         TabIndex        =   116
         Top             =   2220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Sapphire"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   3180
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruby"
         Height          =   255
         Left            =   -71160
         TabIndex        =   87
         Top             =   2340
         Width           =   3375
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Crystal"
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
         Top             =   2340
         Width           =   3375
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Silver"
         Height          =   255
         Left            =   -71160
         TabIndex        =   83
         Top             =   1500
         Width           =   3375
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Gold"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   1500
         Width           =   3375
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         Height          =   255
         Left            =   -71160
         TabIndex        =   79
         Top             =   660
         Width           =   3375
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Red/Blue"
         Height          =   255
         Left            =   -74880
         TabIndex        =   77
         Top             =   660
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   7920
      Width           =   1215
   End
   Begin MSComctlLib.ImageCombo Type2 
      Height          =   330
      Left            =   4800
      TabIndex        =   3
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "Types"
   End
   Begin MSComctlLib.ImageList Types 
      Left            =   7080
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":29D2
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":2F6C
            Key             =   ""
            Object.Tag             =   "Fire"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":3506
            Key             =   ""
            Object.Tag             =   "Water"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":3AA0
            Key             =   ""
            Object.Tag             =   "Electric"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":403A
            Key             =   ""
            Object.Tag             =   "Grass"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":45D4
            Key             =   ""
            Object.Tag             =   "Ice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":4B6E
            Key             =   ""
            Object.Tag             =   "Fighting"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":5108
            Key             =   ""
            Object.Tag             =   "Poison"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":56A2
            Key             =   ""
            Object.Tag             =   "Ground"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":5C3C
            Key             =   ""
            Object.Tag             =   "Flying"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":61D6
            Key             =   ""
            Object.Tag             =   "Psychic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":6770
            Key             =   ""
            Object.Tag             =   "Bug"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":6D0A
            Key             =   ""
            Object.Tag             =   "Rock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":72A4
            Key             =   ""
            Object.Tag             =   "Ghost"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":783E
            Key             =   ""
            Object.Tag             =   "Dragon"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":7DD8
            Key             =   ""
            Object.Tag             =   "Dark"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditHelper.frx":8372
            Key             =   ""
            Object.Tag             =   "Steel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo Type1 
      Height          =   330
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "Types"
   End
   Begin MSComctlLib.ImageCombo Att2 
      Height          =   330
      Left            =   2880
      TabIndex        =   90
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label MovesTemp 
      Caption         =   "LFLevels"
      DataField       =   "LFLevels"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   15
      Left            =   7920
      TabIndex        =   145
      Top             =   9240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "ADVLevels"
      DataField       =   "ADVLevels"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   142
      Top             =   8760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "GSCLevels"
      DataField       =   "GSCLevels"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   13
      Left            =   7920
      TabIndex        =   141
      Top             =   8280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "RBYLevels"
      DataField       =   "RBYLevels"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   12
      Left            =   7920
      TabIndex        =   140
      Top             =   7800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "LF Moves"
      DataField       =   "LF Only"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   11
      Left            =   7920
      TabIndex        =   132
      Top             =   7320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Offset GFX"
      Height          =   255
      Left            =   120
      TabIndex        =   128
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Adv Moves"
      DataField       =   "Advance Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   72
      Top             =   4920
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Adv TMs"
      DataField       =   "Advance TMs"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   71
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Adv Breeding"
      DataField       =   "Advance Breeding"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   70
      Top             =   5880
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Adv Special"
      DataField       =   "Advance Special"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   9
      Left            =   7920
      TabIndex        =   69
      Top             =   6360
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Adv Tutor"
      DataField       =   "Advance Tutor"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   68
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "RBY TMs"
      DataField       =   "R/B/Y Machines"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   64
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Attrib2 
      Caption         =   "Attrib2"
      DataField       =   "Attribute2"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   9840
      TabIndex        =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Attrib1 
      Caption         =   "Attrib1"
      DataField       =   "Attribute1"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   9840
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Color2 
      Caption         =   "Color2"
      DataField       =   "Color2"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   8880
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Color1 
      Caption         =   "Color1"
      DataField       =   "Color1"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   8880
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Attribute 2"
      Height          =   255
      Left            =   2880
      TabIndex        =   53
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Attribute 1"
      Height          =   255
      Left            =   1080
      TabIndex        =   52
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Color 2"
      Height          =   255
      Left            =   2880
      TabIndex        =   51
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Color 1"
      Height          =   255
      Left            =   1080
      TabIndex        =   50
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Adv PDex"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GSC PDex"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   720
      Width           =   855
   End
   Begin VB.Label MoveTutor 
      Caption         =   "Move Tutor"
      DataField       =   "Move Tutor"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   7920
      TabIndex        =   43
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Special Moves"
      DataField       =   "Special Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "R/B/Y Moves"
      DataField       =   "R/B/Y Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Breeding Moves"
      DataField       =   "Breeding Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Machine Moves"
      DataField       =   "Machine Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label MovesTemp 
      Caption         =   "Moves"
      DataField       =   "Moves"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   31
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label PercentFemaleTemp 
      Caption         =   "%Female"
      DataField       =   "Percent Female"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   7920
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Type2Temp 
      Caption         =   "Type2"
      DataField       =   "Type2"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   7920
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Type1Temp 
      Caption         =   "Type1"
      DataField       =   "Type1"
      DataSource      =   "PokeDataControl"
      Height          =   255
      Left            =   7920
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Moves @ L5"
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.Label SliderLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "% Born Female"
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type2"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type1"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "EditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SurpressUpdate As Boolean
Private PokeRecords As ADODB.Recordset
Dim LevelMovesVer() As Boolean
Dim LevelMovesLevel() As Byte

Private Sub AdvTutorTemp_Change()
    Dim X As Integer
    Dim M As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    
    If Not SurpressUpdate Then
        'SurpressUpdate = True
        'For X = 0 To AdvTutor.Count
        '    AdvTutor(X).Value = 0
        'Next
       '
    'Else
        'SurpressUpdate = False
    End If
End Sub

Private Sub Attrib1_Change()
    If Not SurpressUpdate Then
        Att1.SelectedItem = Att1.ComboItems(Val(Attrib1.Caption) + 1)
    Else
        SurpressUpdate = False
    End If
End Sub

Private Sub Attrib2_Change()
    If Not SurpressUpdate Then
        Att2.SelectedItem = Att2.ComboItems(Val(Attrib2.Caption) + 1)
    Else
        SurpressUpdate = False
    End If
End Sub

Private Sub Color1_Change()
    If Not SurpressUpdate Then
        Col1.SelectedItem = Col1.ComboItems(Val(Color1.Caption) + 1)
    Else
        SurpressUpdate = False
    End If
End Sub

Private Sub Color2_Change()
    If Not SurpressUpdate Then
        Col2.SelectedItem = Col2.ComboItems(Val(Color2.Caption) + 1)
    Else
        SurpressUpdate = False
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    PokeDataControl.Recordset.AddNew
    SSTab1.Tab = 0
    Command4.Enabled = False
End Sub

Private Sub Command3_Click()
    Dim CurrentSpot As Integer
    
    CurrentSpot = PokeDataControl.Recordset.AbsolutePosition
    PokeDataControl.Recordset.Save
    PokeDataControl.Recordset.AbsolutePosition = CurrentSpot
End Sub

Private Sub Command4_Click()
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Integer
    
    ListView1(1).ListItems("#8").Checked = True
    ListView1(1).ListItems("#38").Checked = True
    ListView1(1).ListItems("#48").Checked = True
    ListView1(1).ListItems("#59").Checked = True
    ListView1(1).ListItems("#75").Checked = True
    ListView1(1).ListItems("#91").Checked = True
    ListView1(1).ListItems("#149").Checked = True
    ListView1(1).ListItems("#163").Checked = True
    ListView1(1).ListItems("#164").Checked = True
    ListView1(1).ListItems("#190").Checked = True
    ListView1(1).ListItems("#195").Checked = True
    ListView1(1).ListItems("#218").Checked = True
    ListView1(1).ListItems("#235").Checked = True
    SurpressUpdate = True
    Call ListView1_ItemCheck(1, ListView1(1).ListItems("#8"))
End Sub

Private Sub Command5_Click()
    Dim X As Integer
    Dim Temp As String
    
    Temp = InputBox("Jump to which rcord number?", "Jump To")
    X = Val(Temp)
    PokeDataControl.Recordset.AbsolutePosition = X
End Sub

Private Sub Command6_Click()
    Call DoEvoScan(Val(PokeNum.Text))
End Sub

Private Sub Command7_Click()
    PokeChoose.Show
    Me.Enabled = False
    PokeChoose.PokeCombo.SetFocus
End Sub
Public Sub SetPoke(n As Integer)
    PokeDataControl.Recordset.AbsolutePosition = n
End Sub


Private Sub Evolution_Change(Index As Integer)
    If Label21(Index).Caption = "" And Evolution(Index).Text <> "" Then Label21(Index).Caption = "0"
End Sub

Private Sub EvoMethPick_Click(Index As Integer)
    SurpressUpdate = True
    Label21(Index).Caption = EvoMethPick(Index).SelectedItem.Index - 1
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer
    Dim TMOrder(50) As Integer
    Dim HMOrder(8) As Integer
    Dim Sorted() As String
    
    ReDim LevelMovesVer(1 To UBound(Moves), 0 To 3) As Boolean
    ReDim LevelMovesLevel(1 To UBound(Moves), 0 To 3) As Byte
    
    For X = 1 To 17
        Type1.ComboItems.Add X, , Element(X), X
        Type2.ComboItems.Add X, , Element(X), X
    Next
    Type2.ComboItems.Add 18, , "None"
    For X = 0 To UBound(EvoMethod)
        For Y = 0 To 5
            EvoMethPick(Y).ComboItems.Add , , EvoMethod(X)
            EvoMethPick(Y).SelectedItem = EvoMethPick(Y).ComboItems(1)
        Next
    Next
    For X = 0 To UBound(AttributeText)
        Att1.ComboItems.Add , , AttributeText(X)
        Att2.ComboItems.Add , , AttributeText(X)
    Next
    For X = 0 To UBound(ColorText)
        Col1.ComboItems.Add , , ColorText(X)
        Col2.ComboItems.Add , , ColorText(X)
    Next
    For Y = 0 To 11
        Select Case Y
            Case 0, 2, 4
                For X = 1 To UBound(Moves)
                    If Moves(X).GSCMove Then ListView1(Y).ListItems.Add , "#" & X, Moves(X).Name, Moves(X).Type, Moves(X).Type
                Next
            Case 1
                For X = 1 To 50
                    TMOrder(X) = 0
                Next
                For X = 1 To 7
                    HMOrder(X) = 0
                Next
                For X = 1 To UBound(Moves)
                    If Moves(X).NewTM <> "" Then
                        Select Case Left(Moves(X).NewTM, 2)
                            Case "TM"
                                TMOrder(Val(Right(Moves(X).NewTM, 2))) = X
                            Case "HM"
                                HMOrder(Val(Right(Moves(X).NewTM, 2))) = X
                        End Select
                    End If
                Next
                For X = 1 To 50
                    If TMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & TMOrder(X), Moves(TMOrder(X)).NewTM & " - " & Moves(TMOrder(X)).Name, Moves(TMOrder(X)).Type, Moves(TMOrder(X)).Type
                    End If
                Next
                For X = 1 To 7
                    If HMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & HMOrder(X), Moves(HMOrder(X)).NewTM & " - " & Moves(HMOrder(X)).Name, Moves(HMOrder(X)).Type, Moves(HMOrder(X)).Type
                    End If
                Next
            Case 3
                For X = 1 To UBound(Moves)
                    If Moves(X).RBYMove Then ListView1(Y).ListItems.Add , "#" & X, Moves(X).Name, Moves(X).Type, Moves(X).Type
                Next
            Case 5
                For X = 1 To 50
                    TMOrder(X) = 0
                Next
                For X = 1 To 7
                    HMOrder(X) = 0
                Next
                For X = 1 To UBound(Moves)
                    If Moves(X).OldTM <> "" Then
                        Select Case Left(Moves(X).OldTM, 2)
                            Case "TM"
                                TMOrder(Val(Right(Moves(X).OldTM, 2))) = X
                            Case "HM"
                                HMOrder(Val(Right(Moves(X).OldTM, 2))) = X
                        End Select
                    End If
                Next
                For X = 1 To 50
                    If TMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & TMOrder(X), Moves(TMOrder(X)).OldTM & " - " & Moves(TMOrder(X)).Name, Moves(TMOrder(X)).Type, Moves(TMOrder(X)).Type
                    End If
                Next
                For X = 1 To 7
                    If HMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & HMOrder(X), Moves(HMOrder(X)).OldTM & " - " & Moves(HMOrder(X)).Name, Moves(HMOrder(X)).Type, Moves(HMOrder(X)).Type
                    End If
                Next
            Case 6, 8, 9, 11
                For X = 1 To UBound(Moves)
                    If Moves(X).AdvMove Then ListView1(Y).ListItems.Add , "#" & X, Moves(X).Name, Moves(X).Type, Moves(X).Type
                Next
            Case 7
                For X = 1 To 50
                    TMOrder(X) = 0
                Next
                For X = 1 To 8
                    HMOrder(X) = 0
                Next
                For X = 1 To UBound(Moves)
                    If Moves(X).ADVTM <> "" Then
                        Select Case Left(Moves(X).ADVTM, 2)
                            Case "TM"
                                TMOrder(Val(Right(Moves(X).ADVTM, 2))) = X
                            Case "HM"
                                HMOrder(Val(Right(Moves(X).ADVTM, 2))) = X
                        End Select
                    End If
                Next
                For X = 1 To 50
                    If TMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & TMOrder(X), Moves(TMOrder(X)).ADVTM & " - " & Moves(TMOrder(X)).Name, Moves(TMOrder(X)).Type, Moves(TMOrder(X)).Type
                    End If
                Next
                For X = 1 To 8
                    If HMOrder(X) > 0 Then
                        ListView1(Y).ListItems.Add , "#" & HMOrder(X), Moves(HMOrder(X)).ADVTM & " - " & Moves(HMOrder(X)).Name, Moves(HMOrder(X)).Type, Moves(HMOrder(X)).Type
                    End If
                Next
            Case 10
                For X = 1 To UBound(Moves)
                    Select Case X
                        Case 258, 292, 285, 119, 118, 179, 167, 34, 196, 124, 231, 213, 52, 122, 48, 60, 19, 222
                            ListView1(Y).ListItems.Add , "#" & X, Moves(X).Name, Moves(X).Type, Moves(X).Type
                    End Select
                Next
        End Select
        ListView1(Y).Sorted = True
    Next
    
    For X = 0 To 3
        LevelList(X).Clear
        LevelList(X).AddItem "Pre", 0
        For Y = 1 To 100
            LevelList(X).AddItem Y, Y
        Next
    Next
    
    On Error GoTo BadPassword
    With PokeDataControl
        '.Recordset.Close
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SlashPath & "PokeDB.mdb" & ";Mode=ReadWrite;Persist Security Info=False;Jet OLEDB"
        .RecordSource = "Pokemon"
        .Password = "ginyu4ce"
        .Refresh
        .Recordset.StayInSync = True
    End With
    'PokeDataControl.DatabaseName = SlashPath & "PokeDB.mdb"
    'PokeDataControl.RecordSource = "PKMNSorted"
    'PokeDataControl.UpdateControls
    
    Call LoadPKMNData
    PokeChoose.Hide
    ReDim Sorted(1 To UBound(BasePKMN))
    For X = 1 To UBound(Sorted)
        Sorted(X) = BasePKMN(X).Name & " - " & X
    Next X
    Call SortStringArray(Sorted)
    For X = 1 To UBound(Sorted)
        PokeChoose.PokeCombo.AddItem Sorted(X)
    Next X

    Exit Sub
BadPassword:
    MsgBox "Unable to open database.", vbCritical, "Error"
End Sub

Private Sub Form_Resize()
    Dim X As Integer
    
    If EditHelp.WindowState <> vbMinimized Then
        If EditHelp.WindowState <> vbMaximized Then
            If EditHelp.Width < 7935 Then EditHelp.Width = 7935
            If EditHelp.Height < 9045 Then EditHelp.Height = 9045
        End If
        PokeDataControl.Top = EditHelp.Height - 1485
        Label5.Top = EditHelp.Height - 1485
        SliderLabel.Top = EditHelp.Height - 1485
        Command2.Top = EditHelp.Height - 1005
        PercentFemale.Top = EditHelp.Height - 1245
        Command3.Top = EditHelp.Height - 1485
        Command1.Top = EditHelp.Height - 1005
        Command4.Top = EditHelp.Height - 1005
        Command5.Top = EditHelp.Height - 1005
        Command7.Top = EditHelp.Height - 1005
        Command3.Left = EditHelp.Width - 1470
        Command1.Left = EditHelp.Width - 1470
        SSTab1.Width = EditHelp.Width - 135
        SSTab1.Height = EditHelp.Height - 5010
        SSTab2.Width = SSTab1.Width - 360
        SSTab2.Height = SSTab1.Height - 840
        For X = 0 To 11
            ListView1(X).Width = SSTab1.Width - 360
            ListView1(X).Height = SSTab1.Height - 840
        Next
        For X = 0 To 3
            LListView(X).Height = SSTab2.Height - 600
            LListView(X).Width = SSTab2.Width - 1080
            LevelList(X).Height = SSTab2.Height - 705
        Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PokeData.Close
    PokeData.Open
    DBManMain.Command1.Enabled = True
    Unload PokeChoose
End Sub

Private Sub Label21_Change(Index As Integer)
    If Not SurpressUpdate Then
        EvoMethPick(Index).SelectedItem = EvoMethPick(Index).ComboItems(Val(Label21(Index).Caption) + 1)
    Else
        SurpressUpdate = False
    End If
End Sub

Private Sub LevelList_Click(Index As Integer)
    Dim X As Integer
    
    On Error Resume Next
    For X = 1 To LListView(Index).ListItems.Count
        LListView(Index).ListItems(X).Checked = False
    Next
    For X = 1 To UBound(Moves)
        If LevelMovesLevel(X, Index) = LevelList(Index).ListIndex Then LListView(Index).ListItems("#" & X).Checked = True
    Next
End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Integer
    Dim Y As Integer
    
    SurpressUpdate = True
    Temp = ""
    For X = 1 To ListView1(Index).ListItems.Count
        If ListView1(Index).ListItems(X).Checked Then
            Temp2 = ListView1(Index).ListItems(X).Key
            Temp2 = Right(Temp2, Len(Temp2) - 1)
            Temp = Temp + Temp2 + ","
        End If
    Next X
    MovesTemp(Index).Caption = Temp
End Sub

Private Sub LListView_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    Dim TempString As String
    
    TempString = ""
    If Item.Checked Then
        LevelMovesLevel(Right(Item.Key, Len(Item.Key) - 1), Index) = LevelList(Index).ListIndex
    Else
        LevelMovesLevel(Right(Item.Key, Len(Item.Key) - 1), Index) = 0
    End If
    For X = 1 To UBound(Moves)
        If LevelMovesVer(X, Index) Then
            TempString = TempString & Format(X, "000") & Format(LevelMovesLevel(X, Index), "000")
        End If
    Next
    SurpressUpdate = True
    MovesTemp(Index + 12).Caption = TempString
End Sub

Private Sub MovesTemp_Change(Index As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim Temp As String
    
    If SurpressUpdate Then
        SurpressUpdate = False
        Exit Sub
    End If
    If Index < 12 Then Call RefreshChecks(Index)
    If Index = 0 Or Index = 3 Or Index = 6 Then Call RefillLevelList
    If Index >= 12 Then Call RefreshLevelValues(Index)
End Sub


Private Sub MoveTutor_Change()
    Dim Temp As Integer
    
    Temp = Val(MoveTutor.Caption)
    If Temp - 4 >= 0 Then
        MTutor(0).Value = 1
        Temp = Temp - 4
    Else
        MTutor(0).Value = 0
    End If
    If Temp - 2 >= 0 Then
        MTutor(1).Value = 1
        Temp = Temp - 2
    Else
        MTutor(1).Value = 0
    End If
    If Temp - 1 >= 0 Then
        MTutor(2).Value = 1
        Temp = Temp - 1
    Else
        MTutor(2).Value = 0
    End If
    
End Sub

Private Sub MTutor_Click(Index As Integer)
    Dim Temp As Integer
    
    Temp = 0
    If MTutor(0).Value = 1 Then Temp = Temp + 4
    If MTutor(1).Value = 1 Then Temp = Temp + 2
    If MTutor(2).Value = 1 Then Temp = Temp + 1
    MoveTutor.Caption = Temp
End Sub

Private Sub PercentFemale_Change()
    Dim X As Integer
    
    PercentFemaleTemp.Caption = PercentFemale.Value
    If PercentFemale.Value = -1 Then
        SliderLabel.Caption = "Genderless"
    Else
        X = (PercentFemale.Value * 100) / 16
        SliderLabel.Caption = X & "%"
    End If
End Sub

Private Sub PercentFemaleTemp_Change()
    If PercentFemale.Value = Val(PercentFemaleTemp.Caption) Then Exit Sub
    PercentFemale.Value = Val(PercentFemaleTemp.Caption)
End Sub

'Private Sub PokeDataControl_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    PokeDataControl.Caption = PokeDataControl.Recordset.AbsolutePosition & "/" & PokeDataControl.Recordset.RecordCount
'End Sub

Private Sub SSTab1_KeyUp(KeyCode As Integer, Shift As Integer)
    If SSTab1.Tab = 2 Then Command4.Enabled = True Else Command4.Enabled = False
End Sub


Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SSTab1.Tab = 2 Then Command4.Enabled = True Else Command4.Enabled = False
End Sub

Private Sub Type1_Click()
    Dim X As Integer
    
    For X = 1 To 17
        If Type1.ComboItems.Item(X).Selected Then Type1Temp.Caption = X
    Next X
End Sub

Private Sub Type1Temp_Change()
    If Val(Type1Temp.Caption) > 0 Then
        If Type1.ComboItems.Item(Val(Type1Temp.Caption)).Selected Then Exit Sub
    Else
        If Type1.ComboItems.Item(1).Selected Then Exit Sub
    End If
    If Val(Type1Temp.Caption) > 0 Then
        Type1.ComboItems.Item(Val(Type1Temp.Caption)).Selected = True
    Else
        Type1.ComboItems.Item(1).Selected = True
    End If
End Sub

Private Sub Type2_Click()
    Dim X As Integer
    
    If Type2.ComboItems.Item(18).Selected Then Type2Temp.Caption = "0"
    For X = 1 To 17
        If Type2.ComboItems.Item(X).Selected Then Type2Temp.Caption = X
    Next X
End Sub

Private Sub Type2Temp_Change()
    If Val(Type2Temp.Caption) > 0 Then
        If Type2.ComboItems.Item(Val(Type2Temp.Caption)).Selected Then Exit Sub
    Else
        If Type2.ComboItems.Item(18).Selected Then Exit Sub
    End If
    If Val(Type2Temp.Caption) > 0 Then
        Type2.ComboItems.Item(Val(Type2Temp.Caption)).Selected = True
    Else
        Type2.ComboItems.Item(18).Selected = True
    End If
End Sub

Private Sub DoEvoScan(ByVal Number As Integer)
    Dim EvoDataFound As Boolean
    Dim EvoMatrix(5, 6) As Integer
    Dim EvoSMatrix(5, 6) As Integer
    Dim EvoTMatrix(5, 6) As Integer
    Dim X As Integer
    Dim Y As Integer
    
    PokeDataControl.Recordset.Save
    Call LoadPKMNData
    
    For X = 1 To 5
        EvoMatrix(X, 6) = BasePKMN(Number).Evo(X)
        EvoSMatrix(X, 6) = BasePKMN(Number).Stage(X)
        EvoTMatrix(X, 6) = BasePKMN(Number).EvoM(X)
        For Y = 1 To 5
            If Y <> X Then
                EvoMatrix(X, Y) = BasePKMN(Number).Evo(Y)
                EvoSMatrix(X, Y) = BasePKMN(Number).Stage(Y)
                EvoTMatrix(X, Y) = BasePKMN(Number).EvoM(Y)
            Else
                EvoMatrix(X, Y) = Number
                EvoSMatrix(X, Y) = BasePKMN(Number).MyStage
                EvoTMatrix(X, Y) = BasePKMN(Number).MyMethod
            End If
        Next
    Next
    For X = 1 To 5
        If EvoMatrix(X, 6) > 0 Then
            With BasePKMN(EvoMatrix(X, 6))
                For Y = 1 To 5
                    .Evo(Y) = EvoMatrix(X, Y)
                    .EvoM(Y) = EvoTMatrix(X, Y)
                    .Stage(Y) = EvoSMatrix(X, Y)
                Next
                .MyMethod = BasePKMN(Number).EvoM(X)
                .MyStage = BasePKMN(Number).Stage(X)
                PokeData.Execute "UPDATE Pokemon SET Evo1=" & .Evo(1) & ", Evo2=" & .Evo(2) & ", Evo3=" & .Evo(3) & ", Evo4=" & .Evo(4) & ", Evo5=" & .Evo(5) & " WHERE Number=" & .No & ""
                PokeData.Execute "UPDATE Pokemon SET EvoM1=" & .EvoM(1) & ", EvoM2=" & .EvoM(2) & ", EvoM3=" & .EvoM(3) & ", EvoM4=" & .EvoM(4) & ", EvoM5=" & .EvoM(5) & " WHERE Number=" & .No & ""
                PokeData.Execute "UPDATE Pokemon SET Stage1=" & .Stage(1) & ", Stage2=" & .Stage(2) & ", Stage3=" & .Stage(3) & ", Stage4=" & .Stage(4) & ", Stage5=" & .Stage(5) & " WHERE Number=" & .No & ""
                PokeData.Execute "UPDATE Pokemon SET MyStage=" & .MyStage & ", MyMethod=" & .MyMethod & " WHERE Number=" & .No & ""
                PokeData.Close
                PokeData.Open
            End With
        End If
    Next
    With PokeDataControl
        .Recordset.Close
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SlashPath & "PokeDB.mdb" & ";Mode=ReadWrite;Persist Security Info=False;Jet OLEDB:Database Password=ginyu4ce"
        .RecordSource = "Pokemon"
        .Refresh
        .Recordset.StayInSync = True
        .Recordset.AbsolutePosition = Number
    End With
End Sub

Private Sub RefreshChecks(ByVal Index As Integer)
    Dim MoveStore() As Integer
    Dim X As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    
    For X = 1 To ListView1(Index).ListItems.Count
        ListView1(Index).ListItems(X).Checked = False
    Next
    If MovesTemp(Index).Caption = "" Then Exit Sub
    ReDim MoveStore(0)
    P1 = 1
    P2 = InStr(1, MovesTemp(Index).Caption, ",")
    While P2 > 0
        ReDim Preserve MoveStore(UBound(MoveStore) + 1) As Integer
        MoveStore(UBound(MoveStore)) = Mid(MovesTemp(Index).Caption, P1, P2 - P1)
        P1 = P2 + 1
        P2 = InStr(P1, MovesTemp(Index).Caption, ",")
        If P1 = P2 Then P2 = 0
    Wend
    For X = 1 To UBound(MoveStore)
        If MoveStore(X) > 0 Then ListView1(Index).ListItems("#" & MoveStore(X)).Checked = True
    Next
End Sub

Private Sub RefillLevelList()
    Dim X As Integer
    Dim Y As Integer
        
    For X = 1 To UBound(Moves)
        For Y = 0 To 3
            LevelMovesVer(X, Y) = False
        Next
    Next
    
    For X = 1 To ListView1(3).ListItems.Count
        If ListView1(3).ListItems(X).Checked Then
            LevelMovesVer(Val(Right(ListView1(3).ListItems(X).Key, Len(ListView1(3).ListItems(X).Key) - 1)), 0) = True
        End If
    Next
    For X = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(X).Checked Then
            LevelMovesVer(Val(Right(ListView1(0).ListItems(X).Key, Len(ListView1(0).ListItems(X).Key) - 1)), 1) = True
        End If
    Next
    For X = 1 To ListView1(6).ListItems.Count
        If ListView1(6).ListItems(X).Checked Then
            LevelMovesVer(Val(Right(ListView1(6).ListItems(X).Key, Len(ListView1(6).ListItems(X).Key) - 1)), 2) = True
        End If
    Next
    For X = 1 To ListView1(11).ListItems.Count
        If ListView1(11).ListItems(X).Checked Then
            LevelMovesVer(Val(Right(ListView1(6).ListItems(X).Key, Len(ListView1(6).ListItems(X).Key) - 1)), 3) = True
        End If
    Next
    
    For X = 0 To 3
        LListView(X).ListItems.Clear
        For Y = 1 To UBound(Moves)
            If LevelMovesVer(Y, X) Then
                LListView(X).ListItems.Add , "#" & Y, Moves(Y).Name, Moves(Y).Type, Moves(Y).Type
                If LevelMovesLevel(Y, X) = 0 Then LListView(X).ListItems("#" & Y).Checked = True
            End If
        Next
        LevelList(X).ListIndex = 0
    Next
End Sub

Private Sub RefreshLevelValues(ByVal MoveList As Integer)
    Dim X As Integer
    Dim MoveNum As Integer
    Dim LevNum As Integer
    Dim ModeNum As Integer
    
    ModeNum = MoveList - 12
    For X = 1 To UBound(Moves)
        LevelMovesLevel(X, ModeNum) = 0
    Next
    X = 0
    While X < Len(MovesTemp(MoveList).Caption)
        MoveNum = Val(Mid(MovesTemp(MoveList).Caption, X + 1, 3))
        LevNum = Val(Mid(MovesTemp(MoveList).Caption, X + 4, 3))
        LevelMovesLevel(MoveNum, ModeNum) = LevNum
        X = X + 6
    Wend
    Call RefillLevelList
End Sub
