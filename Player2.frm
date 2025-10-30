VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form ChallengeWindowBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Player Details"
   ClientHeight    =   6495
   ClientLeft      =   5535
   ClientTop       =   2265
   ClientWidth     =   8415
   Icon            =   "Player2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FlashTimer 
      Interval        =   1000
      Left            =   1080
      Top             =   6840
   End
   Begin VB.Timer Closer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   600
      Top             =   6840
   End
   Begin VB.Frame RuleFrame 
      Caption         =   "Rules && Modes"
      Height          =   2535
      Left            =   120
      TabIndex        =   28
      Top             =   3840
      Width           =   8175
      Begin MSComctlLib.ListView RuleList 
         Height          =   1395
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2461
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rules"
            Object.Width           =   5080
         EndProperty
      End
      Begin VB.ComboBox TerrainSelector 
         Height          =   315
         ItemData        =   "Player2.frx":1272
         Left            =   4560
         List            =   "Player2.frx":1274
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox ModeSelector 
         Height          =   315
         ItemData        =   "Player2.frx":1276
         Left            =   120
         List            =   "Player2.frx":1278
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   4335
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   7935
         TabIndex        =   30
         Top             =   1920
         Width           =   7935
         Begin VB.CommandButton CancelButton 
            Cancel          =   -1  'True
            Caption         =   "C&ancel"
            Height          =   375
            Left            =   5640
            TabIndex        =   32
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton OKButton 
            Caption         =   "&Challenge"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Label DisplayMode 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame PTInfo 
      Caption         =   "Player  && Team Info"
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   4455
      Begin CCRProgressBar6.ccrpProgressBar PowerBar 
         Height          =   255
         Left            =   2640
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         AutoCaption     =   1
         BackColor       =   0
         Caption         =   "0%"
         FillColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Smooth          =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Power"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label SpeedBar 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image imgBMode 
         Height          =   495
         Left            =   120
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame RecFrame 
      Caption         =   "Battle Record"
      Height          =   1455
      Left            =   4680
      TabIndex        =   16
      Top             =   2280
      Width           =   3615
      Begin VB.Label BattleTotalDisp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnects"
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ties"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Losses"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wins"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   2640
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame TeamFrame 
      Caption         =   "Team"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   8175
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   5
         Left            =   6840
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   9
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   5
            Left            =   120
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   4
         Left            =   5520
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   8
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   4
            Left            =   120
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   3
         Left            =   4200
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   3
            Left            =   120
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   2
         Left            =   2880
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   6
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   2
            Left            =   120
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   1
         Left            =   1560
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   1
            Left            =   120
            Top             =   120
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Index           =   0
         Left            =   240
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         Begin VB.Image Image1 
            Height          =   840
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   6720
         TabIndex        =   15
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   14
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   13
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label PokeName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1455
         Width           =   1335
      End
   End
   Begin VB.PictureBox picSwap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10560
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   10320
      Width           =   375
   End
   Begin VB.CommandButton cmdPlaceholder 
      Caption         =   "Placeholder"
      Height          =   255
      Left            =   -1000
      TabIndex        =   0
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You have been challenged!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
   End
End
Attribute VB_Name = "ChallengeWindowBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PressedOK As Boolean
Private BeenClicked As Boolean
Private CurrentState As Integer
Private RBYModeOK As Boolean
Private GameMode As Byte
Private ListType As Byte
Private NormalGFX As Boolean

'For the nifty RuleBox ToolTips
Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type LVHITTESTINFO
   pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem As Long
End Type
Dim TT As CTooltip
Dim m_lCurItemIndex As Long

Private Sub CancelButton_Click()
    If ICalled Then
        ICalled = False
        If PressedOK Then Call MasterServer.SendData("PCAN:" & ChallengeNumber)
        ChallengeNumber = 0
        ChallengePending = False
        Unload Me
    Else
        Call MasterServer.SendData("PREF:" & ChallengeNumber)
        ChallengeNumber = 0
        ChallengePending = False
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    If TT.Style = TTBalloon Then TT.Style = TTStandard Else TT.Style = TTBalloon
End Sub

Private Sub Closer_Timer()
    Unload Me
End Sub

Private Sub FlashTimer_Timer()
    Dim DummyData As Long
    
    If ICalled Or BeenClicked Then
        DummyData = FlashWindow(ChallengeWindow.hWnd, 0)
        FlashTimer.Enabled = False
        Exit Sub
    End If
    
    If CurrentState = 1 Then
        DummyData = FlashWindow(ChallengeWindow.hWnd, 0)
        CurrentState = 0
    Else
        DummyData = FlashWindow(ChallengeWindow.hWnd, 1)
        CurrentState = 1
    End If
End Sub

Private Sub Form_GotFocus()
    BeenClicked = True
End Sub

'Private Sub Form_Initialize()
'    'InitCommonControls
'End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim TempC() As Boolean
    Dim BattleTotal As Long
    Dim BattleModeItem(3) As Boolean
    Dim Ubers As Byte
    Dim C(1, 3) As Boolean
    Dim G1 As BattleModes
    Dim G2 As BattleModes
    '>>> Call WriteDebugLog("Challenge Window loaded")
    If ChallengeNumber = 0 Then Closer.Enabled = True: Exit Sub
    If Player(ChallengeNumber).Picture = 0 Then Unload Me: Exit Sub
    If Player(ChallengeNumber).ShowTeam Or ChallengeNumber = YourNumber Then
        TeamFrame.Visible = True
        TeamFrame.Top = 360
        PTInfo.Top = 2280
        RecFrame.Top = 2280
        RuleFrame.Top = 3840
        Me.Height = 6975
    Else
        TeamFrame.Visible = False
        PTInfo.Top = 360
        RecFrame.Top = 360
        RuleFrame.Top = 1900
        Me.Height = 5160
    End If
    NormalGFX = True
    ChallengeWindow.Top = (Screen.Height - ChallengeWindow.Height) \ 2
    
'    On Error Resume Next
'    CCompat(0).Picture = LoadResPicture("RBYT", vbResIcon)
'    CCompat(1).Picture = LoadResPicture("GSCT", vbResIcon)
'    CCompat(2).Picture = LoadResPicture("ADV", vbResIcon)
'    CCompat(3).Picture = LoadResPicture("ADV+", vbResIcon)
'    CCompat(4).Picture = LoadResPicture("MOD", vbResIcon)
'    CCompat(5).Picture = LoadResPicture("RBY", vbResIcon)
'    CCompat(6).Picture = LoadResPicture("GSC", vbResIcon)
'    CCompat(7).Picture = MainContainer.Conditions.ListImages(8).Picture
'    CCompat(8).Picture = MainContainer.Conditions.ListImages(8).Picture
    MasterServer.ChallengeLoaded = True
    MasterServer.mnuTeam.Enabled = False
    Call DoImages
    Select Case Player(ChallengeNumber).GameVersion
    Case 0: imgBMode.Picture = LoadResPicture("RBYT", vbResIcon)
    Case 1: imgBMode.Picture = LoadResPicture("GSCT", vbResIcon)
    Case 2: imgBMode.Picture = LoadResPicture("ADV", vbResIcon)
    Case 3: imgBMode.Picture = LoadResPicture("ADV+", vbResIcon)
    Case 4: imgBMode.Picture = LoadResPicture("MOD", vbResIcon)
    Case 5: imgBMode.Picture = LoadResPicture("RBY", vbResIcon)
    Case 6: imgBMode.Picture = LoadResPicture("GSC", vbResIcon)
    End Select
    
    'Figure out which modes are supported & fill in the list
    ListType = 0
    G1 = CompatVersion(Player(YourNumber).GameVersion)
    G2 = CompatVersion(Player(ChallengeNumber).GameVersion)
    If G1 = nbRBYBattle And G2 = nbRBYBattle Then
        BattleModeItem(1) = True
        BattleModeItem(2) = True
        ListType = 1
    ElseIf G1 <> nbAdvBattle And G2 <> nbAdvBattle Then
        BattleModeItem(2) = True
        ListType = 2
        If Player(YourNumber).Compatibility(0) And Player(ChallengeNumber).Compatibility(0) Then
            BattleModeItem(1) = True
            ListType = 1
        End If
    ElseIf G1 = nbAdvBattle And G2 = nbAdvBattle Then
        BattleModeItem(3) = True
        ListType = 3
    End If
    
    ModeSelector.Clear
    If BattleModeItem(1) Then ModeSelector.AddItem "RBY Mode": GameMode = 1
    If BattleModeItem(2) Then ModeSelector.AddItem "GSC Mode": GameMode = 2
    If BattleModeItem(3) Then ModeSelector.AddItem "Advance Mode": ModeSelector.AddItem "Advance Double Battle Mode": GameMode = 3
    ModeSelector.AddItem "RBY Mode (Challenge Cup)"
    ModeSelector.AddItem "GSC Mode (Challenge Cup)"
    ModeSelector.AddItem "Advance Mode (Challenge Cup)"
    ModeSelector.AddItem "Advance Double Battle Mode (Challenge Cup)"
    ModeSelector.ListIndex = 0
    If ListType = 0 Then GameMode = 5
    
'    'Disable terrain selection if not in Advance
'    Select Case GameMode
'        Case 3, 4, 7, 8
            TerrainSelector.Enabled = True
'        Case Else
'            TerrainSelector.Enabled = False
'    End Select

    'Disable Present rule if not in GSC
    'Under an If/Then because this sometimes gets called before RuleList fills in.
    If RuleList.ListItems.count >= nbPresentRule Then
        Select Case GameMode
            Case 2, 6
                RuleList.ListItems(nbPresentRule).Ghosted = False
                RuleList.ListItems(nbPresentRule).ForeColor = vbBlack
            Case Else
                RuleList.ListItems(nbPresentRule).Ghosted = True
                RuleList.ListItems(nbPresentRule).ForeColor = vbRed
        End Select
    End If
    
    'Fill in terrain list
    TerrainSelector.AddItem "Random Terrain"
    For X = LBound(TerrainText) To UBound(TerrainText)
        If TerrainText(X) <> "" Then TerrainSelector.AddItem TerrainText(X)
    Next
    TerrainSelector.ListIndex = 0
    
    If ICalled = False Then
        Label2.Caption = "You have been challenged!"
        Label2.Visible = True
        Me.Caption = Player(ChallengeNumber).Name
        OKButton.Caption = "&Accept"
        CancelButton.Caption = "&Refuse"
        ModeSelector.Visible = False
        TerrainSelector.Visible = False
    Else
        Label2.Caption = Player(ChallengeNumber).Name '"Challenge Setup"
        Label2.Visible = True 'False
        Me.Caption = "Challenge Setup"
        OKButton.Caption = "&Challenge!"
        CancelButton.Caption = "C&ancel"
    End If

    '>>> Call WriteDebugLog("Showing Window...")
    For X = 1 To 6
        If BasePKMN(Player(ChallengeNumber).PKMN(X)).Uber Then Ubers = Ubers + 1
    Next
    If Not Player(ChallengeNumber).ShowTeam Then
        Select Case Ubers
            Case 1
                Me.Caption = Me.Caption & " - 1 Uber"
            Case Is > 1
                Me.Caption = Me.Caption & " - " & Ubers & " Ubers"
        End Select
    End If
    Me.Show
    DoEvents
    
    '>>> Call WriteDebugLog("Loading ToolTip class")
    Set TT = New CTooltip
    TT.Style = TTStandard
    TT.Icon = TTIconInfo
    TT.DelayTime = 400
    TT.VisibleTime = 32767
    
    'Don't ask me why this has to happen, but it screws up without it.
    '>>> Call WriteDebugLog("Adding dummy listitem")
    RuleList.ListItems.Add 1, , "              ."
    RuleList.ListItems.Clear
    '>>> Call WriteDebugLog("Loading List checkboxes")
    RuleList.Checkboxes = ICalled
    If Not ICalled Then
        TempVar = 0
        GameMode = ChallengeMode
        ChallengePending = True
        For X = 1 To UBound(RuleSelected)
            If RuleSelected(X) Then
                TempVar = TempVar + 1
                RuleList.ListItems.Add TempVar, , RuleText(X)
                RuleList.ListItems(TempVar).ToolTipText = RuleToolTip(X)
            End If
        Next
        Select Case GameMode
            Case 1
                DisplayMode.Caption = "RBY Mode"
            Case 2
                DisplayMode.Caption = "GSC Mode"
            Case 3
                DisplayMode.Caption = "Advance Mode"
            Case 4
                DisplayMode.Caption = "Advance Double Battle Mode"
            Case 5
                DisplayMode.Caption = "RBY Mode (Challenge Cup)"
            Case 6
                DisplayMode.Caption = "GSC Mode (Challenge Cup)"
            Case 7
                DisplayMode.Caption = "Advance Mode (Challenge Cup)"
            Case 8
                DisplayMode.Caption = "Advance Double Battle Mode (Challenge Cup)"
        End Select
'        Select Case GameMode
'            Case 3, 4, 7, 8
        If ChallTerrain = 0 Then
            DisplayMode.Caption = DisplayMode.Caption & " - Random Terrain"
        Else
            DisplayMode.Caption = DisplayMode.Caption & " - " & TerrainText(ChallTerrain - 1)
        End If
'        End Select
        Call DoImages
    Else
        '>>> Call WriteDebugLog("Adding rules")
        For X = 1 To UBound(RuleText)
            If RuleText(X) <> "" Then
                RuleList.ListItems.Add X, , RuleText(X)
                'RuleList.ListItems(X).ToolTipText = RuleToolTip(X)
            End If
        Next
        Select Case GameMode
            Case 3, 4, 7, 8
                'TerrainSelector.Enabled = True
                RuleList.ListItems(nbPresentRule).Ghosted = False
                RuleList.ListItems(nbPresentRule).ForeColor = vbBlack
            Case Else
                'TerrainSelector.Enabled = False
                RuleList.ListItems(nbPresentRule).Ghosted = True
                RuleList.ListItems(nbPresentRule).ForeColor = vbRed
        End Select
        '>>> Call WriteDebugLog("Checking RBY status")
'        RBYModeOK = True
'        If Not (Compatibility(0) Or Compatibility(5)) Then RBYModeOK = False
'        If Not (Player(ChallengeNumber).Compatibility(0) Or Player(ChallengeNumber).Compatibility(5)) Then RBYModeOK = False
'        If Not RBYModeOK Then RuleList.ListItems(nbRBYMode).ForeColor = vbRed: RuleList.ListItems(6).Ghosted = True
'        If CompatVersion(Player(ChallengeNumber).GameVersion) <> nbadvBattle Then RuleList.ListItems(nbDoubleBattle).ForeColor = vbRed: RuleList.ListItems(nbDoubleBattle).Ghosted = True
        
        '>>> Call WriteDebugLog("Checking Stadium")
        If Not (Player(ChallengeNumber).StadiumOK And Player(YourNumber).StadiumOK) Then RuleList.ListItems(5).ForeColor = vbRed: RuleList.ListItems(5).Ghosted = True
        ChallengePending = False
        Call ReadRules
        '>>> Call WriteDebugLog("Sending GETS: packet")
        If GetSpeed Then Call MasterServer.SendData("GETS:" & Chr$(ChallengeNumber))
    End If
    PressedOK = False
    ChallengeWindow.Icon = MainContainer.MiniTrainers.ListImages(Player(ChallengeNumber).Picture).Picture
    Label1.Caption = Player(ChallengeNumber).Extra
    BattleTotal = Player(ChallengeNumber).Wins + Player(ChallengeNumber).Losses + Player(ChallengeNumber).Ties + Player(ChallengeNumber).Disconnect
    If BattleTotal > 0 Then
        BRecord(0).Caption = Player(ChallengeNumber).Wins & vbCrLf & Int((Player(ChallengeNumber).Wins * 100) / BattleTotal) & "%"
        BRecord(1).Caption = Player(ChallengeNumber).Losses & vbCrLf & Int((Player(ChallengeNumber).Losses * 100) / BattleTotal) & "%"
        BRecord(2).Caption = Player(ChallengeNumber).Ties & vbCrLf & Int((Player(ChallengeNumber).Ties * 100) / BattleTotal) & "%"
        BRecord(3).Caption = Player(ChallengeNumber).Disconnect & vbCrLf & Int((Player(ChallengeNumber).Disconnect * 100) / BattleTotal) & "%"
    Else
        BRecord(0).Caption = Player(ChallengeNumber).Wins & vbCrLf & "---%"
        BRecord(1).Caption = Player(ChallengeNumber).Losses & vbCrLf & "---%"
        BRecord(2).Caption = Player(ChallengeNumber).Ties & vbCrLf & "---%"
        BRecord(3).Caption = Player(ChallengeNumber).Disconnect & vbCrLf & "---%"
    End If
    BattleTotalDisp.Caption = BattleTotal & " rated battles"
    PowerBar.Caption = Player(ChallengeNumber).Rank
    PowerBar.Value = Val(Player(ChallengeNumber).Rank)
    If Val(Player(ChallengeNumber).Speed) > 0 Then
        SpeedBar.Caption = Player(ChallengeNumber).Speed
        'SpeedBar.Value = Val(Player(ChallengeNumber).Speed)
    Else
        SpeedBar.Caption = "?"
        'SpeedBar.Value = 0
    End If
    Call ColorBars
    If ChallengeNumber = YourNumber Then
        OKButton.Enabled = False
        OKButton.ToolTipText = "Can't challenge self!"
    Else
        OKButton.Enabled = True
        OKButton.ToolTipText = ""
    End If
    If Player(ChallengeNumber).Version <> You.ProgVersion Then
        OKButton.Enabled = False
        OKButton.ToolTipText = "Version Conflict!"
        ChallengeWindow.Caption = ChallengeWindow.Caption & " - Version Conflict!"
    End If
    If MasterServer.ServerVersion <> You.ProgVersion Then
        OKButton.Enabled = False
        OKButton.ToolTipText = "Server Version Conflict!"
        ChallengeWindow.Caption = ChallengeWindow.Caption & " - Server Version Conflict!"
    End If
    'If CompatVersion(Player(ChallengeNumber).GameVersion) <> CompatVersion(Player(YourNumber).GameVersion) Then
    '    OKButton.Enabled = False
    '    OKButton.ToolTipText = "Team Conflict!"
    '    ChallengeWindow.Caption = ChallengeWindow.Caption & " - Team Conflict!"
    'End If
    If Not ICalled And MusicOption = 1 Then Call PlayMusic(9, True)
    Call ModeSelector_Change
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    TT.Destroy
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Call CancelButton_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MasterServer.ChallengeLoaded = False
    MasterServer.mnuTeam.Enabled = True
    If ICalled = False And MusicOption = 1 Then Call StopMusic
    If Not ChallengePending Then ChallengeNumber = 0
End Sub

Private Sub Label12_Click()

End Sub

Private Sub imgTBMode_Click()

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Image1_Click(Index As Integer)
    NormalGFX = Not NormalGFX
    Call DoImages
End Sub

Private Sub ModeSelector_Change()
    If Not ICalled Then Exit Sub
    Select Case ListType
        Case 0
            GameMode = ModeSelector.ListIndex + 5
        Case 1
            If ModeSelector.ListIndex <= 1 Then
                GameMode = ModeSelector.ListIndex + 1
            Else
                GameMode = ModeSelector.ListIndex + 3
            End If
        Case 2
            If ModeSelector.ListIndex = 0 Then
                GameMode = 2
            Else
                GameMode = ModeSelector.ListIndex + 4
            End If
        Case 3
            GameMode = ModeSelector.ListIndex + 3
    End Select
    'Disable terrain selection if not in Advance
'    Select Case GameMode
'        Case 3, 4, 7, 8
'            TerrainSelector.Enabled = True
'        Case Else
'            TerrainSelector.Enabled = False
'    End Select
    'Disable Present rule if not in GSC
    'Under an If/Then because this sometimes gets called before RuleList fills in.
    If RuleList.ListItems.count >= nbPresentRule Then
        Select Case GameMode
            Case 2, 6
                RuleList.ListItems(nbPresentRule).Ghosted = False
                RuleList.ListItems(nbPresentRule).ForeColor = vbBlack
            Case Else
                RuleList.ListItems(nbPresentRule).Ghosted = True
                RuleList.ListItems(nbPresentRule).ForeColor = vbRed
        End Select
    End If
    Call DoImages
End Sub

Private Sub ModeSelector_Click()
    Call ModeSelector_Change
End Sub

Private Sub ModeSelector_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ModeSelector_Change
End Sub

Private Sub ModeSelector_LostFocus()
    Call ModeSelector_Change
End Sub

Private Sub OKButton_Click()
    Dim X As Integer
    
    If ICalled Then
        For X = 1 To RuleList.ListItems.count
            RuleSelected(X) = RuleList.ListItems(X).Checked
        Next
        ChallTerrain = TerrainSelector.ListIndex
        Call MasterServer.SendData("CHLN:" & FixedHex(ChallengeNumber, 3) & FixedHex(GameMode, 1) & FixedHex(ChallTerrain, 1) & FixedHex(MakeBinArray(RuleSelected), 8))
        PressedOK = True
        ChallengePending = True
        OKButton.Enabled = False
        Call SaveRules
    Else
        Player(YourNumber).BattlingWith = ChallengeNumber
        IsServer = False
        RelayServer = True
        Call MasterServer.SendData("PACC:" & FixedHex(ChallengeNumber, 3) & FixedHex(GameMode, 1) & FixedHex(ChallTerrain, 1) & FixedHex(MakeBinArray(RuleSelected), 8))
        On Error Resume Next
        Unload Me
    End If
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub RuleList_Click()
    On Error Resume Next
    If Not ICalled Then RuleList.ListItems(RuleList.SelectedItem.Index).Selected = False
End Sub

Private Sub RuleList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Ghosted Then Item.Checked = False
'    If Item.Index = nbChallengeCup Then
'        Call DoImages
'        If RBYModeOK Or Item.Checked = True Then
'            RuleList.ListItems(nbRBYMode).ForeColor = vbBlack
'            RuleList.ListItems(nbRBYMode).Ghosted = False
'        Else
'            RuleList.ListItems(nbRBYMode).ForeColor = vbRed
'            RuleList.ListItems(nbRBYMode).Ghosted = True
'            RuleList.ListItems(nbRBYMode).Checked = False
'        End If
'    End If
End Sub

Private Sub RuleList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not ICalled Then RuleList.ListItems(Item.Index).Selected = False
End Sub

Private Sub RuleList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim Z As Integer
    Dim Temp As String
    On Error Resume Next
    lvhti.pt.X = X / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(RuleList.hWnd, LVM_HITTEST, 0, lvhti) + 1
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            TT.Destroy
        Else
            Temp = RuleList.ListItems(m_lCurItemIndex).Text
            TT.Title = Temp
            For Z = 1 To UBound(RuleText)
                If RuleText(Z) = Temp Then Exit For
            Next Z
            TT.TipText = Replace(RuleToolTip(Z), "%n", vbNewLine)
            'TT.MaxLen = 300
            TT.Create RuleList.hWnd
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If Val(Player(ChallengeNumber).Speed) > 0 Or Player(ChallengeNumber).Speed = "0000" Then
        If Val(Player(ChallengeNumber).Speed) > 9999 Then Player(ChallengeNumber).Speed = "9999"
        SpeedBar.Caption = Player(ChallengeNumber).Speed
        'SpeedBar.Value = Val(Player(ChallengeNumber).Speed)
        Timer1.Enabled = False
        Call ColorBars
    End If
End Sub

Private Sub ColorBars()
    Select Case PowerBar.Value
        Case 0 To 25
            PowerBar.FillColor = vbRed
        Case 26 To 50
            PowerBar.FillColor = vbYellow
        Case Else
            PowerBar.FillColor = vbGreen
    End Select

'    Select Case SpeedBar.Value
'        Case 0 To 500
'            SpeedBar.FillColor = vbGreen
'        Case 501 To 2000
'            SpeedBar.FillColor = vbYellow
'        Case Else
'            SpeedBar.FillColor = vbRed
'    End Select
End Sub

Private Sub SaveRules()
    Dim X As Integer
    For X = 1 To RuleList.ListItems.count
        SaveSetting "NetBattle", "Battle Rules", "Rule" & Format(X, "00"), RuleList.ListItems(X).Checked
    Next X
    Select Case Player(YourNumber).GameVersion
        Case 0, 5
            SaveSetting "NetBattle", "Battle Rules", "RBYMode", GameMode
        Case 1, 6
            SaveSetting "NetBattle", "Battle Rules", "GSCMode", GameMode
        Case 2, 3
            SaveSetting "NetBattle", "Battle Rules", "Mode", GameMode
            SaveSetting "NetBattle", "Battle Rules", "Terrain", TerrainSelector.ListIndex
    End Select
End Sub

Private Sub ReadRules()
    Dim X As Integer
    Dim GMTemp As Byte
'    If GetSetting("NetBattle", "Battle Rules", "Rule" & Format(nbChallengeCup, "00", False), False) = True Then
'        RuleList.ListItems(nbRBYMode).ForeColor = vbBlack
'        RuleList.ListItems(nbRBYMode).Ghosted = False
'    End If
    For X = 1 To RuleList.ListItems.count
        If Not RuleList.ListItems(X).Ghosted Then
            RuleList.ListItems(X).Checked = GetSetting("NetBattle", "Battle Rules", "Rule" & Format(X, "00"), False)
        End If
    Next X
    Select Case Player(YourNumber).GameVersion
        Case 0, 5
            GMTemp = GetSetting("NetBattle", "Battle Rules", "RBYMode", 0)
        Case 1, 6
            GMTemp = GetSetting("NetBattle", "Battle Rules", "GSCMode", 0)
        Case 2, 3
            GMTemp = GetSetting("NetBattle", "Battle Rules", "Mode", 0)
    End Select
    TerrainSelector.ListIndex = GetSetting("NetBattle", "Battle Rules", "Terrain", 0)
    If GMTemp = 0 Then Exit Sub
    Select Case ListType
        'Only Challenge Cups
        Case 0
            Select Case GMTemp
                Case Is >= 5
                    ModeSelector.ListIndex = GMTemp - 5
                Case Else
                    'Not on this list
            End Select
        'RBY, GSC, CCups
        Case 1
            Select Case GMTemp
                Case Is <= 2
                    ModeSelector.ListIndex = GMTemp - 1
                Case Is >= 5
                    ModeSelector.ListIndex = GMTemp - 3
                Case Else
                    'Not on this list
            End Select
        'GSC,CCups
        Case 2
            Select Case GMTemp
                Case 2
                    ModeSelector.ListIndex = 0
                Case Is >= 5
                    ModeSelector.ListIndex = GMTemp - 4
                Case Else
                    'Not on this list
            End Select
        'Adv, CCups
        Case 3
            Select Case GMTemp
                Case Is >= 3
                    ModeSelector.ListIndex = GMTemp - 3
                Case Else
                    'Not on this list
            End Select
    End Select
End Sub

Private Sub DoImages()
    Dim X As Integer
    Dim Y As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim RandBat As Boolean
    If PTInfo.Top = 360 Then Exit Sub
    On Error GoTo FirstLoad
    'If ICalled Then
    '    RandBat = RuleList.ListItems(nbChallengeCup).Checked
    'Else
    '    RandBat = RuleSelected(nbChallengeCup)
    'End If
    If GameMode > 4 Then RandBat = True Else RandBat = False
    LockWindowUpdate Me.hWnd
    For X = 1 To 6
        If RandBat And NormalGFX Then
            Call MainContainer.DoPicture("0rs.gif")
            PokeName(X - 1).Caption = "Random"
            Image1(X - 1).ToolTipText = "Random"
        ElseIf Player(ChallengeNumber).ShowTeam Or ChallengeNumber = YourNumber Then
            Call MainContainer.DoPicture(Player(ChallengeNumber).PKMNImage(X))
            PokeName(X - 1).Caption = BasePKMN(Player(ChallengeNumber).PKMN(X)).Name
            Image1(X - 1).ToolTipText = BasePKMN(Player(ChallengeNumber).PKMN(X)).Name
        Else
            Call MainContainer.DoPicture("0rs.gif")
            If BasePKMN(Player(ChallengeNumber).PKMN(X)).Uber Then
                PokeName(X - 1).Caption = "Hidden(Über)"
                Image1(X - 1).ToolTipText = "Hidden(Über)"
'            ElseIf BasePKMN(Player(ChallengeNumber).PKMN(X)).Legendary Then
'                PokeName(X - 1).Caption = "Hidden(Legendary)"
'                Image1(X - 1).ToolTipText = "Hidden(Legendary)"
            Else
                PokeName(X - 1).Caption = "Hidden"
                Image1(X - 1).ToolTipText = "Hidden"
            End If
        End If
        picSwap.Picture = MainContainer.SwapSpace.Picture
        Y = GetYOffset(picSwap) * Screen.TwipsPerPixelY
        Image1(X - 1).Picture = MainContainer.SwapSpace.Picture
        TempVar = (Picture1(X - 1).Width - Image1(X - 1).Width) / 2
        TempVar2 = (Picture1(X - 1).Height - Image1(X - 1).Height - Y) / 2
        Image1(X - 1).Left = TempVar
        Image1(X - 1).Top = TempVar2
    Next X
    LockWindowUpdate 0
    Exit Sub
FirstLoad:
    'RandBat = (GetSetting("NetBattle", "Battle Rules", "Rule" & Format(nbChallengeCup, "00"), False) = True)
    'EDIT: Maybe change this later.
    RandBat = False
    Resume Next
End Sub

