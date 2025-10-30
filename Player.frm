VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form ChallengeWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Player Details"
   ClientHeight    =   5970
   ClientLeft      =   5535
   ClientTop       =   2265
   ClientWidth     =   4695
   Icon            =   "Player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRandbat 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   4140
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   14
      Top             =   5460
      Width           =   4515
      Begin VB.CommandButton CancelButton 
         Cancel          =   -1  'True
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   60
         Width           =   1995
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "&Challenge"
         Height          =   375
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   1995
      End
   End
   Begin VB.Frame PTInfo 
      Caption         =   "Player  && Team Info"
      Height          =   2355
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   4455
      Begin CCRProgressBar6.ccrpProgressBar PowerBar 
         Height          =   255
         Left            =   900
         Top             =   1740
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         AutoCaption     =   1
         BackColor       =   0
         Caption         =   "0%"
         FillColor       =   65280
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
         Caption         =   "Ties"
         Height          =   255
         Left            =   3660
         TabIndex        =   20
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Losses"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wins"
         Height          =   255
         Left            =   2340
         TabIndex        =   18
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label lblUbers 
         Alignment       =   2  'Center
         Caption         =   "6 Ubers | Spec."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   4320
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   3780
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   3060
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2340
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   1620
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   900
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   180
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgBMode 
         Height          =   480
         Left            =   120
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Power"
         Height          =   255
         Left            =   900
         TabIndex        =   13
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label1 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   4215
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   3660
         TabIndex        =   11
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label BRecord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   2340
         TabIndex        =   9
         Top             =   1740
         Width           =   675
      End
   End
   Begin VB.Timer FlashTimer 
      Interval        =   1000
      Left            =   9720
      Top             =   6720
   End
   Begin VB.Timer Closer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8760
      Top             =   6720
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   9240
      Top             =   6720
   End
   Begin VB.Frame RuleFrame 
      Caption         =   "Rules && Modes"
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   2820
      Width           =   4455
      Begin MSComctlLib.ListView RuleList 
         Height          =   1875
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3307
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
         ItemData        =   "Player.frx":1272
         Left            =   2820
         List            =   "Player.frx":1274
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1515
      End
      Begin VB.ComboBox ModeSelector 
         Height          =   315
         ItemData        =   "Player.frx":1276
         Left            =   120
         List            =   "Player.frx":1278
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label DisplayMode 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7935
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
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "ChallengeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PressedOK As Boolean
Private BeenClicked As Boolean
Private CurrentState As Integer
Private NormalGFX As Boolean
Private CanBattle(1 To 4) As Boolean
Private RandPoke(1 To 6) As Long
Private GameMode As Long

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
    Dim X As Long
    Dim Y As Long
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim TempC() As Boolean
    Dim BattleTotal As Long
    Dim BattleModeItem(3) As Boolean
    Dim Ubers As Byte
    Dim Spec As Boolean
    Dim ListType As Long
    Dim C(1, 3) As Boolean
    Dim G1 As BattleModes
    Dim G2 As BattleModes
    Dim Temp As String
    '>>> Call WriteDebugLog("Challenge Window loaded")
    With Player(ChallengeNumber)
    'chaos my special version lets me view all teams. :)
    ' be sure to take this out for masamune..
    .ShowTeam = True
        If ChallengeNumber = 0 Then Closer.Enabled = True: Exit Sub
        If .Picture = 0 Then Unload Me: Exit Sub
        NormalGFX = True
        ChallengeWindow.Top = (Screen.Height - ChallengeWindow.Height) \ 2
        
    '    On Error Resume Next
        MasterServer.ChallengeLoaded = True
        MasterServer.mnuTeam.Enabled = False
        Call DoImages
        Select Case .GameVersion
        Case 0: imgBMode.Picture = LoadResPicture("RBYT", vbResIcon)
        Case 1: imgBMode.Picture = LoadResPicture("GSCT", vbResIcon)
        Case 2: imgBMode.Picture = LoadResPicture("ADV", vbResIcon)
        Case 3: imgBMode.Picture = LoadResPicture("ADV+", vbResIcon)
        Case 4: imgBMode.Picture = LoadResPicture("MOD", vbResIcon)
        Case 5: imgBMode.Picture = LoadResPicture("RBY", vbResIcon)
        Case 6: imgBMode.Picture = LoadResPicture("GSC", vbResIcon)
        End Select
        
        'Figure out which modes are supported & fill in the list
        ListType = 3
        G1 = CompatVersion(Player(YourNumber).GameVersion)
        G2 = CompatVersion(.GameVersion)
        ZeroMemory CanBattle(1), 8
        If G1 = nbRBYBattle And G2 = nbRBYBattle Then
            CanBattle(1) = True
            CanBattle(2) = True
            ListType = 0
        ElseIf G1 <> nbAdvBattle And G2 <> nbAdvBattle Then
            CanBattle(2) = True
            ListType = 1
            If Player(YourNumber).Compatibility(0) And .Compatibility(0) Then
                CanBattle(1) = True
                ListType = 0
            End If
        ElseIf G1 = nbAdvBattle And G2 = nbAdvBattle Then
            CanBattle(3) = True
            CanBattle(4) = True
            ListType = 2
        End If
        
        ModeSelector.Clear
        ModeSelector.AddItem "RBY Mode"
        ModeSelector.AddItem "GSC Mode"
        ModeSelector.AddItem "Advance Mode"
        ModeSelector.AddItem "Advance Double Battle Mode"
        ModeSelector.ListIndex = ListType
        
    
    '    'Disable Present rule if not in GSC
    '    'Under an If/Then because this sometimes gets called before RuleList fills in.
    '    If RuleList.ListItems.count >= nbPresentRule Then
    '        Select Case GameMode
    '            Case 2, 6
    '                RuleList.ListItems(nbPresentRule).Ghosted = False
    '                RuleList.ListItems(nbPresentRule).ForeColor = vbBlack
    '            Case Else
    '                RuleList.ListItems(nbPresentRule).Ghosted = True
    '                RuleList.ListItems(nbPresentRule).ForeColor = vbRed
    '        End Select
    '    End If
        
        'Fill in terrain list
        TerrainSelector.AddItem "Random Terrain"
        For X = LBound(TerrainText) To UBound(TerrainText)
            If TerrainText(X) <> "" Then TerrainSelector.AddItem TerrainText(X)
        Next
        TerrainSelector.ListIndex = 0
        
        If ICalled Then
            Label2.Caption = .Name '"Challenge Setup"
            Label2.Visible = True 'False
            Me.Caption = "Challenge Setup"
            OKButton.Caption = "&Challenge!"
            CancelButton.Caption = "C&ancel"
        Else
            Label2.Caption = "You have been challenged!"
            Label2.Visible = True
            Me.Caption = .Name & " - Ping Speed: " & .Speed
            OKButton.Caption = "&Accept"
            CancelButton.Caption = "&Refuse"
            ModeSelector.Visible = False
            TerrainSelector.Visible = False
        End If
    
        '>>> Call WriteDebugLog("Showing Window...")
        Spec = False
        For X = 1 To 6
            If BasePKMN(.PKMN(X)).Uber Then Ubers = Ubers + 1
            If CompatVersion(Player(ChallengeNumber).GameVersion) = nbAdvBattle Then
                If .PKMN(X) = 202 Then Ubers = Ubers + 1 'Ugh wobb.
            End If
            If Not Spec Then
                For Y = X + 1 To 6
                    If .PKMN(X) = .PKMN(Y) Then Spec = True
                Next Y
            End If
        Next
        
        lblUbers.Caption = vbNullString
        If Not .ShowTeam Then
            Select Case Ubers
                Case 1
                    Temp = "1 Uber"
                    'Me.Caption = Me.Caption & " - 1 Uber"
                Case Is > 1
                    Temp = Ubers & " Ubers"
                    'Me.Caption = Me.Caption & " - " & Ubers & " Ubers"
            End Select
            If Spec Then
                If Len(Temp) = 0 Then
                    Temp = "Species Clause"
                Else
                    Temp = Temp & " - Spec."
                End If
            End If
            lblUbers.Caption = Temp
        End If
        
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
            End Select
            
            If ChallTerrain = 0 Then
                DisplayMode.Caption = DisplayMode.Caption & " - Random Terrain"
            Else
                DisplayMode.Caption = DisplayMode.Caption & " - " & TerrainText(ChallTerrain - 1)
            End If
    '        End Select
            Call DoImages
        Else
            If GetSpeed Then Call MasterServer.SendData("GETS:" & Chr$(ChallengeNumber))
        End If
        PressedOK = False
        ChallengeWindow.Icon = MainContainer.MiniTrainers.ListImages(.Picture).Picture
        Label1.Caption = .Extra
        X = .Wins + .Losses + .Ties
        If X > 0 Then
            BRecord(0).Caption = .Wins & vbCrLf & Int((.Wins * 100) / X) & "%"
            BRecord(1).Caption = .Losses & vbCrLf & Int((.Losses * 100) / X) & "%"
            BRecord(2).Caption = .Ties & vbCrLf & Int((.Ties * 100) / X) & "%"
        Else
            BRecord(0).Caption = .Wins & vbCrLf & "---%"
            BRecord(1).Caption = .Losses & vbCrLf & "---%"
            BRecord(2).Caption = .Ties & vbCrLf & "---%"
        End If
        PowerBar.Caption = .Rank
        PowerBar.Value = Val(.Rank)
        If Val(.Speed) > 0 Then
            'SpeedBar.Caption = .Speed
            'SpeedBar.Value = Val(.Speed)
        Else
            'SpeedBar.Caption = "?"
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
        If .Version <> You.ProgVersion Then
            OKButton.Enabled = False
            OKButton.ToolTipText = "Version Conflict!"
            ChallengeWindow.Caption = ChallengeWindow.Caption & " - Version Conflict!"
        End If
        If MasterServer.ServerVersion <> You.ProgVersion Then
            OKButton.Enabled = False
            OKButton.ToolTipText = "Server Version Conflict!"
            ChallengeWindow.Caption = ChallengeWindow.Caption & " - Server Version Conflict!"
        End If
        
        If Not ICalled And MusicOption = 1 Then Call PlayMusic(9, True)
        Call ModeSelector_Change
        Me.Show
        DoEvents
        If ICalled Then
                '>>> Call WriteDebugLog("Adding rules")
            For X = 1 To UBound(RuleText)
                If RuleText(X) <> "" Then
                    RuleList.ListItems.Add X, , RuleText(X)
                    'RuleList.ListItems(X).ToolTipText = RuleToolTip(X)
                End If
            Next
    
            '>>> Call WriteDebugLog("Checking RBY status")
    '        RBYModeOK = True
    '        If Not (Compatibility(0) Or Compatibility(5)) Then RBYModeOK = False
    '        If Not (.Compatibility(0) Or .Compatibility(5)) Then RBYModeOK = False
    '        If Not RBYModeOK Then RuleList.ListItems(nbRBYMode).ForeColor = vbRed: RuleList.ListItems(6).Ghosted = True
    '        If CompatVersion(.GameVersion) <> nbadvBattle Then RuleList.ListItems(nbDoubleBattle).ForeColor = vbRed: RuleList.ListItems(nbDoubleBattle).Ghosted = True
            
            '>>> Call WriteDebugLog("Checking Stadium")
            If Not (.StadiumOK And Player(YourNumber).StadiumOK) Then RuleList.ListItems(5).ForeColor = vbRed: RuleList.ListItems(5).Ghosted = True
            ChallengePending = False
            Call ReadRules
            Call ModeSelector_Change
            '>>> Call WriteDebugLog("Sending GETS: packet")
        End If
    End With
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
    Dim X As Long
    On Error GoTo ETrap
    X = ModeSelector.ListIndex + 1
    With RuleList.ListItems(nbRandbat)
        If CanBattle(X) Then
            .Ghosted = False
            .ForeColor = vbBlack
        Else
            .Ghosted = True
            .ForeColor = RGB(161, 161, 161)
            .Checked = False 'True
            Call RuleList_ItemCheck(RuleList.ListItems(nbRandbat))
        End If
    End With
    
    With RuleList.ListItems(nbPresentRule)
        If X < 3 Then
            .Ghosted = False
            .ForeColor = vbBlack
        Else
            .Ghosted = True
            .ForeColor = RGB(161, 161, 161)
            .Checked = False
        End If
    End With
    
    Call DoImages
ETrap:
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
        Call MasterServer.SendData("CHLN:" & FixedHex(ChallengeNumber, 3) & Player(ChallengeNumber).GameVersion & FixedHex(ModeSelector.ListIndex + 1, 1) & FixedHex(ChallTerrain, 1) & FixedHex(MakeBinArray(RuleSelected), 8))
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
    If Item.Ghosted Then Item.Checked = Not Item.Checked
    If Item.Index = nbRandbat Then
        tmrRandbat.Enabled = Item.Checked
        Call DoImages
    End If
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
        'SpeedBar.Caption = Player(ChallengeNumber).Speed
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
    SaveSetting "NetBattle", "Battle Rules", "Mode", ModeSelector.ListIndex
End Sub

Private Sub ReadRules()
    Dim X As Integer
    Dim GMTemp As Byte
    For X = 1 To RuleList.ListItems.count
        If Not RuleList.ListItems(X).Ghosted Then
            RuleList.ListItems(X).Checked = GetSetting("NetBattle", "Battle Rules", "Rule" & Format(X, "00"), False)
        End If
    Next X
    If ChallengeNumber = YourNumber Then
        RuleList.ListItems(nbRandbat).Checked = False
    End If
    Call RuleList_ItemCheck(RuleList.ListItems(nbRandbat))
    
    GMTemp = GetSetting("NetBattle", "Battle Rules", "Mode", 2)
    If GMTemp > 3 Then GMTemp = 2
    If RuleList.ListItems(nbRandbat).Checked Then
        ModeSelector.ListIndex = GMTemp
    ElseIf CompatVersion(Player(YourNumber).GameVersion) = nbAdvBattle And GMTemp = 3 Then
        ModeSelector.ListIndex = 3
    End If
    
End Sub

Private Sub DoImages()
    Dim X As Integer
    Dim Y As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim RandBat As Boolean
    Dim Graphic As String
    Dim Tip As String
    If PTInfo.Top = 360 Then Exit Sub
    On Error GoTo FirstLoad
    If ICalled Then
        RandBat = RuleList.ListItems(nbRandbat).Checked
    Else
        RandBat = RuleSelected(nbRandbat)
    End If
    If RandBat And Not tmrRandbat.Enabled Then tmrRandbat.Enabled = True
    
    For X = 1 To 6
    'chaos
    Player(ChallengeNumber).ShowTeam = True
    
    
        
        If RandBat Then
            Graphic = Format(RandPoke(X), "000") & "_1.gif"
            Image1(X - 1).ToolTipText = "Random"
        ElseIf Player(ChallengeNumber).ShowTeam Or ChallengeNumber = YourNumber Then
            Graphic = ChooseImage(BasePKMN(Player(ChallengeNumber).PKMN(X)), nbGFXSml)
            Image1(X - 1).ToolTipText = BasePKMN(Player(ChallengeNumber).PKMN(X)).Name
        Else
            Graphic = "0rs.gif"
            If BasePKMN(Player(ChallengeNumber).PKMN(X)).Uber Then
                Image1(X - 1).ToolTipText = "Hidden (Über)"
            Else
                Image1(X - 1).ToolTipText = "Hidden"
            End If
        End If
        If Image1(X - 1).Tag <> Graphic Then
            Image1(X - 1).Tag = Graphic
            Call MainContainer.DoPicture(Graphic)
            Image1(X - 1).Picture = MainContainer.SwapSpace.Picture
        End If
    Next X
    'SetRedraw Me.hWnd, True
    Exit Sub
FirstLoad:
    'RandBat = (GetSetting("NetBattle", "Battle Rules", "Rule" & Format(nbChallengeCup, "00"), False) = True)
    'EDIT: Maybe change this later.
    RandBat = False
    Resume Next
End Sub


Private Sub tmrRandbat_Timer()
    Static X As Long
    Dim Y As Long
    Dim Z As Long
    Dim A As Long
    
    Select Case ModeSelector.ListIndex
    Case 0: Z = 151
    Case 1: Z = 251
    Case 2, 3: Z = 385
    End Select
'    X = X + 1
'    If X = 7 Then X = 1
'    Do
'        RandPoke(X) = Int(Rnd * Z) + 1
'        For Y = 1 To 6
'            If RandPoke(Y) = RandPoke(X) And X <> Y Then Exit For
'        Next Y
'    Loop Until Y = 7 And RandPoke(X) <> 201
    X = X + 1
    If X > 12 Then X = 1
    A = X
    If A > 6 Then A = 13 - X
    Do
        RandPoke(A) = Int(Rnd * Z) + 1
        For Y = 1 To 6
            If RandPoke(Y) = RandPoke(A) And A <> Y Then Exit For
        Next Y
    Loop Until Y = 7 And RandPoke(A) <> 201
DoImages
End Sub

