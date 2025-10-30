VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MasterServer 
   Caption         =   "Stadium: Connecting..."
   ClientHeight    =   4230
   ClientLeft      =   465
   ClientTop       =   1155
   ClientWidth     =   6600
   Icon            =   "MasterServer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   6600
   Begin VB.Timer IMTimer 
      Interval        =   1
      Left            =   6120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   3960
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   30000
   End
   Begin VB.Timer FloodCountTimer 
      Interval        =   1000
      Left            =   5640
      Top             =   600
   End
   Begin VB.Timer MissingDataTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6120
      Top             =   120
   End
   Begin VB.Timer Connector 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5640
      Top             =   120
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   2865
      Left            =   1980
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5054
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"MasterServer.frx":1272
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
   Begin CCRProgressBar6.ccrpProgressBar FCBar 
      Height          =   255
      Left            =   2040
      ToolTipText     =   "If your floodcount reaches 100%, you will be disconnected.  Stop chatting for a while to lower it."
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   1
      AutoCaption     =   1
      BackColor       =   0
      Caption         =   "0%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox ChatBox 
      Height          =   285
      Left            =   1980
      MaxLength       =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4575
   End
   Begin MSComctlLib.ListView UserList 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6588
      View            =   3
      Arrange         =   1
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
         Text            =   "Trainers"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6006
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picResizer 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   1800
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3735
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Floodcheck"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save Log..."
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTeam 
      Caption         =   "&Team"
      Begin VB.Menu mnuTeamItem 
         Caption         =   "&Rearrange Team..."
         Index           =   0
      End
      Begin VB.Menu mnuTeamItem 
         Caption         =   "&Access Box..."
         Index           =   1
      End
      Begin VB.Menu mnuTeamItem 
         Caption         =   "&Change Items..."
         Index           =   2
      End
      Begin VB.Menu mnuTeamItem 
         Caption         =   "&Open Team Builder..."
         Index           =   3
      End
      Begin VB.Menu mnuTeamItem 
         Caption         =   "&Load Another Team..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "&Player"
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Challenge/Info..."
         Index           =   0
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Watch Battle..."
         Index           =   1
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Private Message..."
         Index           =   2
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "C&ontrol Window"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Kick"
         Index           =   5
         Visible         =   0   'False
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
         Caption         =   "&Update Player Speeds"
         Index           =   2
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "Allow Multiline &Paste"
         Index           =   3
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "Show &Version Icons"
         Index           =   4
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Color Names in List"
         Index           =   5
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Away"
         Index           =   7
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
Attribute VB_Name = "MasterServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ChallengeLoaded As Boolean
Public TeamChanged As Boolean
Public TBUserChange As Boolean
Public WatchID As String
Public FormID As Integer
Public FreeWatch As Integer
Public UnloadingBattle As Boolean
Public ServerVersion As String
Public LookingUp As String
Public RTB As RTBClass
Private Watching() As Integer
Private ExitThis As Boolean
Private DataBuffer As String
Private FloodCheck As Integer
Private FloodEnabled As Boolean
Private ServerIssue As Integer
Private BanMessage As String
Private OldName As String
Private NowSwitching As Boolean
Private UseXOR As Boolean
Private PasteOK As Boolean
Private PasteLoop As Boolean
Private WasDisconnected As Boolean
Private ShiftState As Integer
Private ListingBuffer As String
Private ListingPackets As Long
Private ListingRcv As Long
Private ModBuffer As String
Private Sizing As Boolean
Private SizeX As Single
Private IMQueue() As String
Private Const vbDarkGreen As Long = 38400
 
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const EM_LINESCROLL = &HB6

'Resize stuff
Private Const MinWidth = 5900
Private Const MinHeight = 2800

Private Sub ChatBox_Change()
    If PasteLoop Then
        PasteLoop = False
        ChatBox.Text = ""
    End If
End Sub
Private Sub ChatBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim PasteArray() As String
    Dim Temp As String
    Dim X As Integer
    If (((Shift = 2 Or Shift = 3) And KeyCode = vbKeyV) Or (Shift = 1 And KeyCode = vbKeyInsert)) And DoMultiPaste And PasteOK And Clipboard.GetFormat(1) Then
        PasteOK = False
        Temp = Clipboard.GetText
        If Len(Temp) = 0 Then
            ReDim PasteArray(0)
        Else
            PasteArray = Split(Temp, vbNewLine)
        End If
        If UBound(PasteArray) <> 0 Then
            For X = 0 To UBound(PasteArray)
                PasteArray(X) = RTrim(PasteArray(X))
                If Len(PasteArray(X)) > 0 Then
                    Call SendData("CHAT:" & PasteArray(X))
                End If
            Next X
            PasteLoop = True
        End If
    End If
End Sub
Private Sub ChatBox_KeyUp(KeyCode As Integer, Shift As Integer)
    PasteOK = True
End Sub

Private Sub IMTimer_Timer()
    Dim X As Long
    Dim Y As Long
    Dim Data As String
    If UBound(IMQueue) = 0 Then Exit Sub
    On Error GoTo ETrap
    Data = IMQueue(1)
    Select Case ChopString(Data, 5)
    Case "IMCH:"
        X = Asc(ChopString(Data, 1))
        If IMWindowID(X) = 0 Then
            Call Code.NewIMWindow(Player(X).Name, X, Player(X).Picture)
        End If
        With IMWindowArray(IMWindowID(X))
            .LongMsgBuffer = .LongMsgBuffer & Left(Data, 200)
            If Len(Data) < 201 Then
                .LongMsgBuffer = ApplyCSFilter(.LongMsgBuffer)
                If Left$(.LongMsgBuffer, 4) = "/me " Then
                    Call .AddMessage("*** " & Player(X).Name & " " & Right$(.LongMsgBuffer, Len(.LongMsgBuffer) - 4), False, , &HC000C0)
                Else
                    Call .AddMessage(Player(X).Name & ": " & .LongMsgBuffer, False, ":", vbBlue, True, False)
                End If
                .LongMsgBuffer = ""
                If MainContainer.ActiveForm.hWnd <> .hWnd Then
                    IMWindowFlash(IMWindowID(X)) = True
                End If
            End If
        End With
    Case "KILL:"
        Call Code.KillIMWindow(Val(Data))
    Case "SHOW:"
        IMWindowArray(Val(Data)).Show
        IMWindowArray(Val(Data)).ChatBox.SetFocus
    End Select
    For X = 1 To UBound(IMQueue) - 1
        IMQueue(X) = IMQueue(X + 1)
    Next X
    ReDim Preserve IMQueue(X - 1)
ETrap:
End Sub

Private Sub UserList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Shift
    Case 1, 3, 5, 7
        ShiftState = 2
    Case 2, 6
        ShiftState = 1
    Case Else
        ShiftState = 0
    End Select
End Sub

Private Sub UserList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Temp2 As String
    Dim Temp As Integer
    
    If Button <> 2 Then Exit Sub
    On Error Resume Next
    If UserList.SelectedItem Is Nothing Then Exit Sub
    Me.PopupMenu mnuPlayer, , X, Y
End Sub


Private Sub Command1_Click()
    Dim Build As String
    Build = RTrim$(FilterIllegalChars(ChatBox.Text))
    If Len(Build) = 0 Then Exit Sub
    Call SendData("CHAT:" & Build)
    ChatBox.Text = ""
End Sub

Private Sub Command2_Click()
    If Connector.Enabled Then Unload Me: Exit Sub
    If Socket.State = sckConnected Then
        StatusBar1.Panels(1).Text = "Please wait - disconnecting..."
        Call SendData("EXIT:")
        ExitThis = True
    Else
        Unload Me
        Exit Sub
    End If
    If Right(MasterServer.Caption, 12) = "Disconnected" Then Unload Me
    
End Sub

Private Function RandomInteger(HighValue As Long, LowValue As Long) As Long
       RandomInteger = Int((HighValue - LowValue + 1) * Rnd + LowValue)
End Function

Private Sub Connector_Timer()
   ' On Error Resume Next
   ' If Socket.State = sckConnected Then
   '     Connector.Enabled = False
   '     Exit Sub
   ' End If
   ' StatusBar1.Panels(1).Text = "Trying to reconnect..."
   ' Socket.Close
   ' Socket.Connect
   StoredPKMN(1) = PKMN(RandomInteger(6, 1))
   StoredPKMN(2) = PKMN(RandomInteger(6, 1))
StoredPKMN(3) = PKMN(RandomInteger(6, 1))
StoredPKMN(4) = PKMN(RandomInteger(6, 1))
StoredPKMN(5) = PKMN(RandomInteger(6, 1))
StoredPKMN(6) = PKMN(RandomInteger(6, 1))
   Call SendData("TUPD:" & PrepareTeamInfo)
End Sub

Private Sub FloodCountTimer_Timer()
    If FloodTolerance = 0 Then Exit Sub
    If FloodCheck > 0 Then FloodCheck = FloodCheck - 1
    If FloodCheck >= FloodTolerance Then
        ServerIssue = 7
        Call SendData("EXIT:")
        ExitThis = True
    End If
    Call ChangeFCBar(FloodCheck)
End Sub

Private Sub Form_Load()
    Dim BadPKMN As Integer
    Dim X As Integer
    Dim BlankPlayer As MSPlayer
    On Error Resume Next
    Set RTB = New RTBClass
    RTB.SetRTBHook Messages, ChatBox, MinWidth, MinHeight
    RTB.UseTimestamp = True
    RTB.LimitText = True
    If MusicOption = 1 Then Call StopMusic
    If SoundOption = 1 Then Call StopSound
    Me.Height = Int(MainContainer.Height * 0.8)
    Me.Width = Int(MainContainer.Width * 0.9)
    CenterWindow Me
    ChallengeLoaded = False
    UserList.Icons = MainContainer.Trainers
    UserList.SmallIcons = MainContainer.MiniTrainers
    If ServerRunning Then
        ReDim Preserve Player(MaxUsers) As MSPlayer
    Else
        ReDim Player(1) As MSPlayer
    End If
    mnuOptionsItem(7).Enabled = False
    mnuTeam.Enabled = False
    WasDisconnected = False
    FloodCheck = 0
    ModBuffer = vbNullString
    ListingBuffer = ""
    ListingPackets = 0
    ListingRcv = 0
    ExitThis = False
    ICalled = False
    ChallengePending = False
    ChallengeNumber = 0
    Battling = False
    UnloadingBattle = False
    ReDim Watching(0)
    ReDim IMQueue(0)
    WatchID = ""
    ServerIssue = 0
    YourNumber = -1
    FreeWatch = 1
    UseXOR = False
    NowSwitching = False
    If ServerRegName = "" Then ServerRegName = ServerAddress
    StatusBar1.Panels(1).Text = "Attempting to connect to server: " & ServerRegName
    Socket.RemoteHost = Trim$(ServerAddress)
    Socket.RemotePort = MainPort
    RelayServer = True
    If SoundOption = 1 Then mnuOptionsItem(0).Checked = True
    If MusicOption = 1 Then mnuOptionsItem(1).Checked = True
    If GetSpeed = 1 Then mnuOptionsItem(2).Checked = True
    mnuOptionsItem(4).Checked = VerIcons
    mnuOptionsItem(5).Checked = ColorNames
    mnuOptionsItem(3).Checked = DoMultiPaste
    Call DoInitialResize
'    Unload Loader
    On Error GoTo Failed
    StatusBar1.Panels(1).Text = "Connecting to server " & ServerRegName
    Socket.Connect
    Exit Sub
Failed:
    StatusBar1.Panels(1).Text = "Connection failed, waiting to retry..."
    Connector.Enabled = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        'If Me.Width < MinWidth Then Me.Width = MinWidth
        'If Me.Height < MinHeight Then Me.Height = MinHeight
        Command1.Top = Me.Height - 1260
        Command2.Top = Me.Height - 1260
        FCBar.Top = Me.Height - 1260
        Label1.Top = FCBar.Top + FCBar.Height
        Command1.Left = Me.Width - 1230
        Command2.Left = Me.Width - 2310
        ChatBox.Width = Me.Width - ChatBox.Left - 315
        Messages.Width = Me.Width - Messages.Left - 315
        ChatBox.Top = Me.Height - 1620
        UserList.Height = Me.Height - 1005
        Messages.Height = Me.Height - 1845
        picResizer.Height = Messages.Height
        If picResizer.Left > Command2.Left - 1700 Then
            Call UserListResize(Command2.Left - 1700 - picResizer.Left)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim X As Integer
    On Error Resume Next
    Call RTB.UnsetRTBHook
    RelayServer = False
    Connector.Enabled = False
    For X = 1 To MaxUsers
        Call KillIMWindow(X)
    Next X
    If Socket.State = sckConnected And Right(MasterServer.Caption, 12) <> "Disconnected" And Connector.Enabled = False Then
        Call SendData("EXIT:")
        Socket.Close
    End If
    If ServerIssue <> 0 Then
        Select Case ServerIssue
            Case 1
                MsgBox "Server " & ServerRegName & " has shut down.", vbInformation, "Server Quit"
            Case 2
                MsgBox "You have been kicked by Server " & ServerRegName & ".", vbCritical, "Kicked"
            Case 3
                MsgBox "The password was rejected by the server.", vbCritical, "Password Error"
            Case 4
                MsgBox "Your opponent disconnected!", vbCritical, "Disconnect"
            Case 5
                MsgBox "Server " & ServerRegName & " has adjusted the maximum number of users." & vbNewLine & "Your slot was higher than the last slot, so you were disconnected.", vbInformation, "Server Changed"
            Case 6
                MsgBox "There is already a user named " & You.Name & " logged on." & vbNewLine & "If you were disconnected, please wait a few minutes, then try again." & vbNewLine & "If you feel that somebody has stolen your nickname, please contact the server administrator.", vbCritical, "Name in Use"
            Case 7
                MsgBox "You have been auto-kicked for flooding.", vbCritical, "Floodcheck"
            Case 8
                MsgBox "You have been banned from this server." & IIf(BanMessage <> "", vbNewLine & vbNewLine & "The server provided the following message:" & vbNewLine & BanMessage, ""), vbCritical, "Banned"
            Case 9
                MsgBox "This server is not accepting new users.", vbCritical, "Can't connect"
            Case 10
                MsgBox "You have been temporarily banned from this server.  Your ban has " & BanMessage & " minutes remaining.", vbCritical, "Temp Ban"
            Case 11
                MsgBox "The user password does not match the password stored on the server." & vbNewLine & "This may mean that your user name has already been taken.  Try another name.", vbCritical, "Password Error"
            Case 12
                MsgBox "This server only allows " & BanMessage & " connection" & IIf(Val(BanMessage) = 1, "", "s") & " per IP Address.", vbCritical, "Too Many Connections"
            Case 13
                MsgBox "Your team was determined illegal.  The server returned the following message:" & vbNewLine & vbNewLine & BanMessage, vbExclamation, "Illegal Team"
        End Select
        ServerIssue = 0
    End If
    If Me.WindowState <> vbMinimized Then
        If Me.WindowState = vbMaximized Then
            SaveSetting "NetBattle", "Server Window", "Maximized", True
        Else
            SaveSetting "NetBattle", "Server Window", "Maximized", False
        End If
        SaveSetting "NetBattle", "Server Window", "Width", Me.Width
        SaveSetting "NetBattle", "Server Window", "Height", Me.Height
    End If
    ServerAddress = ""
    ServerRegName = ""
    Unload ChallengeWindow
    Unload Battle
    Unload PlayerInfoAdv
    Unload WatchForm(1)
    Unload WatchForm(2)
    Unload WatchForm(3)
    Unload WatchForm(4)
    Unload WatchForm(5)
    If TeamChangeFromMS Then
        TeamChangeFromMS = False
        TeamBuilder.SendBack = False
    Else
        Loader.RefreshBattleButtons
        Loader.Visible = True
    End If
End Sub

Private Sub UserList_DblClick()
    Call mnuPlayerItem_Click(ShiftState)
End Sub
Public Sub OpenChallenge()
    Dim Temp2 As String
    Dim Temp As Integer
    On Error Resume Next
    If ChallengeLoaded Then Exit Sub
    If NowSwitching Then Call AddMessage("You can't view player infomation while changing teams.", , , , , True): Exit Sub
    If ChallengePending Then Call AddMessage("You already have a pending challenge.", , , , , True): Exit Sub
    If Battling Then Call AddMessage("You can't view player information while already in battle.", , , , , True): Exit Sub
    ICalled = True
    If UserList.SelectedItem Is Nothing Then Exit Sub
    Temp2 = UserList.SelectedItem.Key
    Temp = Val(Right(Temp2, Len(Temp2) - 5))
    ChallengeNumber = Temp
    If ChallengeNumber = 0 Then Exit Sub
    Unload ChallengeWindow
    'If DebugMode Then MsgBox "Debug 2!"
    ChallengeWindow.Show
End Sub

Private Sub Messages_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Messages.SelText <> "" Then StatusBar1.Panels(1).Text = "Text copied to Clipboard."
End Sub

Private Sub MissingDataTimer_Timer()
    Dim Temp As String
    
    MissingDataTimer.Enabled = False
'    Temp = DataBuffer
'    DataBuffer = ""
'    If Len(Temp) <= NetChunkSize Then
'        Call DoIncoming(Temp)
'    Else
'        While Len(Temp) > 256
'            Call DoIncoming(left(Temp, 256))
'            Temp = Right(Temp, Len(Temp) - 256)
'        Wend
'        Call DoIncoming(Temp)
'    End If
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileToUse As String
    Dim FileNum As Integer
    Dim Temp As String
    FileNum = FreeFile
    Select Case Index
        Case 0
            With MainContainer.FileBox
                .DialogTitle = "Save Log File"
                .Flags = cdlOFNOverwritePrompt
                .CancelError = True
                .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                .DefaultExt = ".txt"
                .FileName = ""
                Temp = GetSetting("NetBattle", "Options", "InitDir", "")
                If Temp <> "" Then .InitDir = Temp
                On Error GoTo Cancelled
                .ShowSave
                FileToUse = .FileName
                SaveSetting "NetBattle", "Options", "InitDir", Left$(FileToUse, InStrRev(FileToUse, "\"))
            End With
            Open FileToUse For Output As #FileNum
            Print #FileNum, Messages.Text
            Close #FileNum
            StatusBar1.Panels(1).Text = "Log saved"
        Case 2
            Call Command2_Click
    End Select
Cancelled:
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
            Else
                mnuOptionsItem(1).Checked = True
                MusicOption = 1
            End If
            SaveSetting "NetBattle", "Options", "Music", MusicOption
        Case 2
            If mnuOptionsItem(2).Checked Then
                mnuOptionsItem(2).Checked = False
                GetSpeed = 0
            Else
                mnuOptionsItem(2).Checked = True
                GetSpeed = 1
            End If
            SaveSetting "NetBattle", "Options", "GetSpeed", GetSpeed
        Case 3
            DoMultiPaste = Not DoMultiPaste
            mnuOptionsItem(3).Checked = Not mnuOptionsItem(3).Checked
            SaveSetting "NetBattle", "Options", "DoMultiPaste", DoMultiPaste
        Case 4
            VerIcons = Not VerIcons
            mnuOptionsItem(4).Checked = Not mnuOptionsItem(4).Checked
            SaveSetting "NetBattle", "Options", "VerIcons", VerIcons
            Call RefreshListing
        Case 5
            ColorNames = Not ColorNames
            mnuOptionsItem(5).Checked = Not mnuOptionsItem(5).Checked
            SaveSetting "NetBattle", "Options", "ColorNames", ColorNames
            Call RefreshListing
        Case 7
          '  If Player(YourNumber).BattlingWith > 0 And Player(YourNumber).BattlingWith < 1025 Then
          '      MsgBox "Can't be Away while in a battle!", vbCritical, "Error"
         '       Exit Sub
          '  End If
          '  If TeamChangeFromMS Then TeamBuilder.SendBack = False
         '  If Player(YourNumber).BattlingWith = 1025 Then
            
            
           ' Connector.Enabled = True
            
            
    
          '  End If
         '   If Player(YourNumber).BattlingWith = 0 Then
            
         '   End If
    End Select
End Sub

Private Sub mnuPlayerItem_Click(Index As Integer)
    Dim Temp2 As String
    Dim Temp As Integer
    Dim X As Integer
    
    Temp2 = UserList.SelectedItem.Key
    Temp = Val(Right(Temp2, Len(Temp2) - 5))
    
    Select Case Index
        Case 0
            Call OpenChallenge
        Case 1
            If Player(Temp).BattlingWith = 0 Or Player(Temp).BattlingWith = 1025 Then
                Call AddMessage("This player is not battling.")
                Exit Sub
            End If
            For X = 1 To UBound(Watching)
                If Watching(X) = Temp Or Watching(X) = Player(Temp).BattlingWith Then
                    Call AddMessage("You are already watching that battle.")
                    Exit Sub
                End If
            Next X
            If Temp = YourNumber Or Player(Temp).BattlingWith = YourNumber Then
                Call AddMessage("You cannot watch your own battle.")
                Exit Sub
            End If
            mnuPlayerItem(1).Enabled = False
            Call SendData("REQW:" & CStr(Temp))
        Case 2
            If Temp <> YourNumber And IMWindowID(Temp) = 0 Then
                Call Code.NewIMWindow(Player(Temp).Name, Temp, Player(Temp).Picture)
                IMWindowFlash(IMWindowID(Temp)) = False
            End If
        Case 5
            Call SendData("MKCK:" & Temp)
        Case 3
            If LookingUp = "" Then PlayerInfoAdv.Show vbModeless, MainContainer
            PlayerInfoAdv.Text1.Text = Player(Val(Temp)).Name
            PlayerInfoAdv.cmdLookup.SetFocus
            Call PlayerInfoAdv.cmdLookup_Click
'            Call SendData("MBAN:" & Temp)
'        Case 5
'            Call SendData("MBN2:" & Temp)
    End Select
End Sub

Sub DoIncoming(ByVal Temp As String)
    Dim Answer As Integer
    Dim Temp2 As String
    Dim TempString As String
    Dim P1 As Integer
    Dim P2 As Integer
    'Dim ServerVersion As String
    Dim ActiveUsers As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Trimmed As String
    Dim BlankPlayer As MSPlayer
    Dim Command As String * 5
    Dim Data As String
    Dim PNum As Byte
    Dim BNum As Byte
    Dim Worked As Boolean
    Dim B As Boolean
    Dim L As Long
    
    Select Case Left(Temp, 5)
    Case "REQN:", "RPWD:", "BANU:", "NOIP:"
        B = True
    Case Else
        B = False
    End Select
    If Not B And UseXOR Then Temp = XORDecrypt(Temp)
    'Temp = Trim(Temp)
    'Debug.Print "MasterServer.DoIncoming: " & Temp
    '>>> Call WriteDebugLog("Received: " & Temp)
    Call AddMessage(Temp, True, , , , True)
    Command = ChopString(Temp, 5)
    Data = Temp
    Call WriteDebugLog("MasterServer.DoIncoming: " & Command & Data)
    Select Case Command
        'Relay messages (for battle)
        Case "RELAY"
            If Battling Then Call Battle.AddToQueue(Data)
        'Relay messages (for watches)
        Case "WATCH"
            BNum = Dec(ChopString(Data, 3))
            For X = 1 To UBound(WatchForm)
                If WatchLoaded(X) Then
                    If WatchForm(X).WatchID = BNum Then Exit For
                End If
            Next X
            If X <= UBound(WatchForm) Then Call WatchForm(X).AddToQueue(Data)
        'Request (server) PassWorD
        Case "RPWD:"
            UseXOR = CBool(ChopString(Data, 1))
            ServerVersion = Trim(ChopString(Data, 10))
            MaxUsers = Dec(ChopString(Data, 3))
            ReDim Preserve Player(MaxUsers) As MSPlayer
            FloodTolerance = Dec(ChopString(Data, 2))
            ActiveUsers = Dec(ChopString(Data, 3))
            Temp = Trim$(ChopString(Data, 20))
            If Len(Temp) <> 0 Then ServerAddress = Temp
            Call AddMessage("Server " & ServerAddress & " - NetBattle v." & ServerVersion, , , vbBlue, True)
            Call AddMessage("Currently " & ActiveUsers & " trainer(s) online, with a maximum of " & MaxUsers & " trainers.", , , vbBlue, True)
            FCBar.Max = FloodTolerance
            Call AddMessage("Flood count is set to " & FloodTolerance, , , vbRed, True)
            Call AddMessage("Your team's power is: " & Ranking & "%", , , , True)
            PasswordBoxTitle = "Server Password"
            PasswordBoxCaption = "This server is password protected." & vbNewLine & "Please enter the password."
            PWWindow.Show 1
            Call SendData("USER:" & PrepareUserInfo(Data, ServerPassword))
        '(Server) PassWorD Refused
        Case "PWDR:"
            ServerIssue = 3
            ExitThis = True
            Call SendData("EXIT:")
        'REQuest Name
        Case "REQN:"
            UseXOR = CBool(ChopString(Data, 1))
            ServerVersion = Trim(ChopString(Data, 10))
            MaxUsers = Dec(ChopString(Data, 3))
            ReDim Preserve Player(MaxUsers) As MSPlayer
            FloodTolerance = Dec(ChopString(Data, 2))
            ActiveUsers = Dec(ChopString(Data, 3))
            Temp = Trim$(ChopString(Data, 20))
            If Len(Temp) <> 0 Then ServerAddress = Temp
            Call AddMessage("Server " & ServerAddress & " - NetBattle v." & ServerVersion, , , vbBlue, True)
            Call AddMessage("Currently " & ActiveUsers & " trainer(s) online, with a maximum of " & MaxUsers & " trainers.", , , vbBlue, True)
            Call SendData("USER:" & PrepareUserInfo(Data))
            FCBar.Max = FloodTolerance
            Call AddMessage("Flood count is set to " & FloodTolerance, , , vbRed, True)
            Call AddMessage("Your team's power is: " & Ranking & "%", , , , True)
        'NAMe Refused
        '(Duplicate player)
        Case "NAMR:"
            ServerIssue = 6
            ExitThis = True
            Call SendData("EXIT:")
        'IP is banned
        Case "BANU:"
            ServerIssue = 8
            ExitThis = True
            BanMessage = Data
            Call SendData("EXIT:")
        'Temp Banned
        Case "TBAN:"
            ServerIssue = 10
            ExitThis = True
            BanMessage = Data
            Call SendData("EXIT:")
        'Server doesn't accept new players
        Case "NNPL:"
            ServerIssue = 9
            ExitThis = True
            Call SendData("EXIT:")
        'Server is full
        Case "BUSY:"
            Socket.Close
            MsgBox "Server is full - try again later.", vbCritical, "Error"
            Unload Me
        'ReQuest TeaM
        '(New format in 0.9.00)
        Case "RQTM:"
            Temp2 = Data
            P1 = 1
            P2 = InStr(P1, Temp2, ",")
            YourNumber = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Authority = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Wins = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Losses = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Ties = Mid(Temp2, P1, P2 - P1)
            Player(YourNumber).Disconnect = Right(Temp2, Len(Temp2) - P2)
            Player(YourNumber).Name = You.Name
            Player(YourNumber).Version = You.ProgVersion
            Player(YourNumber).Extra = You.Extra
            Player(YourNumber).Picture = You.Picture
            Player(YourNumber).Rank = Val(Ranking)
            Call RefreshAuth
            For X = 0 To 4
                Player(YourNumber).Compatibility(X) = Compatibility(X)
            Next
            'Player(YourNumber).Active = True
            Player(YourNumber).BattlingWith = 0
            For X = 1 To 6
                Player(YourNumber).PKMN(X) = PKMN(X).No
                Player(YourNumber).PKMNImage(X) = PKMN(X).Image
            Next
            Call RefreshListing
            Call SendData("TEAM:" & PrepareTeamInfo)
        'The initial player listing
        Case "/WHO:"
            If ListingBuffer = "" Then
                ListingPackets = Bin2Dec(Chr2Bin(Mid(Data, 1, 2)))
            End If
            ListingRcv = ListingRcv + 1
            StatusBar1.Panels(1).Text = "Receiving Player List: " & CStr(Round(ListingRcv / ListingPackets, 2) * 100) & "%"
            ListingBuffer = ListingBuffer & Left$(Data, 200)
            If Len(Data) <= 200 Then
                StatusBar1.Panels(1).Text = "Interpreting Player Information..."
                Temp = ChopString(ListingBuffer, 4)
                L = Bin2Dec(Chr2Bin(Mid(Temp, 3, 2)))
                MainContainer.Compressor.DecompressString ListingBuffer, L
                L = Len(ListingBuffer)
                Do Until L = 0
                    L = L - 1
                    X = Asc(Left$(ListingBuffer, 1))
                    If L < X Then
                        Call SendData("EXIT:")
                        Exit Sub
                    End If
                    ListingBuffer = Right$(ListingBuffer, L)
                    L = L - X
                    Call InterpretPlayerData(Left$(ListingBuffer, X), True)
                    ListingBuffer = Right$(ListingBuffer, L)
                Loop
                Call RefreshListing
                Call ParsePlayerList
                mnuOptionsItem(7).Enabled = True
                mnuTeam.Enabled = True
            End If
        'Serverside DB Mod
        Case "DBMD:"
            ModBuffer = ModBuffer & Left$(Data, 200)
            If Len(Data) <> 201 Then
                Call AddMessage("WARNING!  This server uses a modified database.", , , vbRed, True)
                If Len(DBModStr) > 0 And PKMN(1).GameVersion = nbModAdv Then
                    If StrComp(ModBuffer, DBModStr) <> 0 Then
                        MsgBox "This server's Database Mod differs from the one currently loaded.  Your team may be seen as illegal.  If it is, open Team Builder and it will automatically correct your team.", vbExclamation, "Database Mod"
                    End If
                End If
                DBModStr = ModBuffer
                ModBuffer = vbNullString
                DBModName = ServerAddress
                Call ApplyDBMod
                Call SaveDBMod
            End If
        'REQuest Version
        Case "REQV:"
            Call SendData("VERS:" & You.ProgVersion)
        'REQuest Picture
        Case "REQP:"
            Call SendData("PICT:" & Str(You.Picture))
        'Request USer Password
        Case "RUSP:"
            PasswordBoxTitle = "User Password"
            PasswordBoxCaption = "Your user password is not saved." & vbNewLine & "Please enter the password."
            If SavedPassword = "" Then
                PWWindow.Show 1
            Else
                ServerPassword = SavedPassword
            End If
            Call SendData("UPWS:" & MD5(ServerPassword))
        'User PassWord Refused
        Case "UPWR:"
            ServerIssue = 11
            ExitThis = True
            Call SendData("EXIT:")
        'Too many connections
        Case "NOIP:"
            ServerIssue = 12
            ExitThis = True
            BanMessage = Data
            Call SendData("EXIT:")
        'Authority
        Case "AUTH:" 'THIS ONE IS STILL USED IN 0.9.0 FOR AUTH CHANGES
            X = Dec(ChopString(Data, 3))
            Player(X).Authority = Val(Data)
            Call RefreshListing
            If X = YourNumber Then Call RefreshAuth
        'Your NUMber
        Case "YNUM:"
            Temp2 = Data
            P1 = 1
            P2 = InStr(P1, Temp2, ",")
            YourNumber = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Wins = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Losses = Val(Mid(Temp2, P1, P2 - P1))
            P1 = P2 + 1
            P2 = InStr(P1, Temp2, ",")
            Player(YourNumber).Ties = Mid(Temp2, P1, P2 - P1)
            Player(YourNumber).Disconnect = Right(Temp2, Len(Temp2) - P2)
            Player(YourNumber).Name = You.Name
            Player(YourNumber).Version = You.ProgVersion
            Player(YourNumber).Extra = You.Extra
            Player(YourNumber).Picture = You.Picture
            Player(YourNumber).Rank = Val(Ranking)
            For X = 0 To 4
                Player(YourNumber).Compatibility(X) = Compatibility(X)
            Next
            Player(YourNumber).Active = True
            Player(YourNumber).BattlingWith = 0
            Player(YourNumber).ShowTeam = AllowViewing
            For X = 1 To 6
                Player(YourNumber).PKMN(X) = PKMN(X).No
                Player(YourNumber).PKMNImage(X) = PKMN(X).Image
            Next
            Call RefreshListing
            Call SendData("LIST:")
        'Player LiST
        Case "PLST:"
            StatusBar1.Panels(1).Text = "Received Trainer List"
            Call ParsePlayerList(Data)
        'Player iNFO
        Case "PNFO:"
            StatusBar1.Panels(1).Text = "Received Trainer Data"
            Call InterpretPlayerData(Data, Not MsgToggle(1))
            Call RefreshListing
            Call ParsePlayerList
        'Player UPDate
        Case "PUPD:"
            StatusBar1.Panels(1).Text = "Received Trainer Data"
            If ChopString(Data, 1) = "T" Then
                X = Asc(Mid(Data, 1, 1))
                If ChallengeNumber = X Then
                    On Error Resume Next
                    Unload ChallengeWindow
                End If
                'OldName = Player(X).Name
                TempString = Trim(Mid(Data, 29, 20))
                If X = YourNumber And OldName <> You.Name Then
                    If MsgToggle(10) Then Call AddMessage(OldName & " has changed teams and is now known as " & TempString & ".")
                ElseIf TempString <> Player(X).Name Then
                    If MsgToggle(10) Then Call AddMessage(Player(X).Name & " has changed teams and is now known as " & TempString & ".")
                    If IMWindowID(X) > 0 Then
                         MainContainer.IMWindowList.Buttons("U:" & Player(X).Name).Caption = TempString
                         MainContainer.IMWindowList.Buttons("U:" & Player(X).Name).Key = "U:" & TempString
                         IMWindowArray(IMWindowID(X)).Caption = TempString
                    End If
                Else
                    If MsgToggle(9) Then Call AddMessage(TempString & " has changed teams.")
                End If
                If X = YourNumber Then
                    mnuTeam.Enabled = True
                    NowSwitching = False
                End If
            End If
            Call InterpretPlayerData(Data, True)
            RefreshListing
        'New PLaYer
        Case "NPLY:"
            StatusBar1.Panels(1).Text = "New Trainer signed on!"
            Player(Val(Data)).Active = True
            If SoundOption = 1 Then Call PlaySound(1)
            Call ParsePlayerList
        'Player DISconnect
        Case "PDIS:"
            P1 = Val(Data)
            If P1 = YourNumber Then
                Socket.Close
                Unload Me
            Else
                StatusBar1.Panels(1).Text = "Trainer signed off"
                If Len(Player(P1).Name) > 0 And MsgToggle(2) Then
                    If Player(P1).BattlingWith > 0 And Player(P1).BattlingWith < 1025 Then
                        Call AddMessage(Player(P1).Name & " has disconnected from battle.", False)
                    Else
                        Call AddMessage(Player(P1).Name & " signed off.", False)
                    End If
                End If
                If IMWindowID(P1) > 0 Then
                    Call Code.KillIMWindow(P1)
                End If
                On Error Resume Next
                If ChallengeNumber = P1 Then
                    ChallengePending = False
                    ChallengeNumber = 0
                    ChallengeWindow.Closer.Enabled = True
                End If
                X = Player(P1).BattlingWith
                Temp = Player(P1).Name
                Player(P1) = BlankPlayer
                Call RefreshListing
                If Player(YourNumber).BattlingWith = P1 And YourNumber = X Then
                    Battle.OppDiscon
                End If
                For Y = 1 To 5
                    If WatchLoaded(Y) Then
                        If WatchForm(Y).WatchP1 = P1 Then
                            Call WatchForm(Y).OppDiscon(1)
                        ElseIf WatchForm(Y).WatchP2 = P1 Then
                            Call WatchForm(Y).OppDiscon(2)
                        End If
                    End If
                Next Y
'                If Battling Then
'                    If Battle.IsWatching(P1) Then
'                        Call Battle.AddToQueue("DONW:" & "1" & Chr$(P1) & Temp)
'                    End If
'                End If
            End If
        'Player returned from Away/Battling
        Case "BACK:"
            P1 = Val(Data)
            If Player(P1).BattlingWith = 0 Then Exit Sub
            If Player(P1).BattlingWith = 1025 Then
                If MsgToggle(7) Then Call AddMessage(Player(P1).Name & " has returned.")
            Else
                If MsgToggle(5) Then Call AddMessage(Player(P1).Name & " is done battling.")
                For X = 1 To 5
                    If WatchLoaded(X) Then
                        If WatchForm(X).WatchP1 = P1 Or WatchForm(X).WatchP2 = P1 Then
                            Unload WatchForm(X)
                        End If
                    End If
                Next X
                If Player(P1).BattlingWith = YourNumber And Player(YourNumber).BattlingWith = P1 And UnloadingBattle = False Then
                    Player(P1).BattlingWith = 0
                    Call Battle.AddToQueue("BOVER")
                End If
            End If
            Player(P1).BattlingWith = 0
            If P1 = YourNumber Then mnuOptionsItem(7).Checked = False
            Call RefreshListing
        Case "RPUD:"
            Call SendData("USER:" & PrepareUserInfo(Data))
        'Chat MeSsaGe
        Case "CMSG:"
            Temp2 = Data
            X = InStr(1, Data, ":")
            If Left(Data, 4) = "*** " Then
                Data = ApplyCSFilter(Data)
                Call AddMessage(Data, , , &HC000C0)
            ElseIf X > 0 Then
                Temp2 = Left$(Data, X - 1)
                Mid(Data, X) = ApplyCSFilter(Mid$(Data, X))
                If Temp2 = Player(YourNumber).Name Then
                    Call AddMessage(Data, , ":", vbRed, True)
                Else
                    For X = 1 To MaxUsers
                        If Player(X).Name = Temp2 Then Exit For
                    Next X
                    If X = MaxUsers + 1 Then
                        Call AddMessage(Data, , ":", &H99AA00, True)
'                    ElseIf Player(X).Authority > 1 Then
'                        Call AddMessage(Data, , ":", vbBlue, True, , True)
                    Else
                        Call AddMessage(Data, , ":", vbBlue, True)
                    End If
                End If
            Else
                Data = ApplyCSFilter(Data)
                Call AddMessage(Data)
            End If
            If SoundOption = 1 Then Call PlaySound(2)
        'IM CHat
        Case "IMCH:"
            Call AddToIMQueue("IMCH:" & Data)
        'Server MeSsaGe
        Case "SMSG:"
            Temp2 = Data
            Call AddMessage("Welcome Message: " & Temp2, , ":", vbRed, True)
            If SoundOption = 1 Then Call PlaySound(2)
        'SerVeR QUit
        Case "SVRQU"
            Call AddMessage("SERVER IS SHUTTING DOWN", , , vbRed, True)
            Call SendData("EXIT:")
            ServerIssue = 1
            ExitThis = True
        'Player kicked
        Case "KICK:"
            Call AddMessage("SERVER HAS KICKED TRAINER: " & Player(Val(Data)).Name, , , vbRed, True)
        Case "KCKU:"
            Call SendData("EXIT:")
            ExitThis = True
            ServerIssue = 2
        Case "ILLM:"
            ServerIssue = 13
            ExitThis = True
            BanMessage = Data
            Call SendData("EXIT:")
        'Max User ChanGe
        Case "MUCG:"
            StatusBar1.Panels(1).Text = "Maximum # of users has changed."
            Call AddMessage("Server has changed maximum # of users to " & MaxUsers, , , vbRed, True)
            MaxUsers = Val(Data)
            If YourNumber > MaxUsers Then
                Call SendData("EXIT:")
                ExitThis = True
                ServerIssue = 5
            End If
        'The following few are for challenging.
        Case "CHLN:"
            X = Dec(ChopString(Data, 3))
            If ChallengeLoaded Or Battling Or NowSwitching Or Player(YourNumber).BattlingWith = 1025 Then
                Call SendData("PBSY:" & X)
            Else
                ChallengeMode = Dec(ChopString(Data, 1))
                ChallTerrain = Dec(ChopString(Data, 1))
                Call ReadBinArray(Dec(ChopString(Data, 8)), RuleSelected)
                Player(X).Speed = Data
                ChallengeNumber = X
                ICalled = False
                On Error Resume Next
                Unload ChallengeWindow
                ChallengeWindow.Show
                'ChallengeWindow.Caption = ChallengeWindow.Caption & " - Ping Speed: " & Player(ChallengeNumber).Speed
                'ChallengeWindow.SpeedBar.Caption = Player(ChallengeNumber).Speed
                'ChallengeWindow.SpeedBar.Value = Val(Player(ChallengeNumber).Speed)
                If SoundOption = 1 And MusicOption = 0 Then Call PlaySound(3)
            End If
        Case "PBSY:"
            X = Val(Data)
            Call AddMessage(Player(X).Name & " can't be challenged at this time." & vbNewLine & "They may be in battle, switching teams, waiting for a challenge, or have the challenge window open.", , , , , True)
            ChallengePending = False
            If ICalled Then Unload ChallengeWindow
        Case "SBSY:"
            X = Val(Data)
            Call AddMessage(Player(X).Name & " is already in a battle.", , , , , True)
            ChallengeWindow.OKButton.Enabled = True
            ChallengePending = False
            Unload ChallengeWindow
        Case "PREF:"
            X = Val(Data)
            Call AddMessage(Player(X).Name & " has refused the challenge.", , , , , True)
            ChallengeWindow.OKButton.Enabled = True
            ChallengePending = False
            Unload ChallengeWindow
        Case "SREF:"
            X = Val(Data)
            ChallengeWindow.OKButton.Enabled = True
            ChallengePending = False
            Unload ChallengeWindow
        Case "PACC:"
            X = Val(Data)
            If ChallengePending = False Then
                Call SendData("PCAN:" & X)
            ElseIf NowSwitching Or ChallengeNumber <> X Then
                Call SendData("PBSY:" & X)
            Else
                Call AddMessage(Player(X).Name & " has accepted the challenge!", , , , , True)
                Player(YourNumber).BattlingWith = ChallengeNumber
                StatusBar1.Panels(1).Text = "Challenge Accepted!"
                
            End If
        'Start Battle
        'PNum will be used for multi-battling, not used yet.
        Case "SBAT:"
            PNum = Asc(ChopString(Data, 1))
            BattleTemp = Data
            mnuTeam.Enabled = False
            On Error Resume Next
            Unload Battle
            OpenedAsReplay = False
            Battle.Show
        'Start Watch
        Case "SWAT:"
            For X = 1 To 5
                If Not WatchLoaded(X) Then Exit For
            Next X
            If X > UBound(WatchLoaded) Then
                Call SendData("DONW:" & Dec(ChopString(Data, 3)))
                Exit Sub
            End If
            WatchID = Data & FixedHex(X, 3)
            WatchLoaded(X) = True
            On Error Resume Next
            Unload WatchForm(X)
            OpenedAsReplay = False
            WatchForm(X).Show
            mnuPlayerItem(1).Enabled = True
        Case "WREF:"
            Select Case Val(Data)
            Case 1
                Call AddMessage("This player is not battling.")
            Case 2
                Call AddMessage("Spectators have been disallowed for this battle.")
            Case 3
                Call AddMessage("You are already watching this battle.")
            Case 4
                Call AddMessage("This battle has not finished initializing.  Please wait a few seconds and try again.")
            Case 5
                Call AddMessage("Either this battle started under a now-outdated database, or your database is outdated.  If the latter, sign off and on.")
            End Select
            mnuPlayerItem(1).Enabled = True
        Case "PCAN:"
            X = Val(Data)
            Call AddMessage(Player(X).Name & " has cancelled the challenge.", , , , , True)
            ChallengePending = False
            Unload ChallengeWindow
        'Silent kick
        'Mostly used when new stuff is introduced on the server.
        Case "BOOT:"
            'ExitThis = True
            Call SendData("EXIT:")
        'So the server doesn't time you out.
        Case "PING:"
            Call SendData("PONG:")
        'FloodcounT ChanGe
        Case "FTCG:"
            X = Val(Data)
            Call AddMessage("Server changed floodcount to " & X, , , vbRed, True)
            FloodTolerance = X
            FCBar.Value = 0
            FCBar.Max = FloodTolerance
            Call ChangeFCBar(FloodCheck)
        'Players aRe BuSy
        Case "PRBS:"
            Temp2 = Data
            P1 = Val(Left(Temp2, InStr(1, Temp2, ",") - 1))
            P2 = Val(Right(Temp2, Len(Temp2) - InStr(1, Temp2, ",")))
            'If P1 <> YourNumber Then Player(P1).BattlingWith = P2
            'If P2 <> YourNumber Then Player(P2).BattlingWith = P1
            Player(P1).BattlingWith = P2
            Player(P2).BattlingWith = P1
            If MsgToggle(3) Then Call AddMessage(Player(P1).Name & " and " & Player(P2).Name & " are battling.")
            Call RefreshListing
        'Player is away via menu option
        Case "AWAY:"
            Temp2 = Data
            Player(Val(Temp2)).BattlingWith = 1025
            If Val(Temp2) = YourNumber Then mnuOptionsItem(7).Checked = True
            If MsgToggle(6) Then Call AddMessage(Player(Val(Temp2)).Name & " is away.")
            Call RefreshListing
        Case "PSPD:"
            X = Asc(ChopString(Data, 1))
            Player(X).Speed = Data
            If Player(X).Speed = "WAIT" Then
                Call SendData("GETS:" & Chr$(X))
                Exit Sub
            End If
            If ChallengeNumber = X Then
                ChallengeWindow.Caption = ChallengeWindow.Caption & " - Ping Speed: " & Player(ChallengeNumber).Speed
                'ChallengeWindow.SpeedBar.Caption = Player(X).Speed
                'ChallengeWindow.SpeedBar.Value = Val(Player(x).Speed)
            End If
        Case "MKCK:"
            Temp2 = Data
            X = InStr(1, Temp2, ":")
            P1 = Left(Temp2, X - 1)
            P2 = Right(Temp2, Len(Temp2) - X)
            If Player(P1).Authority = 2 Then
                TempString = "Mod "
            Else
                TempString = "Admin "
            End If
            Call AddMessage(TempString & Player(P1).Name & " has kicked " & Player(P2).Name, , , vbRed, True)
        Case "MBAN:"
            X = Asc(ChopString(Data, 1))
            Call AddMessage("Admin " & Player(X).Name & " has banned " & Data, , , vbRed, True)
        Case "BRLT:"
            If LookingUp <> "" Then
                With PlayerInfoAdv
                    Select Case Val(Data)
                    Case 0:
                        .StatusBar.SimpleText = "Ban successful."
                        .txtInfo(2).Text = "Banned"
                        .Command(0).Enabled = False
                        .Command(2).Enabled = False
                    Case 1: .StatusBar.SimpleText = "Request rejected: User already banned."
                    Case 2: .StatusBar.SimpleText = "Request rejected: User not found."
                    Case 3: .StatusBar.SimpleText = "Request rejected: Authority conflict."
                    Case 4:
                        .StatusBar.SimpleText = "TempBan successful."
                        .txtInfo(2).Text = "TempBanned"
                        .Command(0).Enabled = False
                        .Command(2).Enabled = True
                    End Select
                End With
            End If
        Case "LOOK:"
            If LookingUp <> "" Then
                If Left(Data, 1) = "|" Then
                    PlayerInfoAdv.StatusBar.SimpleText = "Server reports no such user."
                    PlayerInfoAdv.cmdLookup.Enabled = True
                Else
                    Temp = ChopString(Data, InStr(1, Data, "|") - 1)
                    If LCase(LookingUp) = LCase(Temp) Then
                        LookingUp = Temp
                        With PlayerInfoAdv
                            .txtInfo(0).Text = LookingUp
                            Y = Val(Mid(Data, 2, 1))
                            Select Case Y
                            Case 0: .txtInfo(2) = "Banned"
                            Case 1: .txtInfo(2) = "TempBanned"
                            Case 2: .txtInfo(2) = "Normal User"
                            Case 3: .txtInfo(2) = "Moderator"
                            Case 4: .txtInfo(2) = "Administrator"
                            End Select
                            X = IIf(Player(YourNumber).Authority >= Y, 1, 0)
                            .Command(0).Enabled = (X = 1 And Y > 1)
                            .Command(2).Enabled = (X = 1 And Y > 0)
                            .txtInfo(3) = Mid(Data, 3, 21)
                            Call ChopString(Data, 23)
                            Temp = ChopString(Data, InStr(1, Data, "|") - 1)
                            .txtInfo(1).Text = IIf(.PlayerOnline, "Currently Online", "Last Online on " & Temp)
                            Call ChopString(Data, 1)
                            If .PlayerOnline Then
                                .txtInfo(4) = Data
                                .Command(1).Enabled = (X = 1)
                            Else
                                .txtInfo(4) = "[Player not connected]"
                                .Command(1).Enabled = False
                            End If
                            .cmdLookup.Enabled = True
                            .cmdLookup.SetFocus
                            .StatusBar.SimpleText = "Query successful."
                        End With
                    End If
                End If
            End If
        Case "ALIA:"
            With PlayerInfoAdv
                If Len(Data) = 240 Then
                    For X = 1 To 221 Step 20
                        .NickList.ListItems.Add , , Trim$(Mid$(Data, X, 20))
                    Next X
                Else
                    Do While Len(Data) > 1
                        .NickList.ListItems.Add , , Trim$(ChopString(Data, 20))
                    Loop
                    .cmdALook.Enabled = True
                    For X = 0 To 2
                        .Command(X).Enabled = True
                    Next X
                    .cmdLookup.Enabled = True
                    .StatusBar.SimpleText = "Query successful."
                End If
            End With
        Case "BANL:"
            If LookingUp <> "" Then
                With PlayerInfoAdv
                    .iBanList = .iBanList & Left(Data, 200)
                    If Len(Data) <= 200 Then Call .ProcessBanList
                End With
            End If
        Case "MTBN:"
            X = Asc(ChopString(Data, 1))
            Y = Str2Int(ChopString(Data, 2))
            If Player(X).Authority = 3 Then
                TempString = "Admin "
            Else
                TempString = "Mod "
            End If
            Call AddMessage(TempString & Player(X).Name & " has TempBanned " & Data & " for " & CStr(Y) & " minutes", , , vbRed, True)
        Case "MASS:" 'Mass Message from the server registry
            Call AddMessage("NETWORK-WIDE MESSAGE: " & Data, , ":", &H80FF&, True)
    End Select
End Sub
Public Sub AddToIMQueue(Data As String)
    ReDim Preserve IMQueue(UBound(IMQueue) + 1)
    IMQueue(UBound(IMQueue)) = Data
End Sub
Sub RefreshListing()
    Dim X As Integer
    'Dim Index As Integer
    Dim TempItem As ListItem
    Dim Temp As String
    'On Error Resume Next
    Call SetRedraw(Me.hWnd, False)
    For X = 1 To MaxUsers
        If Player(X).Name <> "" Then
            If Player(X).Authority = 0 Then Player(X).Authority = 1
            If PListed(X) Then
                Set TempItem = UserList.ListItems("USER: " & CStr(X))
            Else
                Set TempItem = UserList.ListItems.Add(, "USER: " & CStr(X), Player(X).Name)
            End If
            With TempItem
                Temp = Trim(Player(X).Name & " " & String(Player(X).Authority - 1, "*"))
                If .Text <> Temp Then .Text = Temp
                If VerIcons Then
                    If .SmallIcon <> Player(X).GameVersion + 16 Then .SmallIcon = Player(X).GameVersion + 16
                Else
                    If Player(X).Picture > 0 And .SmallIcon <> Player(X).Picture Then .SmallIcon = Player(X).Picture
                End If
                If ColorNames Then
                    If Player(X).BattlingWith > 0 And Player(X).BattlingWith < 1025 Then
                        .ForeColor = vbBlue
                    Else
                        Select Case Player(YourNumber).GameVersion
                            Case 0, 5, 1, 6
                                Select Case Player(X).GameVersion
                                    Case 0, 5, 1, 6
                                    .ForeColor = vbBlack
                                    Case 2 To 4
                                        .ForeColor = vbRed
                                End Select
                            Case 2 To 4
                                Select Case Player(X).GameVersion
                                    Case 0, 5, 1, 6
                                        .ForeColor = vbRed
                                    Case 2 To 4
                                        .ForeColor = vbBlack
                                End Select
                        End Select
                    End If
                Else
                    .ForeColor = vbBlack
                End If
                If Player(X).BattlingWith = 1025 Then
                    .Ghosted = True
                    .ToolTipText = "Away"
                ElseIf Player(X).BattlingWith > 0 Then
                    .Ghosted = True
                    If Player(Player(X).BattlingWith).BattlingWith <> X Then
                        .ToolTipText = "No opponent"
                    Else
                        .ToolTipText = "Battling with " & Player(Player(X).BattlingWith).Name
                    End If
                Else
                    .ToolTipText = "Not battling"
                    .Ghosted = False
                End If
                'UserList.Refresh
            End With
        Else
            If PListed(X) Then UserList.ListItems.Remove "USER: " & CStr(X)
        End If
    Next X
    UserList.Sorted = True
    Call SetRedraw(Me.hWnd, True)
End Sub

Private Function PListed(PNum As Integer) As Boolean
    Dim X As Integer
    On Error Resume Next
    Err.Number = 0
    X = UserList.ListItems("USER: " & CStr(PNum)).Index
    PListed = (Err.Number = 0)
End Function

Sub ParsePlayerList(Optional ByVal Listing As String = "")
    Dim X As Integer
    Dim Y As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim Request As Integer
    
    If Listing <> "" Then
        For X = 1 To Len(Listing)
            Player(Asc(Mid(Listing, X, 1))).Active = True
        Next X
    End If

    Request = -1
    For X = 1 To MaxUsers
        If Player(X).Active = True And Player(X).Name = "" And X <> YourNumber Then Request = X: Exit For
    Next X
    Call RefreshListing
    If Request > -1 Then
        Call SendData("PREQ:" & Request)
    Else
        StatusBar1.Panels(1).Text = "All online players displayed."
    End If
End Sub

Sub InterpretPlayerData(ByVal BuildString As String, ByVal SkipMessage As Boolean)
    Dim Number As Integer
    Dim X As Byte
    Dim G As Byte
    Dim TempPKMN As Pokemon
    Dim Temp As String
    Temp = Chr2Bin(ChopString(BuildString, 28))
    Number = Bin2Dec(ChopString(Temp, 8))
    With Player(Number)
        .GameVersion = Bin2Dec(ChopString(Temp, 3))
        .Picture = Bin2Dec(ChopString(Temp, 4))
        .GFXVer = Bin2Dec(ChopString(Temp, 4))
        .Authority = Bin2Dec(ChopString(Temp, 2))
        .ShowTeam = Bin2Bool(Bin2Dec(ChopString(Temp, 1)))
        .StadiumOK = Bin2Bool(Bin2Dec(ChopString(Temp, 1)))
        For X = 1 To 6
            .PKMN(X) = Bin2Dec(ChopString(Temp, 9))
            TempPKMN.No = .PKMN(X)
            TempPKMN.UnownLetter = Bin2Dec(ChopString(Temp, 5))
            TempPKMN.Shiny = CBool(Bin2Dec(ChopString(Temp, 1)))
            .PKMNImage(X) = ChooseImage(TempPKMN, .GFXVer)
        Next
        .Wins = Bin2Dec(ChopString(Temp, 16))
        .Losses = Bin2Dec(ChopString(Temp, 16))
        .Ties = Bin2Dec(ChopString(Temp, 16))
        .Disconnect = Bin2Dec(ChopString(Temp, 16))
        Call ReadBinArray(Bin2Dec(ChopString(Temp, 16)), .Compatibility)
        .Rank = Bin2Dec(ChopString(Temp, 16))
        .BattlingWith = Bin2Dec(ChopString(Temp, 11))
        .Name = Trim(ChopString(BuildString, 20))
        .Version = Trim(ChopString(BuildString, 8))
        .Extra = BuildString
        .Speed = "WAIT"
        .Active = True
    End With
    If Not SkipMessage Then Call AddMessage(Player(Number).Name & " signed on.")
End Sub

Sub DoChallenge(ByVal Number As Integer)
    If ChallengeLoaded Then
        Call SendData("PBSY:" & Number)
    Else
        ChallengePending = True
        ChallengeNumber = Number
        ICalled = False
        On Error Resume Next
        Unload ChallengeWindow
        'If DebugMode Then MsgBox "Debug!"
        ChallengeWindow.Show
    End If
End Sub

Public Sub SendData(ByVal SendMe As String)
    If Socket.State <> sckConnected Then Exit Sub
    '>>> Call WriteDebugLog("Sent: " & SendMe)
    Call AddMessage("Sent: " & SendMe, True, ":", vbGreen, True)
    Socket.SendData FormatPacket(SendMe, UseXOR)
    If Left(SendMe, 5) = "CHAT:" Then FloodCheck = FloodCheck + 1
    Call ChangeFCBar(FloodCheck)
    Call WriteDebugLog("MasterServer.SendData: " & SendMe)
End Sub

Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal Underline As Boolean = False)
    If DebugMessage And Not DebugMode Then Exit Sub
    Call RTB.AddMessage(Message, BreakChar, Color, Bold, Italic, False, Underline)
End Sub

Sub ChangeFCBar(ByVal Value As Integer)
    If FloodTolerance = 0 Then Exit Sub
    If Value > FCBar.Max Then
        FCBar.Value = FCBar.Max
    Else
        FCBar.Value = Value
    End If
    If FCBar.Value < FCBar.Max / 2 Then FCBar.FillColor = vbGreen
    If FCBar.Value >= FCBar.Max / 2 And FCBar.Value < FCBar.Max - (FCBar.Max / 4) Then FCBar.FillColor = vbYellow
    If FCBar.Value >= FCBar.Max - (FCBar.Max / 4) Then FCBar.FillColor = vbRed
    If FloodTolerance = 0 Then Exit Sub
    If FloodCheck >= FloodTolerance And ServerIssue <> 7 Then
        Command1.Enabled = False
        ServerIssue = 7
        Call SendData("EXIT:")
        ExitThis = True
    End If
End Sub

Public Sub mnuTeamItem_Click(Index As Integer)
    Dim X As Integer
    If Not TeamChangeFromMS Then
        TeamChanged = False
        mnuTeam.Enabled = False
        TeamChangeFromMS = True
        NowSwitching = True
        OldName = You.Name
        Select Case Index
            Case 0
                Rearrange.Show vbModal
            Case 1
                BoxArrange.Show vbModal
            Case 2
                ItemChange.Show vbModal
            Case 3
                'Me.Enabled = False
                TeamBuilder.Show
                Exit Sub
            Case 4
                Call TeamLoader.OpenTheFile
        End Select
    End If
    TeamChangeFromMS = False
    If TeamChanged Then
        For X = 1 To 6
            StoredPKMN(X) = PKMN(X)
        Next X
        Select Case Index
        Case Is < 3
            Call SendData("TUPD:" & PrepareTeamInfo)
        Case 3
            If TBUserChange Then
                Call SendData("RKEY:")
                mnuOptionsItem(7).Checked = False
            Else
                Call SendData("TUPD:" & PrepareTeamInfo)
            End If
        Case 4
            Call SendData("RKEY:")
            mnuOptionsItem(7).Checked = False
        End Select
        StatusBar1.Panels(1).Text = "Team Changed"
    Else
        NowSwitching = False
        mnuTeam.Enabled = True
    End If
End Sub

Private Sub picResizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SizeX = X
        Sizing = True
    End If
End Sub

Private Sub picResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Single
    Dim Z As Single
    If Button = 1 And Sizing Then
        Z = (X - SizeX)
        A = picResizer.Left + Z
        If A < 30 Then
            Z = 30 - picResizer.Left
        End If
        If A > Command2.Left - 1700 Then
            Z = Command2.Left - 1700 - picResizer.Left
        End If
        If Z = 0 Then Exit Sub
        SetRedraw Me.hWnd, False
        Call UserListResize(Z)
        SetRedraw Me.hWnd, True
    End If
End Sub
Private Sub UserListResize(Z As Single)
    picResizer.Left = picResizer.Left + Z
    UserList.Width = UserList.Width + Z
    UserList.ColumnHeaders(1).Width = UserList.ColumnHeaders(1).Width + Z
    Messages.Left = Messages.Left + Z
    Messages.Width = Messages.Width - Z
    ChatBox.Left = ChatBox.Left + Z
    ChatBox.Width = ChatBox.Width - Z
    Label1.Left = Label1.Left + Z
    FCBar.Left = FCBar.Left + Z
End Sub
Private Sub picResizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sizing = False
End Sub

Private Sub Socket_Close()
    Dim X As Integer
    If WasDisconnected Then Exit Sub
    MasterServer.Caption = "Stadium: Disconnected"
    WasDisconnected = True
    Call AddMessage("You have been disconnected from the server.", , , vbRed, True, True)
    If Battling Then Call Battle.AddMessage("You have been disconnected from the server.", , , vbRed, True, True)
    For X = 1 To 5
        If WatchLoaded(X) Then Call WatchForm(X).AddMessage("You have been disconnected from the server.", , , vbRed, True, True)
    Next X
    If ExitThis = True Then Unload Me
End Sub

Private Sub Socket_Connect()
    Connector.Enabled = False
    Command1.Enabled = True
    MasterServer.Caption = "Stadium: " & ServerRegName
End Sub

Private Sub Socket_DataArrival(ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(Socket, BytesTotal, Packet)
    If Worked Then
        StatusBar1.Panels(3).Text = BytesTotal & " bytes rcd."
        For X = 1 To UBound(Packet)
            Call DoIncoming(Packet(X))
        Next X
    Else
        Call AddMessage("Data Arrival error.  Disconnecting...", , , vbRed, True)
        Call SendData("EXIT:")
    End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    StatusBar1.Panels(1).Text = "Network Error: " & Description
    If Socket.State <> sckConnected Then Call Socket_Close
    'Connector.Enabled = True
End Sub

Function PrepareUserInfo(LogOnKey As String, Optional ByVal PSWD As String = "") As String
    Dim Build As String
    Dim Temp As String
    Dim X As Integer
    If PSWD <> "" Then Build = Pad(PSWD, 16) Else Build = ""


    Build = Build & Pad(You.Name, 20) & Pad(You.ProgVersion, 7) & StationID
    PasswordBoxTitle = "User Password"
    PasswordBoxCaption = "Your user password is not saved." & vbNewLine & "Please enter the password."
    If Len(SavedPassword) = 0 Then
        PWWindow.Show 1
        If Len(ServerPassword) <> 0 Then ServerPassword = MD5(ServerPassword)
    Else
        ServerPassword = SavedPassword
    End If
    If Len(ServerPassword) = 0 Then
        Build = Build & String$(15, vbNullChar)
    Else
        Temp = ServerPassword
        For X = 1 To 32 Step 2
            Build = Build & Chr$(Val("&H" & Mid(Temp, X, 2)))
        Next X
    End If
    Temp = String(10, "0")
    For X = 1 To 10
        If MsgToggle(X) Then Mid$(Temp, X, 1) = "1"
    Next X
    Temp = Bin2Chr(Temp)
    Build = Build & Temp & FixedHex(You.Picture, 2) & Bool2Bin(AllowViewing) & Chr$(Bin2Dec(Bool2Bin(BadSID) & Dec2Bin(Val(Ranking), 7))) & FixedHex(MakeBinArray(Compatibility), 2) & LogOnKey & You.Extra
    PrepareUserInfo = Build
End Function

Public Function PrepareTeamInfo() As String
    Dim Build As String
    Dim X As Integer
    Dim Y As Integer
    Dim Z As String
'    Call ReadBinArray(CompatCheck, Compatibility)
'    Build = FixedHex(CompatCheck, 2)
    Build = Build & Bool2Bin(AllowViewing)
    Build = Build & Chr$(Ranking)
    Build = Build & FixedHex(You.Version, 1)
    For X = 1 To 6
        Build = Build & PKMN2Str(PKMN(X))
    Next X
    PrepareTeamInfo = Build
End Function

Sub RefreshAuth()
    Select Case Val(Player(YourNumber).Authority)
        Case 0, 1
            mnuPlayerItem(3).Visible = False
            mnuPlayerItem(4).Visible = False
            mnuPlayerItem(5).Visible = False
        Case 2, 3
            mnuPlayerItem(3).Visible = True
            mnuPlayerItem(4).Visible = True
            mnuPlayerItem(5).Visible = True
'        Case 3
'            mnuPlayerItem(3).Visible = True
'            mnuPlayerItem(4).Visible = True
'            mnuPlayerItem(5).Visible = True
    End Select
End Sub

Private Sub DoInitialResize()
    Dim Maximized As Boolean
    Dim Width As Long
    Dim Height As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Maximized = GetSetting("NetBattle", "Server Window", "Maximized", False)
    If Maximized Then Me.WindowState = vbMaximized: Exit Sub
    Width = GetSetting("NetBattle", "Server Window", "Width", 6735)
    Height = GetSetting("NetBattle", "Server Window", "Height", 5040)
    If Width < MinWidth Then Width = MinWidth
    If Height < MinHeight Then Height = MinHeight
    Me.Width = Width
    Me.Height = Height
    Call CenterWindow(Me)
End Sub

