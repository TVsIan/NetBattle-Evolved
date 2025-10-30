VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerWindow 
   Caption         =   "NetBattle Master Server"
   ClientHeight    =   4050
   ClientLeft      =   2850
   ClientTop       =   5475
   ClientWidth     =   6225
   Icon            =   "ServerWindow.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   6225
   Begin VB.Timer ScriptTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1200
   End
   Begin VB.Timer RndTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   1200
   End
   Begin VB.Timer SaveDataTimer 
      Interval        =   60000
      Left            =   5640
      Top             =   720
   End
   Begin VB.Timer RegTimer 
      Interval        =   15000
      Left            =   5640
      Top             =   240
   End
   Begin VB.Timer UserChangeTimer 
      Interval        =   5000
      Left            =   5640
      Top             =   1200
   End
   Begin VB.Timer tmrKickTimer 
      Interval        =   1000
      Left            =   4680
      Top             =   1200
   End
   Begin VB.Timer IdleTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   5160
      Top             =   240
   End
   Begin MSWinsockLib.Winsock ClientSocket 
      Index           =   0
      Left            =   4200
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   30000
   End
   Begin MSWinsockLib.Winsock RegSocket 
      Left            =   4680
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer SendAllQueue 
      Interval        =   10
      Left            =   4680
      Top             =   720
   End
   Begin VB.Timer MissingDataTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   4200
      Top             =   720
   End
   Begin VB.Timer ChannelScanner 
      Interval        =   60000
      Left            =   5160
      Top             =   720
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   2865
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   5054
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ServerWindow.frx":1272
   End
   Begin VB.TextBox ChatBox 
      Height          =   285
      Left            =   1980
      MaxLength       =   200
      TabIndex        =   1
      Top             =   3000
      Width           =   4170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3795
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2778
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6376
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Trainers"
         Object.Width           =   2672
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   635
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save Log..."
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Start &Logging..."
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuServerItem 
         Caption         =   "&Options..."
         Index           =   1
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "&Server Data..."
         Index           =   2
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "Soc&ket Status..."
         Index           =   3
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "Scrip&t Window..."
         Index           =   4
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "Data&base Changes..."
         Index           =   5
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "&Relay Spy"
         Index           =   7
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "Raw &Data"
         Index           =   8
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "Confirm E&xit"
         Checked         =   -1  'True
         Index           =   9
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "&Clear Messages"
         Index           =   11
      End
      Begin VB.Menu mnuServerItem 
         Caption         =   "R&eload Database Files"
         Index           =   12
      End
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "&Player"
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Info..."
         Index           =   0
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Kick"
         Index           =   1
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Ban (IP)"
         Index           =   2
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "Ban (&Name/ID)"
         Index           =   3
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPlayerItem 
         Caption         =   "&Refresh"
         Index           =   5
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
Attribute VB_Name = "ServerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type SendAllType
    Data As String
    Packet As String
End Type
Private Type TempBanType
    SID As String
    PName As String
    TimeLeft As Integer
End Type
Private Type DNSQueueType
    PNum As Long
    IP As Long
End Type
Private RTB As RTBClass
Private Exiting As Boolean
Private RawData As Boolean
Private Logging As Boolean
Private DataBuffer(256) As String
Private SentPing(256) As Boolean
'Private BattleOver(256) As Boolean
Private SendAllData(256, 500) As SendAllType
Private ReceiveQueue(256, 500) As String
Private ServerBattle() As BattleData
Private BattleReady(256) As Boolean
Private ConfirmExit As Boolean
Private BlankPlayer As MSPlayer
Private ClearLoops As Byte
Private SinceLastSave As Long
Private TodaysDate As Date
Private LogDate As Date
Private BlankSend As SendAllType
Private TempBanInfo() As TempBanType
Private UserCount As Integer
Private ResizeOK As Boolean
Private ServerCode As String
Public ShuttingDown As Boolean
Public SNRegged As Boolean
Public DNSPtr As Long
Private DNSQueue() As DNSQueueType
Private Const MinWidth = 5900
Private Const MinHeight = 2800


Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3

Private Sub BWMon_Timer()

End Sub

Public Sub RemoveTempban(SID As String)
    Dim X As Integer
    For X = 1 To UBound(TempBanInfo)
        If TempBanInfo(X).SID = SID Then
            TempBanInfo(X).TimeLeft = 0
            TempBanInfo(X).PName = ""
            TempBanInfo(X).SID = ""
        End If
    Next X
End Sub

Function SendFullList(ByVal Index As Long) As String
    Dim X As Long
    Dim Temp As String
    Dim Temp2 As String
    Dim Build As String
    Dim Packet() As String
    For X = 1 To MaxUsers
        If Player(X).Name <> "" Then
            Temp = PreparePlayerData(X)
            Temp2 = ChopString(Temp, 5)
            Temp = Chr$(Len(Temp)) & Temp
            Build = Build & Temp
        End If
    Next X
    X = Len(Build)
    MainContainer.Compressor.CompressString Build
    Build = "  " & Bin2Chr(Dec2Bin(X, 16)) & Build
    ReDim Packet(1)
    While Len(Build) > 200
        Packet(UBound(Packet)) = "/WHO:" & ChopString(Build, 200) & vbNullChar
        ReDim Preserve Packet(UBound(Packet) + 1)
    Wend
    Packet(UBound(Packet)) = "/WHO:" & Build
    Mid(Packet(1), 6, 2) = Bin2Chr(Dec2Bin(UBound(Packet), 16))
    For X = 1 To UBound(Packet)
        Call AddToQueue(Index, Packet(X))
    Next X
End Function

Private Sub ChatBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer <> 0 Then Me.MousePointer = 0
End Sub


Private Sub ClientSocket_Close(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        ClientSocket(0).Listen
    Else
        If Not Disconnecting(Index) Then
            If Player(Index).DCReason = "" Then Player(Index).DCReason = "Socket Closed."
            Call DisconnectPlayer(Index)
        End If
    End If
End Sub

Private Sub ClientSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim ThisUser As Long
    Dim X As Long
    Dim Y As Integer
    Dim ActivePlayers As Integer
    Dim Temp As String
    Dim TooManyIPs As Boolean
    Dim IsBanned As Boolean
    On Error Resume Next
    If MainContainer.ServerKiller.Enabled Then Exit Sub
    StatusBar1.Panels(1).Text = "Incoming connection request received."
    Call AddMessage("Connect request received")
    'Debug.Print ClientSocket(0).RemoteHost
    ThisUser = MaxUsers + 1
    For X = 1 To MaxUsers
        If Not IsLoaded(X) Then
            IsLoaded(X) = True
            ThisUser = X
            Exit For
        End If
    Next
    If ThisUser <= MaxUsers Then
        With Player(ThisUser)
            For X = 1 To 500
                SendAllData(ThisUser, X) = BlankSend
                ReceiveQueue(ThisUser, X) = ""
            Next
            SentPing(ThisUser) = False
            Chances(ThisUser) = 0
            Unload ClientSocket(ThisUser)
            Load ClientSocket(ThisUser)
            Load MissingDataTimer(ThisUser)
            ClientSocket(ThisUser).Close
            ClientSocket(ThisUser).Tag = ""
            ClientSocket(ThisUser).Accept requestID
            .Address = ClientSocket(ThisUser).RemoteHostIP
                    
            .CommandLock = "USER:"
            ReDim .Ignore(0)
            ActivePlayers = ListView1.ListItems.count + 1
            For X = 1 To UBound(PANum)
                PANum(X).Value(ThisUser) = 0
            Next X
            For X = 1 To UBound(PATxt)
                PATxt(X).Value(ThisUser) = """"""
            Next X
            Call AddMessage("Request connected on slot " & ThisUser & " from IP " & .Address)
            
            Temp = ""
            
            'Even with the encryption, it would still be possible
            'to copy the encrypted packet of an incoming user, then
            'send the same packet to another server in place of your
            'own sign on information.  The ServerCode thing kills
            'that idea.
            
            Y = 0
            For X = 1 To MaxUsers
               If .Address = Player(X).Address Then Y = Y + 1
            Next X
            TooManyIPs = (Y > MaxIPs) And (MaxIPs <> 0)
            
            'If the player is IP or ISP banned, put them in LockDown.
            'Nothing gets in or out other than a single BANU: packet.
            Temp = ""
            IsBanned = IPIsBanned(.Address, Temp)
            If Not (IsBanned Or TooManyIPs) Then
                If Left$(.Address, 7) <> "168.192" Then
                    .DNSAddress = ClientSocket(ThisUser).RemoteHost
'                    .DNSAddress = "[DNS Pending]"
'
'                    ReDim Preserve DNSQueue(UBound(DNSQueue) + 1)
'                    DNSQueue(UBound(DNSQueue)).IP = inet_addr(.Address)
'                    DNSQueue(UBound(DNSQueue)).PNum = ThisUser
'                    If UBound(DNSQueue) = 1 Then
'                        SetDNSIP DNSQueue(1).IP, DNSQueue(1).PNum
'                        Call CreateThread(ByVal 0&, ByVal 0&, DNSPtr, ByVal 0&, 0, ByVal 0)
'                    End If
'                    X = 0
'                    Do
'                        X = X + 1
'                        DoEvents
'                        Sleep 1
'                    Loop Until X > 2000 Or .DNSAddress <> "[DNS Pending]" Or Not IsLoaded(ThisUser)
'                    If Not IsLoaded(ThisUser) Then Exit Sub
'                    If X <= 200 Then
'                        Debug.Print .DNSAddress
'                        IsBanned = ISPIsBanned(.DNSAddress, Temp)
'                    Else
'                        Debug.Print "Timeout"
'                    End If
                End If
                
            End If
            If IsBanned Then
                Call AddMessage("Banned user attemping to log on.  Disconnecting.")
                .LockDown = True
                If Temp = "" Then Temp = DefaultBanMsg
                Call AddToQueue(ThisUser, "BANU:" & Temp)
            ElseIf TooManyIPs Then
                Call AddToQueue(ThisUser, "NOIP:" & CStr(MaxIPs))
            ElseIf ServerPassword = "" Then
                Call AddToQueue(ThisUser, "REQN:" & CStr(Abs(UseXOR)) & Pad(You.ProgVersion, 10) & FixedHex(MaxUsers, 3) & FixedHex(FloodTolerance, 2) & FixedHex(ActivePlayers, 3) & Pad(ServerName, 20) & ServerCode)
            Else
                Call AddToQueue(ThisUser, "RPWD:" & CStr(Abs(UseXOR)) & Pad(You.ProgVersion, 10) & FixedHex(MaxUsers, 3) & FixedHex(FloodTolerance, 2) & FixedHex(ActivePlayers, 3) & Pad(ServerName, 20) & ServerCode)
            End If
        End With
    Else
        ClientSocket(0).Accept requestID
        Call AddToQueue(0, "BUSY:")
    End If
End Sub

Private Sub ClientSocket_DataArrival(Index As Integer, ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(ClientSocket(Index), BytesTotal, Packet)
    If Worked Then
        If Player(Index).LockDown Then
            If Player(Index).DCReason = "" Then Player(Index).DCReason = "Player is banned."
            Call DisconnectPlayer(Index)
            Exit Sub
        End If
        IsLoaded(Index) = True
        StatusBar1.Panels(2).Text = BytesTotal & " bytes rcd."
        Chances(Index) = 0
        SentPing(Index) = False
        For X = 1 To UBound(Packet)
            Call DoIncoming(Index, Packet(X))
        Next X
    Else
        Player(Index).DCReason = "Data Arrival error."
        Call DisconnectPlayer(Index)
    End If
End Sub

Private Sub ClientSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    StatusBar1.Panels(1).Text = "Error: " & Description & " on channel " & Index
    On Error Resume Next
    If Index = 0 Then
        ClientSocket(0).Close
        ClientSocket(0).Listen
    Else
        If IsLoaded(Index) Then
            If Player(Index).DCReason = "" Then Player(Index).DCReason = "Socket Error: " & Description
            Call DisconnectPlayer(Index)
        End If
    End If
End Sub

 
'Private Sub Command2_Click()
'    Call ScriptTest
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim X1 As Single
    Dim X2 As Single
    Dim X3 As Single
    Dim NewX1 As Single
    Dim NewX2 As Single
    Dim NewW1 As Single
    Dim NewW2 As Single
    Dim NewW3 As Single
    Dim MoveOK As Boolean
    X1 = ListView1.Width + ListView1.Left
    X2 = Messages.Left
    X3 = Abs(X1 - X2) / 2
    If X > X1 - 100 And X < X2 + 100 Then
        ResizeOK = True 'If Me.MousePointer <> 9 Then Me.MousePointer = 9
    Else
        If Button <> 1 Then ResizeOK = False
    End If
    If Button = 1 And ResizeOK Then 'Me.MousePointer = 9 Then
        NewX1 = X - X3
        NewX2 = X + X3
        NewW1 = ListView1.Width + NewX1 - X1
        NewW2 = Messages.Width - NewX2 + X2
        If NewW2 < 1500 Then
            NewX2 = Messages.Width + X2 - 1500
            NewW2 = Messages.Width - NewX2 + X2
            NewX1 = NewX2 - X3 - X3
            NewW1 = ListView1.Width + NewX1 - X1
        End If
        If NewW1 < 1000 Then
            NewX1 = X1 - ListView1.Width + 1000
            NewW1 = ListView1.Width + NewX1 - X1
            NewX2 = NewX1 + X3 + X3
            NewW2 = Messages.Width - NewX2 + X2
        End If
        NewW3 = ListView1.ColumnHeaders(1).Width + NewX1 - X1
        If NewW3 < 300 Then NewW3 = 300
        Messages.Left = NewX2
        ChatBox.Left = NewX2
        Messages.Width = NewW2
        ChatBox.Width = Messages.Width
        If NewW3 < ListView1.ColumnHeaders(1).Width Then
            ListView1.ColumnHeaders(1).Width = NewW3
        End If
        ListView1.Width = NewW1
        If NewW3 > ListView1.ColumnHeaders(1).Width Then
            ListView1.ColumnHeaders(1).Width = NewW3
        End If
        ListView1.Refresh
        Messages.Refresh
        ChatBox.Refresh
        Me.Refresh
        DoEvents
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, Shift, X, Y)
End Sub


Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If ListView1.Width - X < 100 Then
'        Call Form_MouseMove(Button, Shift, ListView1.left + X, Y)
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
    End If
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
  
Public Sub VScrollTextBox(ByRef TBox As TextBox, ByVal ScrollDown As Boolean, ByVal PageMode As Boolean)
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

Private Sub ChannelScanner_Timer()
    Dim X As Integer
    
    'ChannelScanner Timer
    On Error Resume Next
    For X = 1 To MaxUsers
        If IsLoaded(X) Then
            If SentPing(X) Then
                If Chances(X) = 3 Then
                    Player(X).DCReason = "PONG Timeout"
                    'Call AddMessage("Disconnecting player: " & Player(X).Name & " due to inactivity.")
                    Call DisconnectPlayer(X)
                ElseIf Chances(X) < 3 Then
                    Chances(X) = Chances(X) + 1
                    Call AddToQueue(X, "PING:")
                End If
            Else
                Call AddToQueue(X, "PING:")
                SentPing(X) = True
                Chances(X) = 1
            End If
        End If
    Next
    If RegSocket.State <> sckConnected And PublicServer Then
        RegSocket.Close
        RegSocket.Connect
        RegTimer.Enabled = True
    End If
    
    'Tempban Timer
    For X = 1 To UBound(TempBanInfo)
        With TempBanInfo(X)
            If .TimeLeft <> 0 Then
                .TimeLeft = .TimeLeft - 1
                If .TimeLeft = 0 Then
                    Call ServerDB.DelSIDBan(.PName)
                    .PName = ""
                End If
            End If
        End With
    Next X
    For X = UBound(TempBanInfo) To 1 Step -1
        If TempBanInfo(X).TimeLeft <> 0 Then Exit For
    Next X
    If X <> UBound(TempBanInfo) Then ReDim Preserve TempBanInfo(X)
End Sub

Private Sub Command1_Click()
    Dim Temp As String
    Dim X As Integer
    Dim E As String
    On Error Resume Next
    If ChatBox.Text = "" Then Exit Sub
    Command1.Enabled = False
    If Left(ChatBox.Text, 1) = "/" Then
        Temp = LineCheck(ChatBox.Text, E)
        ChatBox.Text = ""
        If E = "" Then E = ScriptMod.Exec(Temp, BlankSource, False)
        If E <> "" Then AddMessage ("ERROR: " & E)
    Else
        Temp = "CMSG:SERVER MESSAGE: " & ChatBox.Text
        ChatBox.Text = ""
        Call AddMessage(Right(Temp, Len(Temp) - 5))
        Call SendAll(Temp)
    End If
    Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim TempVar As String
    
    On Error Resume Next
    
    Set RTB = New RTBClass
    RTB.SetRTBHook Messages, ChatBox, MinWidth, MinHeight
    'RTB.UseTimestamp = True
    RTB.LimitText = True
    
    Call CenterWindow(ServerWindow)
    ListView1.Icons = MainContainer.Trainers
    ListView1.SmallIcons = MainContainer.MiniTrainers
    MaxUsers = GetSetting("NetBattle", "Master Server", "Max Users", 100)
    'Set up array sizes
    ReDim Player(MaxUsers)
    For X = 1 To MaxUsers
        ReDim Player(X).Ignore(0)
    Next
    ReDim DNSQueue(0)
    ReDim Chances(MaxUsers)
    ReDim Disconnecting(MaxUsers)
    ReDim ServerBattle(Int(MaxUsers / 2))
    ReDim TempBanInfo(0)
    ReDim BlankPlayer.Ignore(0)
    For X = 1 To 5
        ServerCode = ServerCode & Chr$(Int(Rnd * 256))
    Next X
    DNSPtr = GetDNSPtr(Me.hWnd)
    NB_DNSADDR = RegisterWindowMessage("NB_DNSADDR")
    ShuttingDown = False
    SNRegged = False
    ServerMessage = GetSetting("NetBattle", "Master Server", "Server Message", "")
    AllowNewUsers = GetSetting("NetBattle", "Master Server", "Allow New Users", 1)
    AllowOldVersions = GetSetting("NetBattle", "Master Server", "Allow Old Versions", 1)
    ServerPassword = GetSetting("NetBattle", "Master Server", "Server Password", "")
    If Len(ServerPassword) > 16 Then ServerPassword = Left(ServerPassword, 16)
    FloodTolerance = GetSetting("NetBattle", "Master Server", "Floodcount Tolerance", 10)
    ListenWrong = GetSetting("NetBattle", "Master Server", "16150 Listen", False)
    SendTimer = GetSetting("NetBattle", "Master Server", "Timer Interval", 10)
    ConfirmExit = GetSetting("NetBattle", "Master Server", "Confirm Exit", True)
    NumLines = GetSetting("NetBattle", "Master Server", "Display Lines", 0)
    NumLines = Int(NumLines / 100) * 100
    If NumLines > 10000 Then NumLines = 10000
    If NumLines < 0 Then NumLines = 0
    RTB.LineLimit = NumLines
    If NumLines = 0 Then RTB.LimitText = False
    PurgeDays = GetSetting("NetBattle", "Master Server", "Purge Days", 90)
    If PurgeDays > 360 Then PurgeDays = 360
    If PurgeDays < 30 Then PurgeDays = 30
    TodaysDate = Date
    UseXOR = GetSetting("NetBattle", "Master Server", "Encrypt", True)
    MaxIPs = GetSetting("NetBattle", "Master Server", "MaxIPs", 0)
    If MaxIPs > 16 Then MaxIPs = 16
    DefaultBanMsg = GetSetting("NetBattle", "Master Server", "Ban Message", "")
    If GetSetting("NetBattle", "Master Server", "AutoLogging", 0) = 1 Then Call StartAutoLog
    DBModStr = Bin2Chr(GetSetting("NetBattle", "Server", "dbmod", vbNullString))
    
    mnuServerItem(5).Checked = ConfirmExit
    BMessStyle = 1
    If SendTimer > 0 Then
        SendAllQueue.Interval = SendTimer
    Else
        SendAllQueue.Enabled = False
    End If
    For X = 1 To UBound(ServerBattle)
        Set ServerBattle(X) = New BattleData
    Next
    Exiting = False
    ServerRunning = True
    Me.Height = Int(MainContainer.Height * 0.8)
    Me.Width = Int(MainContainer.Width * 0.9)
    CenterWindow Me
    Me.Show
    Call AddMessage("Server Version " & You.ProgVersion)
    Call AddMessage("Loading Pokemon Database...")
    If BasePKMN(1).No <> 1 Then PokeLoader.Show
    Call AddMessage("Pokemon Database successfully opened")
    Call ApplyDBMod
    Call AddMessage("Initializing player database...")
    Call ServerDB.InitDB
    Call ServerDB.PurgeUsers(PurgeDays)
    Call AddMessage("All data ready!")
    Call ScriptMod.ScriptInit
    Logging = False
    RunningServer = True
    TaskBarIcon.Show
    TaskBarIcon.Visible = False
    PublicServer = GetSetting("NetBattle", "Master Server", "Public", True)
    RealIP = GetSetting("NetBattle", "Master Server", "Real IP", "")
    ServerName = Left$(GetSetting("NetBattle", "Master Server", "Name", ""), 20)
    Admin = Left$(GetSetting("NetBattle", "Master Server", "Admin", ""), 20)
    ServerDesc = GetSetting("NetBattle", "Master Server", "Desc", "")
    UseTrueRnd = GetSetting("NetBattle", "Master Server", "UseTrueRnd", False)
    RndState = rEmpty
    RndCache = 0
    BitPos = 0
    TempVar = GetSetting("NetBattle", "Master Server", "RndGroup", 1024)
    If Not IsNumeric(TempVar) Then
        RndGroup = 1024
        SaveSetting "NetBattle", "Master Server", "RndGroup", 1024
    ElseIf CLng(TempVar) < 256 Or CLng(TempVar) > 16384 Then
        RndGroup = 1024
        SaveSetting "NetBattle", "Master Server", "RndGroup", 1024
    Else
        RndGroup = CLng(TempVar)
    End If
    TempVar = GetSetting("NetBattle", "Master Server", "RndThresh", Int(RndGroup / 8))
    If Not IsNumeric(TempVar) Then
        RndThresh = Int(RndGroup / 8)
        SaveSetting "NetBattle", "Master Server", "RndThresh", RndThresh
    ElseIf CLng(TempVar) < 8 Or CLng(TempVar) > Int(RndGroup / 8) Then
        RndThresh = Int(RndGroup / 8)
        SaveSetting "NetBattle", "Master Server", "RndThresh", RndThresh
    Else
        RndThresh = CLng(TempVar)
    End If
    RndTimer.Enabled = UseTrueRnd
    RegSocket.RemotePort = RegPort
    RegSocket.RemoteHost = RegAddress
    If Not PublicServer Then
        Call ServerStartup
        Exit Sub
    End If
    If ServerName = "" Or Admin = "" Or ServerDesc = "" Then
        ServerName = InputBox("New to version 0.9.0 is the NetBattle Server Registry, which keeps track of all online servers so that people can find them easily.  Please enter a name for your server.  Note that this is NOT an IP Address or a Redirect; these things are stored internally.  The name is just something for your server to be known as.", "Registry Options", ServerName)
        If ServerName = "" Then
            Call AddMessage("Did not supply Server Name; cannot connect to registry.")
            Call AddMessage("Please set your Server Name in the options screen.")
            Call ServerStartup
            Exit Sub
        Else
            SaveSetting "NetBattle", "Master Server", "Name", ServerName
        End If
        Admin = InputBox("Please enter the NetBattle Trainer Name that you most commonly use while battling.", "Registry Options", Admin)
        If Admin = "" Then
            Call AddMessage("Did not supply Admin name; cannot connect to registry.")
            Call AddMessage("Please set your Admin name in the options screen.")
            PublicServer = False
            Call ServerStartup
            Exit Sub
        Else
            SaveSetting "NetBattle", "Master Server", "Admin", Admin
        End If
        ServerDesc = Left(InputBox("Please enter a brief description of your server.  (190 character limit)", "Registry Options", ServerDesc), 190)
        If ServerDesc = "" Then
            Call AddMessage("Did not supply Description; cannot connect to registry.")
            Call AddMessage("Please set your Description in the options screen.")
            PublicServer = False
            Call ServerStartup
            Exit Sub
        Else
            SaveSetting "NetBattle", "Master Server", "Desc", ServerDesc
        End If
    End If
    Call AddMessage("Attempting to connect to server registry at " & RegAddress)
    RegTimer.Enabled = True
    RegSocket.Connect
End Sub

Private Sub ServerStartup()
    Call AddMessage("Server Startup")
    ScriptTimer.Enabled = True
    Call BlockExec(23)
    ClientSocket(0).LocalPort = MainPort
    ClientSocket(0).Listen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Call ExitProgram(Cancel)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If ServerWindow.WindowState <> vbMinimized Then
        If ServerWindow.Width < 3000 Then ServerWindow.Width = 3000
        If ServerWindow.Height < 2000 Then ServerWindow.Height = 2000
        Command1.Top = ServerWindow.Height - 1260
        Command1.Left = ServerWindow.Width - 1260
        Messages.Width = ServerWindow.Width - Messages.Left - 200
        ChatBox.Width = Messages.Width
        Messages.Height = ServerWindow.Height - 1845
        ChatBox.Top = ServerWindow.Height - 1620
        ListView1.Height = ServerWindow.Height - 1005
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Logging Then Close #LogFileNum
    Call ServerDB.WriteDB
    Call ScriptMod.CleanUpScript
End Sub

Private Sub ListView1_DblClick()
    Call mnuPlayerItem_Click(0)
 End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuPlayer
    End If
End Sub

Private Sub Messages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer <> 0 Then Me.MousePointer = 0
End Sub

Private Sub MissingDataTimer_Timer(Index As Integer)
'    Dim Temp As String
'    Dim X As Integer
'
'    On Error Resume Next
'    MissingDataTimer(Index).Enabled = False
'    Temp = DataBuffer(Index)
'    DataBuffer(Index) = ""
'    If Len(Temp) <= NetChunkSize Then
'        Call DoIncoming(Index, Temp)
'    Else
'        While Len(Temp) > 256
'            Call DoIncoming(Index, left(Temp, 256))
'            Temp = Right(Temp, Len(Temp) - 256)
'        Wend
'        Call DoIncoming(Index, Temp)
'    End If
'
'
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileToUse As String
    Dim YourChoice As Integer
    Dim FileNum As Integer
        
    FileNum = FreeFile
    On Error Resume Next
    Select Case Index
        Case 0
            With MainContainer.FileBox
                .DialogTitle = "Save Server Log"
                .Flags = cdlOFNOverwritePrompt
                .CancelError = True
                .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                .DefaultExt = ".txt"
                .FileName = ""
                .InitDir = SlashPath
                On Error GoTo Cancelled
                .ShowSave
                FileToUse = .FileName
            End With
            Open FileToUse For Output As #FileNum
            Print #FileNum, Messages.Text
            Close #FileNum
            StatusBar1.Panels(1).Text = "Server log saved"
Cancelled:
        Case 1
            If Logging = False Then
                With MainContainer.FileBox
                    .DialogTitle = "Save Server Log"
                    .Flags = cdlOFNOverwritePrompt
                    .CancelError = True
                    .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                    .DefaultExt = ".txt"
                    .FileName = ""
                    .InitDir = SlashPath
                    On Error GoTo Cancelled
                    .ShowSave
                    FileToUse = .FileName
                End With
                Open FileToUse For Output As #FileNum
                Print #FileNum, Messages.Text
                StatusBar1.Panels(1).Text = "Server logging enabled"
                Logging = True
            Else
                Close #FileNum
                Logging = False
                StatusBar1.Panels(1).Text = "Server logging ended"
            End If
        Case 3
            Call ExitProgram(YourChoice)
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

Private Sub mnuPlayerItem_Click(Index As Integer)
    Dim Temp As String
    Dim Temp2 As Integer
    Dim X As Integer
    Dim Cancel As Boolean
    On Error GoTo NoSel
    Select Case Index
        Case 0
            Temp = ListView1.SelectedItem.Key
            Temp2 = Val(Right(Temp, Len(Temp) - 5))
            ChallengeNumber = Temp2
            If PInfo.Visible Then
                MsgBox "You can only view one player at a time!", , "Error"
                Exit Sub
            End If
            On Error Resume Next
            Unload PInfo
            PInfo.Show
        Case 1
            Temp = ListView1.SelectedItem.Key
            Temp2 = Val(Right(Temp, Len(Temp) - 5))
            Cancel = False
            Call ScriptMod.BlockExec(15, Cancel, 0, Temp2)
            If Cancel Then Exit Sub
            Call AddMessage(Player(Temp2).Name & " has been kicked.")
            Call SendAll("KICK:" & Temp2)
            Call AddToQueue(Temp2, "KCKU:")
            Call ScriptMod.BlockExec(16, 0, Temp2)
        Case 2
            Temp = ListView1.SelectedItem.Key
            Temp2 = Val(Right(Temp, Len(Temp) - 5))
            Cancel = False
            Call ScriptMod.BlockExec(17, Cancel, 0, Temp2)
            If Cancel Then Exit Sub
            Call AddMessage(Player(Temp2).Address & " has been banned.")
            Call BanUser(Temp2)
            Call ScriptMod.BlockExec(18, , 0, Temp2)
        Case 3
            Temp = ListView1.SelectedItem.Key
            Temp2 = Val(Right(Temp, Len(Temp) - 5))
            Cancel = False
            Call ScriptMod.BlockExec(17, Cancel, 0, Temp2)
            If Cancel Then Exit Sub
            Call AddMessage(Player(Temp2).Name & " has been banned.")
            Call SIDBanUser(Temp2)
            Call ScriptMod.BlockExec(18, , 0, Temp2)
        Case 5
            Call RefreshListing(True)
    End Select
NoSel:
End Sub

Private Sub mnuServerItem_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1 'Server Options
            Unload SetUsers
            SetUsers.Show vbModeless, MainContainer
            SetUsers.Slider1.Value = MaxUsers
        Case 2
            Unload UserEdit
            UserEdit.Show vbModeless, MainContainer
        Case 3
            Unload SocketStatus
            SocketStatus.Show
        Case 4
            Unload ScriptForm
            ScriptForm.Show
        Case 5
            Unload DatabaseMod
            DatabaseMod.Show
        Case 7
            DebugMode = Not DebugMode
            mnuServerItem(7).Checked = DebugMode
            If DebugMode Then
                StatusBar1.Panels(1).Text = "Relay Spy disabled"
            Else
                StatusBar1.Panels(1).Text = "Relay Spy enabled"
            End If
        Case 8
            If RawData Then
                RawData = False
                StatusBar1.Panels(1).Text = "Full messages disabled"
                mnuServerItem(1).Checked = False
            Else
                RawData = True
                StatusBar1.Panels(1).Text = "Full messages enabled"
                mnuServerItem(1).Checked = True
            End If
        Case 9
            mnuServerItem(5).Checked = Not mnuServerItem(5).Checked
            ConfirmExit = mnuServerItem(5).Checked
            SaveSetting "NetBattle", "Master Server", "Confirm Exit", ConfirmExit
        Case 11
            Messages.Text = ""
        Case 12
            Call ServerDB.InitDB
            AddMessage "Databases reloaded."
    End Select
End Sub

Sub ProcessIncoming(ByVal Index As Long, ByVal Temp As String)
    Dim Orig As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim TempString As String
    Dim RelayString As String
    Dim Temp2 As String
    Dim BreakPoint As Integer
    Dim PNum As Integer
    Dim SendString As String
    Dim SentPassword As String
    Dim ActivePlayers As Integer
    Dim BattleID As Integer
    Dim BlankBattle As BattleData
    Dim Command As String * 5
    Dim Data As String
    Dim GotoNext As Boolean
    Dim Cancel As Boolean
    Dim B As Boolean
    Dim Watch() As String
    Dim Packet() As String

    On Error Resume Next
    Orig = Temp
    'Debug.Print "ServerWindow.ProcessIncoming: " & Temp
    Data = Temp
    Command = ChopString(Data, 5)
    If Player(X).KickTimer <> 0 And Command <> "EXIT:" Then Exit Sub
    Call AddMessage("Received: " & Temp, , True)
    GotoNext = False
    Call WriteDebugLog("ServerWindow.ProcessIncoming: " & Command & Data & " (From Player " & X & ")")
    Select Case Command 'These commands can be received whenever and cannot be limited.
        Case "PONG:"
            SentPing(Index) = False
            Player(Index).PongCount = Player(Index).PongCount + 1
            If Player(Index).PongCount > 3 Then Call Tempban(Index)
        Case "CHAT:"
            TempString = ""
            If Player(Index).Name = "" Or UserIsBanned(Index) Or Player(Index).Active = False Then Exit Sub
            Temp = RTrim$(FilterIllegalChars(WordFilter(Data)))
            If Len(Temp) = 0 Then Exit Sub
            Temp2 = vbNullString
            If Left(Temp, 1) = "/" Then
                Temp2 = Temp
                B = SlashCommand(Index, Temp2)
            End If
            If B Then
                If Temp2 <> "" Then Temp = Temp2 Else Exit Sub
                TempString = Right(Temp2, Len(Temp2) - 5)
                Call ScriptMod.BlockExec(11, Cancel, Index, , TempString)
                If Cancel Then Exit Sub
            Else
                Cancel = False
                TempString = Temp
                Call ScriptMod.BlockExec(11, Cancel, Index, , Temp)
                If Cancel Then Exit Sub
                Temp2 = "CMSG:" & Player(Index).Name & ": " & Temp
            End If
            Call AddMessage(Right(Temp2, Len(Temp2) - 5))
            Call SendAll(Temp2, False, Index)
            If TempString <> "" Then Call ScriptMod.BlockExec(12, , Index, , TempString)
            Player(Index).FloodCount = Player(Index).FloodCount + 1
            If Player(Index).FloodCount > FloodTolerance + 4 Then
                Player(Index).DCReason = "Kicked for flooding."
                Call DisconnectPlayer(Index)
            End If
        Case "EXIT:"
'            Call ScriptMod.BlockExec(5, , Index)
'            If Not BattleOver(Index, True) Then
'                X = IIf(ServerBattle(Player(Index).BattleID).Unrated, 1, 0)
'                Call ScriptMod.BlockExec(7, , Player(Index).BattlingWith, Index, "FOR" & IIf(X = 1, "*", ""))
'                If X = 0 Then
'                    Call ChangePlayerStats(Index, 2)
'                    Call ChangePlayerStats(Player(Index).BattlingWith, 1)
'                End If
'                Call ScriptMod.BlockExec(8, , Player(Index).BattlingWith, Index, "FOR" & IIf(X = 1, "*", ""))
'            End If
            If Player(Index).DCReason = "" Then Player(Index).DCReason = "Normal exit."
            Call DisconnectPlayer(Index)
        Case "GETS:"
            X = Asc(Data)
            If X = 0 Or X > MaxUsers Then Exit Sub
            Call AddToQueue(Index, "PSPD:" & Data & Pad(Player(X).Speed, 4))
        Case "PWDS:", "NAME:"
            Player(Index).SkipXOR = True
            Call AddToQueue(Index, "CMSG:Due to security issues and different database formats, NetBattle no longer supports versions older than 0.9.0.  Please visit www.netbattle.net to download the latest version.")
            Player(Index).DCReason = "Version below 0.9.0"
            Call AddToQueue(Index, "BOOT:")
        Case Else
            GotoNext = True
    End Select
    If Not GotoNext Then Exit Sub
    
    If Player(Index).CommandLock <> "" And Player(Index).CommandLock <> Command Then
        'The only way this could trigger is if someone messed with the
        'code.  It's for when the server expects one and only one
        'command, such as the initial sign-on's "USER:" packet.
        Player(Index).DCReason = "CommandLock Conflict."
        Call DisconnectPlayer(Index)
        Exit Sub
    End If
    
    Select Case Command
        Case "USER:"
            Player(Index).SkipXOR = False
            'Okay, first take the server password (If there is one)
            If Len(ServerPassword) > 0 And Not Player(Index).ChangingTeams Then
                SentPassword = Trim(ChopString(Data, 16))
                If UCase(SentPassword) <> UCase(ServerPassword) Then
                    Call AddToQueue(Index, "PWDR:")
                    Call AddMessage("Bad server password received on channel " & Index)
                    Exit Sub
                End If
            End If
            With Player(Index)
                .Name = Trim(ChopString(Data, 20))
                
                If .Name <> CorrectText(.Name, True) Then
                    .DCReason = "Invalid Username"
                    Call AddToQueue(Index, "CMSG:Your user name is invalid.  Please try a different one.")
                    Call AddToQueue(Index, "BOOT:")
                    Exit Sub
                End If
                
                X = QueryName(.Name)
                If X > 0 Then
                    If SIDIsTempBanned(ServerDB.GetSIDByNum(X), Y) Then
                        Player(Index).LockDown = True
                        Call AddToQueue(Index, "BANU:" & "Your ban has " & CStr(Y) & " minutes remaining.")
                        Exit Sub
                    End If
                    If UserIsBanned(0, TempString, X) Then
                        Player(Index).LockDown = True
                        If TempString = "" Then TempString = DefaultBanMsg
                        Call AddToQueue(Index, "BANU:" & TempString)
                        Exit Sub
                    End If
                End If
                
                .Version = Trim(ChopString(Data, 7))
                .SID = ChopString(Data, 13)
                TempString = ChopString(Data, 16)
                If TempString = String$(16, vbNullChar) Then
                    SentPassword = vbNullString
                Else
                    SentPassword = nSpace(32)
                    Y = 0
                    For X = 1 To 32 Step 2
                        Y = Y + 1
                        Mid(SentPassword, X, 2) = FixedHex(Asc(Mid$(TempString, Y, 1)), 2)
                    Next X
                End If
                TempString = Chr2Bin(ChopString(Data, 2))
                For X = 1 To 10
                    .MessageAllow(X) = (Mid$(TempString, X, 1) = "1")
                Next X
                .Picture = Dec(ChopString(Data, 2))
                .ShowTeam = Bin2Bool(Val(ChopString(Data, 1)))
                TempString = Dec2Bin(Asc(ChopString(Data, 1)), 8)
                .SID = DecompressSID(.SID, Bin2Bool(ChopString(TempString, 1)))
                .Rank = Str(Bin2Dec(ChopString(TempString, 7)))
                Call Code.ReadBinArray(Dec(ChopString(Data, 2)), .Compatibility)
                TempString = ChopString(Data, 5)
                .Extra = WordFilter(Data)
                If .Picture < 1 Or .Picture > 15 Then .Picture = 1
                'Now, set all the default junk.
                .BattlingWith = 0
                .Unrated = False
                .Authority = GetAuthority(Index)
                Call GetRanking(Index)
                BattleReady(Index) = False
                'BattleOver(Index) = True
                ReDim .Ignore(0)
                .Speed = "WAIT"
            End With
            
            
            'Disconnect if version not up to date.
            If BetaRel = "" Then
                If Not IsVersionAt(Player(Index).Version, App.Major, App.Minor, App.Revision) Then
    '                If AllowOldVersions = 1 Then
    '                    Call AddToQueue(Index, "CMSG:Your version is out-of-date.  Please update as soon as possible to avoid problems.")
    '                Else
                        Call AddToQueue(Index, "CMSG:Your version is out-of-date.  This server requires the most current version to connect.")
                        Player(Index).DCReason = "Version out of date."
                        Call AddToQueue(Index, "BOOT:")
                        Exit Sub
    '                End If
                End If
            Else
                If Player(Index).Version <> You.ProgVersion Then
                    Call AddToQueue(Index, "CMSG:Your version is out-of-date.  This server requires the most current version to connect.")
                    Player(Index).DCReason = "Version out of date."
                    Call AddToQueue(Index, "BOOT:")
                    Exit Sub
                End If
            End If
            
            'Disconnect if name is taken
            For X = 1 To MaxUsers
                If IsLoaded(X) Then
                    If Player(X).Name = Player(Index).Name And X <> Index And Player(X).Name <> "" Then
                        Call AddToQueue(Index, "NAMR:")
                        Exit Sub
                    End If
                End If
            Next X
            
            If TempString <> ServerCode Then
                If InVBMode Then Stop
                Player(Index).DCReason = "SignOn Code conflict."
                Call DisconnectPlayer(Index)
            End If
            TempString = ""
            
            'Disconnect if Banned.
            If IsTempBanned(Index, X) Then
                Player(Index).LockDown = True
                Call AddToQueue(Index, "BANU:" & "Your ban has " & CStr(X) & " minutes remaining.")
                Exit Sub
            End If
            If UserIsBanned(Index, TempString) Then
                Player(Index).LockDown = True
                If TempString = "" Then TempString = DefaultBanMsg
                Call AddToQueue(Index, "BANU:" & TempString)
                Exit Sub
            End If
                
            Call AddToQueue(Index, "PING:")
            
            
            'Disconnect if new user and AllowNewUsers is off.
            If AllowNewUsers = 0 And GetUserPassword(Player(Index).Name) = "NULL" Then
                Call AddToQueue(Index, "NNPL:")
                Exit Sub
            End If
            Call ProcessLogon(Player(Index).Name)
            
            'Disconnect if bad password
            TempString = Trim(GetUserPassword(Player(Index).Name))
            If TempString = "NULL" Then
                If Len(SentPassword) = 32 Then
                    Call AddUserPassword(Player(Index).Name, SentPassword, Player(Index).SID)
                End If
            ElseIf TempString <> SentPassword Then
                Call AddToQueue(Index, "UPWR:")
                Call AddMessage("Bad logon for " & Player(Index).Name)
                Exit Sub
            End If
            'And if they made it this far, lets move on to the next step!
            Player(Index).CommandLock = "TEAM:"
            If Not Player(Index).ChangingTeams Then Call SendDBMod(Index)
            Call AddToQueue(Index, "RQTM:" & Index & "," & Player(Index).Authority & "," & Player(Index).Wins & "," & Player(Index).Losses & "," & Player(Index).Ties & "," & Player(Index).Disconnect)
         Case "TEAM:"
            Player(Index).CommandLock = ""
            Temp = Right(Temp, Len(Temp) - 5)
            If Not InterpretTeamInfo(Temp, Index) Then
                Call AddMessage("Hacked team detected.  Disconnecting " & Player(Index).Name)
                Exit Sub
            End If
            Call UpdateSID(Player(Index).Name, Player(Index).SID)
            If Player(Index).ChangingTeams Then
                Call ScriptMod.BlockExec(19, , Index)
                Call AddMessage(Player(Index).Name & " has loaded a new team.")
                Call SendAll(PreparePlayerData(Index, True, True), False, Index)
                Player(Index).ChangingTeams = False
                Call ScriptMod.BlockExec(20, , Index)
            Else
                Cancel = False
                Call ScriptMod.BlockExec(3, Cancel, Index)
                StatusBar1.Panels(1).Text = "Finished recieving information for Player: " & Player(Index).Name
                Call AddMessage(Player(Index).Name & " (" & Player(Index).Address & " - " & Player(Index).DNSAddress & ") is fully connected.")
                If Not Cancel Then 'In case player was script kicked
                    Call SendAll(PreparePlayerData(Index), True, Index)
                    If ServerMessage <> "" Then Call AddToQueue(Index, "SMSG:" & ServerMessage)
                    Call SendFullList(Index)
                    Player(Index).Active = True
                    Call SendAll(MakeList) ', True, Index)
                    Call ScriptMod.BlockExec(4, , Index)
                End If
            End If
            'Scan if there's an battle this user needs to finish
            Temp = vbNullString
            For X = 1 To UBound(ServerBattle)
                With ServerBattle(X)
                    If .DisconStall Then
                        If StrComp(.ModHash, Player(Index).ModHash) = 0 Or Not .ModsInvolved Then
                            If .PlayerName(1) = Player(Index).Name Then
                                Temp = .TeamChecksum1
                                Y = .Player2
                                Z = 1
                            End If
                            If .PlayerName(2) = Player(Index).Name Then
                                Temp = .TeamChecksum2
                                Y = .Player1
                                Z = 2
                            End If
                            If Len(Temp) > 0 Then
                                If Temp = Player(Index).TeamChecksum Then
                                    If Z = 1 Then .Player1 = Index Else .Player2 = Index
                                    Player(Index).BattleID = X
                                    Player(Index).BattlingWith = Y
                                    Player(Y).BattlingWith = Index
                                    Call SendAll("PRBS:" & Y & "," & Index)
                                    Call AddToQueue(Index, "SBAT:" & Chr$(Y) & Chr$(X) & CStr(Z) & .ActNum)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End With
            Next X
            Call RefreshListing
        Case "TUPD:"
            Call ScriptMod.BlockExec(19, , Index)
            If Not InterpretTeamInfo(Data, Index) Then
                Call AddMessage("Hacked team detected.  Disconnecting " & Player(Index).Name)
                Exit Sub
            End If
            StatusBar1.Panels(1).Text = "Team changed for Player: " & Player(Index).Name
            Call AddMessage(Player(Index).Name & " changed teams.")
            Call SendAll(PreparePlayerData(Index, True, True), False, Index)
            Call ScriptMod.BlockExec(20, , Index)
        Case "RKEY:"
            Player(Index).ChangingTeams = True
            Call AddToQueue(Index, "RPUD:" & ServerCode) 'Request Player Update
        Case "PWDS:"
            SentPassword = Right(Temp, Len(Temp) - 5)
            If UCase(SentPassword) = UCase(ServerPassword) Then
                Call AddToQueue(Index, "REQN:" & You.ProgVersion & "," & MaxUsers & "," & FloodTolerance & "," & ListView1.ListItems.count + 1)
            Else
                Call AddToQueue(Index, "PWDR:")
                Call AddMessage("Bad server password received on channel " & Index)
            End If
        Case "NAME:"
            Temp = Right(Temp, Len(Temp) - 5)
            If InStr(1, Temp, ":") > 1 Then
                Player(Index).Name = WordFilter(Left(Temp, InStr(1, Temp, ":") - 1))
                Player(Index).SID = Right(Temp, Len(Temp) - InStr(1, Temp, ":"))
            Else
                Player(Index).Name = WordFilter(Temp)
                Player(Index).SID = Player(Index).Name
            End If
            Player(Index).BattlingWith = 0
            Player(Index).Unrated = False
            Player(Index).Authority = GetAuthority(Index)
            Call GetRanking(Index)
            BattleReady(Index) = False
            'BattleOver(Index) = True
            If Not UserIsBanned(Index) Then
                Call AddToQueue(Index, "REQV:")
            Else
                Call AddToQueue(Index, "YNUM:" & Index & "," & Player(Index).Wins & "," & Player(Index).Losses & "," & Player(Index).Ties & "," & Player(Index).Disconnect)
                Call AddToQueue(Index, "BANU:")
            End If
        Case "LIST:"
            Call AddToQueue(Index, MakeList)
        Case "PREQ:"
            StatusBar1.Panels(1).Text = "Player" & Player(Index).Name & " has requested player data."
            X = Val(Data)
            Call AddToQueue(Index, PreparePlayerData(X))
        Case "IMCH:"
            X = Asc(ChopString(Data, 1))
            If Not IsIgnoring(X, Index) Then
                Call AddToQueue(X, "IMCH:" & Chr(Index) & WordFilter(Data))
            End If
        Case "CHLN:"
            X = Dec(ChopString(Data, 3))
            If Not IsLoaded(X) Then Exit Sub
            If Val(ChopString(Data, 1)) <> Player(X).GameVersion Then
                Call AddToQueue(Index, "SREF:" & X)
                Exit Sub
            End If
            
            If Player(Index).GameVersion = nbModAdv Or Player(X).GameVersion = nbModAdv Then
                If StrComp(Player(Index).ModHash, DBModHash) <> 0 Then
                    Call AddToQueue(Index, "SREF:" & X)
                    Call AddToQueue(Index, "CMSG:Your currently loaded database mod doesn't match the server's.  Please sign off and on before challenging.")
                    Exit Sub
                End If
                If StrComp(Player(X).ModHash, DBModHash) <> 0 Then
                    Call AddToQueue(Index, "SREF:" & X)
                    Call AddToQueue(Index, "CMSG:This player's currently loaded database mod doesn't match the server's and needs to sign off and on before accepting challenges.")
                    Exit Sub
                End If
            End If
            
            Cancel = False
            Call ScriptMod.BlockExec(21, Cancel, Index, X, StrReverse$(Dec2Bin(Dec(Right$(Data, 8)), UBound(RuleText))))
            If Cancel Then
                Call AddToQueue(Index, "SREF:" & X)
                Exit Sub
            End If
            StatusBar1.Panels(1).Text = "Player " & Player(Index).Name & " has issued a challenge."
            If IsIgnoring(X, Index) Then
                Call AddToQueue(Index, "PREF:" & X)
                Call AddToQueue(Index, "CMSG:Because this player is ignoring you, all challenges are automatically refused.")
                Exit Sub
            End If
            If X = Index Then
                'Call AddMessage("Bad challenge request received from " & Player(X).Name)
                Player(Index).DCReason = "Bad challenge request."
                Call DisconnectPlayer(X)
            End If
            Call AddMessage("Challenge issued from " & Player(Index).Name & " to " & Player(X).Name)
            Call AddToQueue(X, "CHLN:" & FixedHex(Index, 3) & Data & Player(Index).Speed)
            Call ScriptMod.BlockExec(22, , Index, X, StrReverse$(Dec2Bin(Right$(Data, 8), UBound(RuleText))))
        Case "PBSY:"
            StatusBar1.Panels(1).Text = "Player " & Player(Index).Name & " is busy."
            X = Val(Right(Temp, Len(Temp) - 5))
            Call AddToQueue(X, "PBSY:" & Index)
            Call AddMessage(Player(Index).Name & " responded to the challenge as busy.")
        Case "PREF:"
            StatusBar1.Panels(1).Text = "Player " & Player(Index).Name & " has refused a challenge."
            X = Val(Right(Temp, Len(Temp) - 5))
            Call AddToQueue(X, "PREF:" & Index)
            Call AddMessage(Player(Index).Name & " refused the challenge.")
        Case "PACC:"
            X = Dec(ChopString(Data, 3))
            If Player(Index).GameVersion = nbModAdv Or Player(X).GameVersion = nbModAdv Then
                If StrComp(Player(Index).ModHash, DBModHash) <> 0 Or StrComp(Player(X).ModHash, DBModHash) <> 0 Then
                    Call AddToQueue(Index, "SREF:" & X)
                    Call AddToQueue(X, "SREF:" & X)
                    Call AddToQueue(Index, "CMSG:Your currently loaded database mod doesn't match the server's.  Please sign off and on before challenging.")
                    Call AddToQueue(X, "CMSG:Your currently loaded database mod doesn't match the server's.  Please sign off and on before challenging.")
                    Exit Sub
                End If
            End If
            Call ScriptMod.BlockExec(9, , X, Index, StrReverse$(Dec2Bin(Dec(Data), UBound(RuleText))))
            Call AddMessage("Player " & Player(Index).Name & " has accepted a challenge.")
            If Not Player(X).Active Then
                Call AddToQueue(Index, "PBSY:" & X)
                Exit Sub
            End If
            For Y = 1 To UBound(ServerBattle)
                With ServerBattle(Y)
                    If .Player1 = 0 Or .Player2 = 0 Then
                        Exit For
                    Else
                        If Player(.Player1).BattleID <> Y And Player(.Player2).BattleID <> Y Then
                            Exit For
                        End If
                    End If
                End With
            Next
            If Y > UBound(ServerBattle) Then
                'This should technically never happen, but just in case..
                ReDim Preserve ServerBattle(Y)
                Set ServerBattle(Y) = New BattleData
            End If
            Call AddMessage(Player(Index).Name & " has accepted the challenge.")
            Player(Index).BattlingWith = X
            Player(X).BattlingWith = Index
            With ServerBattle(Y)
                .ResetBattle
                .ModHash = Player(Index).ModHash
                .Player1 = X
                .Player2 = Index
                Z = Dec(ChopString(Data, 1)) - 1
                .BattleMode = Cap(Z, 2)
                If Z = 3 Then .ActNum = 4 Else .ActNum = 2
                Call .SetVer(1, Player(X).GFXVer)
                Call .SetVer(2, Player(Index).GFXVer)
                .TeamChecksum1 = Player(X).TeamChecksum
                .TeamChecksum2 = Player(Index).TeamChecksum
                Z = Dec(ChopString(Data, 1))
                If Z = 0 Then Z = GetSmallRnd(UBound(TerrainText)) Else Z = Z - 1
                .Terrain = Z
                .Rules = Data
                Player(Index).BattleID = Y
                Player(X).BattleID = Y
                Call AddToQueue(X, "PACC:" & Index)
                Call SendAll("PRBS:" & X & "," & Index)
                Call AddToQueue(X, "SBAT:" & Chr$(Index) & Chr$(Y) & "1" & .ActNum)
                Call AddToQueue(Index, "SBAT:" & Chr$(X) & Chr$(Y) & "2" & .ActNum)
                Call RefreshListing
            End With
        Case "REQW:"
            X = Player(Val(Data)).BattleID
            With ServerBattle(X)
                If X = 0 Then
                    Call AddToQueue(Index, "WREF:1")
                ElseIf .NoWatch And Player(Index).Authority < 2 Then
                    Call AddToQueue(Index, "WREF:2")
                ElseIf .IsWatching(Index) Then
                    Call AddToQueue(Index, "WREF:3")
                ElseIf Not .WatchOK Then
                    Call AddToQueue(Index, "WREF:4")
                ElseIf .ModsInvolved And StrComp(.ModHash, Player(Index).ModHash) <> 0 Then
                    Call AddToQueue(Index, "WREF:5")
                Else
                    Player(Index).WatchID = X
                    Call AddToQueue(Index, "SWAT:" & FixedHex(X, 3) & FixedHex(.Player1, 3) & FixedHex(.Player2, 3) & CStr(.ActNum)) 'CStr(Bool2Bin(.NoWatchChat)))
                End If
            End With
        Case "WWRDY"
            X = Player(Index).WatchID
            If X = 0 Then Exit Sub
            Player(Index).WatchID = 0
            With ServerBattle(X)
                Temp = .VariableSync
                Do Until Len(Temp) <= 230
                    Call AddToQueue(Index, "WATCH" & FixedHex(X, 3) & "SYNC:" & ChopString(Temp, 230) & vbNullChar)
                Loop
                Call AddToQueue(Index, "WATCH" & FixedHex(X, 3) & "SYNC:" & Temp)
'                Temp = .TextLog
'                Do Until Len(Temp) <= 230
'                    Call AddToQueue(Index, "WATCH" & FixedHex(X, 3) & "TLOG:" & ChopString(Temp, 230) & vbNullChar)
'                Loop
'                Call AddToQueue(Index, "WATCH" & FixedHex(X, 3) & "TLOG:" & Temp)
                Call AddToQueue(.Player1, "RELAYWTCH:" & FixedHex(Index, 3))
                Call AddToQueue(.Player2, "RELAYWTCH:" & FixedHex(Index, 3))
                Call .NewWatcher(Index)
            End With
        Case "DONW:"
            If ServerBattle(Val(Data)).IsWatching(Index) Then
                Call ServerBattle(Val(Data)).RemoveWatcher(Index)
                Call AddToQueue(ServerBattle(Val(Data)).Player1, "RELAYDONW:" & "0" & FixedHex(Index, 3))
                Call AddToQueue(ServerBattle(Val(Data)).Player2, "RELAYDONW:" & "0" & FixedHex(Index, 3))
            End If
        Case "PCAN:"
            StatusBar1.Panels(1).Text = "Player " & Player(Index).Name & " has already cancelled."
            X = Val(Right(Temp, Len(Temp) - 5))
            Call AddToQueue(X, "PCAN:" & Index)
            Call AddMessage(Player(Index).Name & " cancelled before accepting.")
        Case "AWAY:"
            If Player(Index).BattlingWith = 1025 Then
                Call DisconnectPlayer(Index)
                Exit Sub
            End If
            Cancel = False
            Call ScriptMod.BlockExec(13, Cancel, Index)
            If Cancel Then Exit Sub
            Player(Index).BattlingWith = 1025
            Call AddMessage(Player(Index).Name & " is away.")
            Call SendAll("AWAY:" & Index)
            Call RefreshListing
            Call ScriptMod.BlockExec(14, , Index)
'        Case "BOVER"
'            X = Player(Index).BattlingWith
'            If X > 0 Then
'                If Not BattleOver(Index, True) Then
'                    Call ChangePlayerStats(Index, 4)
'                    If Player(X).Name <> "" And Player(Index).Name <> "" Then
'                        Call AddMessage(Player(Index).Name & " has finished battling with " & Player(X).Name)
'                        SendString = "CMSG:" & Player(Index).Name & " finished battling with " & Player(X).Name
'                        Call SendAll(SendString)
'                        'Call SendAll("BACK:" & Player(Index).BattlingWith)
'                    End If
'                End If
'                Call SendAll("BACK:" & Index)
'                Player(Player(Index).BattlingWith).BattlingWith = 0
'                Player(Player(Index).BattlingWith).Unrated = False
'                BattleReady(Player(Index).BattlingWith) = False
'                Player(Index).BattlingWith = 0
'                BattleReady(Index) = False
'            End If
'            Call RefreshListing
        Case "BACK:"
            X = Player(Index).BattlingWith
            If X = 1025 Then
                Call SendAll("BACK:" & Index)
            ElseIf X > 0 Then
                With ServerBattle(Player(Index).BattleID)
                    If .Unrated Then Y = 1 Else Y = 0
                    If .Player1 = Index Then Z = 1 Else Z = 2
                    If .DisconStall Then
                        Call SendAll("BACK:" & Index)
                        .DisconStall = False
                    Else
                        If Not BattleOver(Index, True) Then
                            Call ScriptMod.BlockExec(7, , Player(Index).BattlingWith, Index, "FOR" & IIf(Y = 1, "*", ""))
                            If Y = 0 Then
                                Call ChangePlayerStats(.PlayerName(Z), 2)
                                Call ChangePlayerStats(.PlayerName(OtherTeam(Z)), 1)
                            End If
                            Call ScriptMod.BlockExec(8, , Player(Index).BattlingWith, Index, "FOR" & IIf(Y = 1, "*", ""))
        
                            If Player(X).Name <> "" And Player(Index).Name <> "" Then
                                Call AddMessage(Player(Index).Name & " has forfeited the battle with " & Player(X).Name)
                                SendString = "CMSG:" & Player(Index).Name & " has forfeited the battle with " & Player(X).Name
                                Call SendAll(SendString, , , 4)
                            End If
                        End If
                        Call SendAll("BACK:" & Index)
                    End If
                End With
            End If
            BattleReady(Index) = False
            Player(Index).BattlingWith = 0
            Player(Index).BattleID = 0
            Player(Index).Unrated = False
            Call AddMessage(Player(Index).Name & " is back.")
            Call RefreshListing
        Case "MKCK:"
            If Player(Index).Authority <= Player(Val(Data)).Authority Then
                Call AddToQueue(Index, "CMSG:" & "Not authorized to perform that action.")
            Else
                Cancel = False
                Call ScriptMod.BlockExec(15, Cancel, Index, Val(Data))
                If Cancel Then Exit Sub
                Call SendAll("MKCK:" & Index & ":" & Val(Right(Temp, Len(Temp) - 5)))
                Call AddToQueue(Val(Right(Temp, Len(Temp) - 5)), "KCKU:")
                Call ScriptMod.BlockExec(16, , Index, Val(Data))
            End If
        Case "MBAN:"
            If Player(Index).Authority <= Player(Val(Right(Temp, Len(Temp) - 5))).Authority Or Player(Index).Authority < 2 Then
                Call AddToQueue(Index, "CMSG:" & "Not authorized to perform that action.")
            Else
                Call SendAll("MBAN:" & Index & ":" & Val(Right(Temp, Len(Temp) - 5)))
                Call BanUser(Val(Right(Temp, Len(Temp) - 5)))
                Call ScriptMod.BlockExec(18, Cancel, Index, Val(Data))
            End If
        Case "MBN2:"
            If Player(Index).Authority <= Player(Val(Right(Temp, Len(Temp) - 5))).Authority Or Player(Index).Authority < 2 Then
                Call AddToQueue(Index, "CMSG:" & "Not authorized to perform that action.")
            Else
                Cancel = False
                Call ScriptMod.BlockExec(17, Cancel, Index, Val(Data))
                If Cancel Then Exit Sub
                Call SendAll("MBAN:" & Index & ":" & Val(Right(Temp, Len(Temp) - 5)))
                Call SIDBanUser(Val(Right(Temp, Len(Temp) - 5)))
                Call ScriptMod.BlockExec(18, Cancel, Index, Val(Data))
            End If
        Case "IPBN:"
            For X = 1 To MaxUsers
                If LCase(Player(X).Name) = LCase(Data) Then
                    Data = Player(X).Name
                    Exit For
                End If
            Next
            If X = MaxUsers + 1 Then
                X = 2
            Else
                If Player(Index).Authority <= Player(X).Authority Or Player(Index).Authority < 3 Then
                    X = 3
                ElseIf Player(Index).Authority <= ServerDB.GetMaxAuth(ServerDB.QueryName(Player(X).Name)) Then
                    X = 3
                Else
                    Z = IIf(ServerDB.AddIPBan(Player(X).Address), 0, 1)
                    If Z = 0 Then
                        Call SendAll("MBAN:" & Chr$(Index) & Data)
                        For Y = 1 To MaxUsers
                            If Player(Y).Address = Player(X).Address Then
                                Call AddToQueue(Y, "BANU:")
                            End If
                        Next Y
                    End If
                    X = Z
                End If
            End If
            Call AddToQueue(Index, "BRLT:" & X)
        Case "IDBN:"
            X = ServerDB.QueryName(Data)
            If X = 0 Then
                X = 2
            Else
                Data = ServerDB.GetNameByNum(X)
                Y = ServerDB.GetAuthByNum(X)
                If Player(Index).Authority <= Y Or Player(Index).Authority < 3 Then
                    X = 3
                ElseIf Player(Index).Authority <= ServerDB.GetMaxAuth(X) Then
                    X = 3
                Else
                    Temp = ServerDB.GetSIDByNum(X)
                    If SIDIsTempBanned(Temp, X) Then DelSIDBanBySID Temp
                    X = IIf(ServerDB.AddSIDBan(Data), 0, 1)
                    If X = 0 Then
                        Call SendAll("MBAN:" & Chr$(Index) & Data)
                        For Y = 1 To MaxUsers
                            If Player(Y).SID = Temp Then
                                Call AddToQueue(Y, "BANU:")
                            End If
                        Next Y
                    End If
                End If
            End If
            Call AddToQueue(Index, "BRLT:" & X)
        Case "TMPB:"
            Z = Val(ChopString(Data, 4))
            X = ServerDB.QueryName(Data)
            If X = 0 Then
                X = 2
            Else
                Data = ServerDB.GetNameByNum(X)
                Y = ServerDB.GetAuthByNum(X)
                If Player(Index).Authority <= Y Or Player(Index).Authority < 2 Then
                    X = 3
                ElseIf Player(Index).Authority <= ServerDB.GetMaxAuth(X) Then
                    X = 3
                Else
                    Temp = ServerDB.GetSIDByNum(X)
                    X = IIf(ServerDB.AddSIDBan(Data), 4, 1)
                    If X = 4 Then
                        Call SendAll("MTBN:" & Chr$(Index) & Int2Str(Z) & Data)
                        For Y = 1 To UBound(TempBanInfo)
                            If TempBanInfo(Y).TimeLeft = 0 Then Exit For
                        Next Y
                        If Y = UBound(TempBanInfo) + 1 Then ReDim Preserve TempBanInfo(Y)
                        TempBanInfo(Y).PName = Data
                        TempBanInfo(Y).SID = Temp
                        TempBanInfo(Y).TimeLeft = Z
                        For Y = 1 To MaxUsers
                            If Player(Y).SID = Temp Then
                                Call AddToQueue(Y, "TBAN:" & Z)
                            End If
                        Next Y
                    End If
                End If
            End If
            Call AddToQueue(Index, "BRLT:" & X)
        Case "LOOK:"
            If Player(Index).Authority < 2 Then
                Call DisconnectPlayer(Index)
                Exit Sub
            End If
            Temp2 = ""
            For X = 1 To MaxUsers
                If LCase(Player(X).Name) = LCase(Data) Then
                    Temp2 = Player(X).Address
                    Exit For
                End If
            Next X
            Temp = ServerDB.GetLookupString(Data, Temp2)
            Temp = Temp & "|" & Temp2
            Call AddToQueue(Index, "LOOK:" & Temp)
        Case "ALIA:"
            'Call AddToQueue(Index, "ALIA:MasamuneXGP        ")
            If Player(Index).Authority < 2 Then
                Call DisconnectPlayer(Index)
                Exit Sub
            End If
            Erase Packet
            Packet = ServerDB.GetAliases(Data)
            If (Not Packet) <> -1 Then 'The array is actually defined.
                X = 0
                Temp = Join(Packet, "") & " "
                Do While Len(Temp) > 240
                    Packet(X) = "ALIA:" & ChopString(Temp, 240)
                    X = X + 1
                Loop
                Packet(X) = "ALIA:" & Temp
                For X = 0 To X
                    Call AddToQueue(Index, Packet(X))
                Next X
            Else
                Call AddToQueue(Index, "ALIA:X")
            End If
        Case "UBAN:"
            If Player(Index).Authority > 1 Then
                If Data <> "" Then
                    Temp = ChopString(Data, 1)
                    If Temp = "I" Then
                        If Player(Index).Authority = 3 Then
                            Call ServerDB.DelIPBan(Data)
                        End If
                    Else
                        For X = 1 To ServerDB.GetSIDBanMax
                            Temp2 = ServerDB.GetSIDNameByNum(X)
                            For Y = 1 To UBound(TempBanInfo)
                                If UCase(TempBanInfo(Y).PName) = UCase(Data) Then Exit For
                            Next Y
                            If Y = UBound(TempBanInfo) + 1 Then
                                If Player(Index).Authority = 3 Then Call ServerDB.DelSIDBan(Data)
                            Else
                                If Player(Index).Authority > 1 Then Call ServerDB.DelSIDBan(Data)
                            End If
                        Next X
                    End If
                End If
            
                Temp = ""
                For X = 1 To ServerDB.GetSIDBanMax
                    Temp2 = ServerDB.GetSIDNameByNum(X)
                    For Y = 1 To UBound(TempBanInfo)
                        If UCase(TempBanInfo(Y).PName) = UCase(Temp2) Then Exit For
                    Next Y
                    If Y <> UBound(TempBanInfo) + 1 Then
                        Temp = Temp & "," & Temp2 & ":" & TempBanInfo(Y).TimeLeft
                    End If
                Next X
                Temp = Temp & "|"
                For X = 1 To ServerDB.GetSIDBanMax
                    Temp2 = ServerDB.GetSIDNameByNum(X)
                    For Y = 1 To UBound(TempBanInfo)
                        If UCase(TempBanInfo(Y).PName) = UCase(Temp2) Then Exit For
                    Next Y
                    If Y = UBound(TempBanInfo) + 1 Then
                        Temp = Temp & "," & Temp2
                    End If
                Next X
                Temp = Temp & "|"
                For X = 1 To ServerDB.GetIPBanMax
                    Temp2 = ServerDB.GetIPByNum(X)
                    Temp = Temp & "," & ServerDB.GetIPByNum(X)
                Next X
                ReDim Packet(1)
                While Len(Temp) > 200
                    Packet(UBound(Packet)) = "BANL:" & ChopString(Temp, 200) & vbNullChar
                    ReDim Preserve Packet(UBound(Packet) + 1)
                Wend
                Packet(UBound(Packet)) = "BANL:" & Temp
                For X = 1 To UBound(Packet)
                    Call AddToQueue(Index, Packet(X))
                Next X
            Else
                Call DisconnectPlayer(Index)
            End If
            
        'This scans the relay messages for certain things.
        'Most messages will just pass through.
        Case "RELAY"
            BattleID = Asc(ChopString(Data, 1))
            With ServerBattle(BattleID)
                If .Player1 = Index Then
                    PNum = .Player2
                ElseIf .Player2 = Index Then
                    PNum = .Player1
                ElseIf .IsWatching(Index) And Left(Data, 5) = "WMSG:" And Not .NoWatchChat Then
                    PNum = Index
                Else
                    'Uh, oh - somebody's hacking...
                    If InVBMode Then Stop
                    Player(Index).DCReason = "Illegal Relay."
                    Call DisconnectPlayer(Index)
                    Exit Sub
                End If
            End With
            RelayString = Data
            Call RelayScanner(Index, PNum, RelayString, BattleID)
        Case Else
            'Uh, oh - somebody's hacking...
            If InVBMode Then Stop
            'Call AddMessage("Bad packet received on Channel " & Index & ".  Disconnecting.")
            Player(Index).DCReason = "Bad packet received."
            Call AddMessage("Bad packet: " & Command & Data, , True)
            Call DisconnectPlayer(Index)
    End Select
End Sub

Public Sub DisconnectPlayer(ByVal Number As Integer)
    Dim X As Integer
    Dim NumberConnected As Integer
    Dim Temp As String
    '>>> Call WriteDebugLog("Disconnecting Player " & Player(Number).Name & ", slot " & Number & ")")
    On Error Resume Next
    Disconnecting(Number) = True
    Call ScriptMod.BlockExec(5, , Number)
    '>>> Call WriteDebugLog("DP: Checking Ignore lists.")
    For X = 1 To UBound(Player)
        ''>>> Call WriteDebugLog("DP: Ignore Loop- " & X)
        If IsIgnoring(X, Number) Then Call StopIgnore(X, Number)
    Next X
    '>>> Call WriteDebugLog("DP: Checking Watch lists.")
    For X = 1 To UBound(ServerBattle)
        If ServerBattle(X).IsWatching(Number) Then
            Call ServerBattle(X).RemoveWatcher(Number)
'            Call AddToQueue(ServerBattle(X).Player1, "RELAYDONW:" & FixedHex(Number, 3))
'            Call AddToQueue(ServerBattle(X).Player2, "RELAYDONW:" & FixedHex(Number, 3))
        End If
    Next X
    '>>> Call WriteDebugLog("DP: Checking BattleOver.")
    If Not BattleOver(Number) Then
        With ServerBattle(Player(Number).BattleID)
            If .DisconStall Then
                X = IIf(.Unrated, 1, 0)
                Call ScriptMod.BlockExec(7, , Player(Number).BattlingWith, Number, "DIS" & IIf(X = 1, "*", ""))
'                If X = 0 Then
'                    If .Player1 = Number Then
'                        Call ChangePlayerStats(.PlayerName(2), 4)
'                    Else
'                        Call ChangePlayerStats(.PlayerName(1), 4)
'                    End If
'                End If
                Call ScriptMod.BlockExec(8, , Player(Number).BattlingWith, Number, "DIS" & IIf(X = 1, "*", ""))
                BattleReady(Player(Number).BattlingWith) = False
                .DisconStall = False
            Else
                .SavedSyncStr = .VariableSync
                .DisconStall = True
            End If
        End With
    End If
'    '>>> Call WriteDebugLog("DP: Checking IP Ban status.")
'    If Not IPIsBanned(Player(Number).Address) Then
        Call SendAll("PDIS:" & Number)
'    End If
    '>>> Call WriteDebugLog("DP: Adding Disconnect message.  Reason: " & Player(Number).DCReason)
    If Player(Number).Name <> "" Then
        Temp = Player(Number).Name & " disconnected."
        If Player(Number).DCReason <> "" Then
            Temp = Temp & "  (" & Player(Number).DCReason & ")"
        End If
        Call AddMessage(Temp)
    End If
    '>>> Call WriteDebugLog("DP: Reseting Player variable.")
    Player(Number) = BlankPlayer
    BattleReady(Number) = False
    SentPing(Number) = False
    Chances(Number) = 0
    '>>> Call WriteDebugLog("DP: Closing Socket.")
    If IsLoaded(Number) Then
        ClientSocket(Number).Close
        'Unload ClientSocket(Number)
        IsLoaded(Number) = False
    End If
    '>>> Call WriteDebugLog("DP: Refreshing Listing.")
    Call RefreshListing
    Disconnecting(Number) = False
    Call ScriptMod.BlockExec(6, , Number)
    '>>> Call WriteDebugLog("Player successfully disconnected.")
End Sub

Sub RefreshListing(Optional ByVal ShowList As Boolean = False)
    Dim ConnectedString As String
    Dim X As Integer
    'Dim Index As Integer
    Dim TempItem As ListItem
    Dim Temp As String
    'On Error Resume Next
    'Index = 0
'    ListView1.ListItems.Clear
'    ListView1.Sorted = False
    ConnectedString = "Current Players: "
    For X = 1 To MaxUsers
        If Player(X).Name <> "" Then
            'Index = Index + 1
            If Player(X).Authority = 0 Then Player(X).Authority = 1
            If PListed(X) Then
                Set TempItem = ListView1.ListItems("USER: " & CStr(X))
            Else
                Set TempItem = ListView1.ListItems.Add(, "USER: " & CStr(X), Player(X).Name)
            End If
            Temp = Trim(Player(X).Name & " " & String(Player(X).Authority - 1, "*"))
            If TempItem.Text <> Temp Then TempItem.Text = Temp
            If TempItem.SubItems(1) <> CStr(X) Then TempItem.SubItems(1) = CStr(X)
            If Player(X).Picture > 0 And TempItem.SmallIcon <> Player(X).Picture Then TempItem.SmallIcon = Player(X).Picture
            If Player(X).BattlingWith > 0 Then
                TempItem.Ghosted = True
                If Player(X).BattlingWith = 1025 Then
                    TempItem.ToolTipText = "Away"
                Else
                    TempItem.ToolTipText = "Battling with " & Player(Player(X).BattlingWith).Name
                End If
            Else
                TempItem.Ghosted = False
                TempItem.ToolTipText = "Not battling"
            End If
            ConnectedString = ConnectedString + Player(X).Name & " - " & X & " "
        Else
            If PListed(X) Then ListView1.ListItems.Remove "USER: " & CStr(X)
        End If
    Next X
    If ShowList Then Call AddMessage(ConnectedString)
    ListView1.Sorted = True
End Sub
Private Function PListed(PNum As Integer) As Boolean
    Dim X As Integer
    On Error Resume Next
    Err.Number = 0
    X = ListView1.ListItems("USER: " & CStr(PNum)).Index
    PListed = (Err.Number = 0)
End Function

Private Sub Messages_Change()
    On Error Resume Next
    
    Messages.SelStart = Len(Messages.Text)
End Sub

Sub ExitProgram(ByRef Cancel As Integer)
    Dim YourChoice As Integer
    Dim X As Long
    On Error Resume Next
    YourChoice = vbYes
    If ConfirmExit Then YourChoice = MsgBox("This will disconnect all players and exit the program." & vbNewLine & "Are you sure you want to do this?", vbQuestion + vbYesNo + vbDefaultButton2, "Disconnect?")
    If YourChoice = vbNo Then
        Cancel = True
    Else
        ConfirmExit = False
        ShuttingDown = True
        StatusBar1.Panels(1).Text = "Closing Connections..."
        If ListView1.ListItems.count <> 0 Then
            Cancel = True
            Call SendAll("SVRQU")
        End If
        For X = 1 To UBound(TempBanInfo)
            If TempBanInfo(X).PName <> "" Then
                Call ServerDB.DelSIDBan(TempBanInfo(X).PName)
            End If
        Next X
        MainContainer.ServerKiller.Enabled = True
    End If
End Sub

Function PreparePlayerData(ByVal Number As Integer, Optional ByVal Update As Boolean = False, Optional ByVal TeamChange As Boolean = False) As String
    Dim BuildString As String
    Dim X As Byte
    On Error Resume Next
    
    
    BuildString = ""
    If Player(Number).Name = "" Then
        PreparePlayerData = "PDIS:" & Number
        Exit Function
    End If
    With Player(Number)
        BuildString = BuildString & Dec2Bin(Number, 8)
        BuildString = BuildString & Dec2Bin(.GameVersion, 3)
        BuildString = BuildString & Dec2Bin(.Picture, 4)
        BuildString = BuildString & Dec2Bin(.GFXVer, 4)
        BuildString = BuildString & Dec2Bin(.Authority, 2)
        BuildString = BuildString & Bool2Bin(.ShowTeam)
        BuildString = BuildString & Bool2Bin(.StadiumOK)
        For X = 1 To 6
            BuildString = BuildString & Dec2Bin(.PKMN(X), 9)
            BuildString = BuildString & Dec2Bin(.PokeData(X).UnownLetter, 5)
            BuildString = BuildString & Bool2Bin(.PokeData(X).Shiny)
        Next
        BuildString = BuildString & Dec2Bin(.Wins, 16)
        BuildString = BuildString & Dec2Bin(.Losses, 16)
        BuildString = BuildString & Dec2Bin(.Ties, 16)
        BuildString = BuildString & Dec2Bin(.Disconnect, 16)
        BuildString = BuildString & Dec2Bin(MakeBinArray(.Compatibility), 16)
        BuildString = BuildString & Dec2Bin(Val(.Rank), 16)
        BuildString = BuildString & Dec2Bin(.BattlingWith, 11)
        BuildString = Bin2Chr(BuildString)
        BuildString = BuildString & Pad(.Name, 20)
        BuildString = BuildString & Pad(.Version, 8)
        BuildString = BuildString & .Extra
    End With
    If Update Then
        BuildString = "PUPD:" & IIf(TeamChange, "T", "N") & BuildString
    Else
        BuildString = "PNFO:" & BuildString
    End If
    PreparePlayerData = BuildString
End Function

Private Sub SendData(ByVal Channel As Integer, SendMe As SendAllType)
    Dim XORSendMe As String
    Dim SPacket As Boolean
    On Error GoTo ErrorTrap
    If Channel > 0 Then If Not IsLoaded(Channel) Then Exit Sub
    
    If Player(Channel).LockDown And Left(SendMe.Data, 5) <> "BANU:" Then Exit Sub
    If Player(Channel).KickTimer <> 0 Then Exit Sub
    
    'If SendMe.Data = "CMSG:TESTING:5" Then MsgBox "Sending: " & XORDecrypt(Right$(SendMe.Packet, Len(SendMe.Packet) - 1))
    
    Call AddMessage("Sent " & SendMe.Data & " on channel " & Channel, , True)
    If Left(SendMe.Data, 5) = "PING:" Then Player(Channel).PingTime = Timer
    ClientSocket(Channel).SendData SendMe.Packet
    Player(Channel).KickTimer = 5
    Select Case Left(SendMe.Data, 5)
    Case "BOOT:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Silently kicked."
    Case "KCKU:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Kicked."
    Case "BANU:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Banned."
    Case "ILLM:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Illegal team."
    Case "NAMR:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Name taken."
    Case "NNPL:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "No New Players active."
    Case "SVRQU"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Server Shutdown."
    Case "UPSR:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "User password refused."
    Case "NOIP:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Too many connections."
    Case "TBAN:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Temp Banned."
    Case "UPWR:"
        If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Bad User Password."
    Case Else
        Player(Channel).KickTimer = 0
    End Select
    'Debug.Print "ServerWindow.SendData: " & Channel & " - " & SendMe.Data
    Call WriteDebugLog("ServerWindow.SendData: " & SendMe.Data & " (To Player " & Channel & ")")
    Exit Sub
ErrorTrap:
    If Player(Channel).DCReason = "" Then Player(Channel).DCReason = "Data send error."
    Call DisconnectPlayer(Channel)
End Sub

Public Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal RawMessage As Boolean = False)
    Dim X As Integer
    Dim LineCount As Integer
    Dim Cancel As Boolean
    On Error Resume Next
    If DebugMessage = True And DebugMode = False Then Exit Sub
    If RawMessage = True And RawData = False Then Exit Sub
    If Not DebugMessage And Not RawMessage Then Call ScriptMod.BlockExec(1, Cancel, , , Message)
    If Cancel Then Exit Sub
    RTB.AddMessage Time & " - " & Message
    If AutoLogging = 1 And Date <> LogDate Then
        StopAutoLog
        StartAutoLog
    End If
    If Logging Then Print #LogFileNum, Time & " - " & Message
    If Not DebugMessage And Not RawMessage And ProcessScript Then Call ScriptMod.BlockExec(2, , , , Message)
End Sub

Sub SendAll(ByVal Data As String, Optional ByVal NoRepeat As Boolean = False, Optional ByVal CallerNumber As Integer = 0, Optional ByVal IgnoreNum As Integer = -1)
    Dim X As Integer
    Dim Y As Integer
    Dim i As Boolean
    Dim B As Boolean
    Dim Formatted As String
    '>>> Call WriteDebugLog("CALL: SendAll(" & Data & ", " & NoRepeat & ", " & CallerNumber & ")")
    Formatted = FormatPacket(Data, UseXOR)
    For X = 1 To MaxUsers
        If IsLoaded(X) And Player(X).Active Then
            If IgnoreNum <> -1 Then
                B = Not Player(X).MessageAllow(IgnoreNum)
            End If
            If Not B And (Not NoRepeat Or X <> CallerNumber) Then
                If Not IsIgnoring(X, CallerNumber) Or Left(Data, 5) <> "CMSG:" Or (InStr(1, Right(Data, Len(Data) - 5), ":") = 0 And InStr(1, Right(Data, Len(Data) - 5), "*") = 0) Then
                    For Y = 1 To 500
                        If SendAllData(X, Y).Data = "" Then
                            SendAllData(X, Y).Data = Data
                            SendAllData(X, Y).Packet = Formatted
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        i = False
    Next
 End Sub

Sub AddToQueue(ByVal Number As Integer, ByVal Data As String)
    Dim X As Integer
    Dim E As Boolean
    '>>> Call WriteDebugLog("CALL: AddToQueue(" & Number & ", " & Data & ")")
    Select Case Left(Data, 5)
    Case "REQN:", "RPWD:", "BANU:", "NOIP:"
        E = True
    Case Else
        E = False
    End Select
    If IsLoaded(Number) Then
'        If SendTimer = 0 Then
'            Call SendData(Number, Data)
'        Else
            For X = 1 To 500
                If SendAllData(Number, X).Data = "" Then
                    SendAllData(Number, X).Data = Data
                    SendAllData(Number, X).Packet = FormatPacket(Data, UseXOR And Not (E Or Player(Number).SkipXOR))
                    StatusBar1.Panels(3).Text = "Ch." & Number & " - " & X
                    Exit For
                End If
            Next
'        End If
    End If
End Sub
Sub DoIncoming(ByVal Number As Integer, ByVal Data As String)
    Dim X As Integer
    If Left(Data, 5) <> "PWDS:" And Left(Data, 5) <> "NAME:" And UseXOR Then
        Data = XORDecrypt(Data)
    End If
    If Player(Number).KickTimer <> 0 And Data <> "EXIT:" Then Exit Sub
    '>>> Call WriteDebugLog("CALL: DoIncoming(" & Number & ", " & Trim(Data) & ")")
    If Data = "PONG:" Then
        If Timer - Player(Number).PingTime > 9.999 Then
            Player(Number).Speed = "9999"
        Else
            Player(Number).Speed = Format((Timer - Player(Number).PingTime) * 1000, "0000")
        End If
        Exit Sub
    End If
    If SendTimer = 0 Then
        Call ProcessIncoming(Number, Data)
    Else
        For X = 1 To 500
            If ReceiveQueue(Number, X) = "" Then
                ReceiveQueue(Number, X) = Data
                StatusBar1.Panels(4).Text = "Ch." & Number & " - " & X
                Exit For
            End If
        Next
    End If
End Sub

Private Sub RndTimer_Timer()
    If UseTrueRnd And RndState = rEmpty Then
        RndState = rQuerying
        On Error GoTo NoInetOCX
        '>>> Call WriteDebugLog("Loading RndForm")
        RndForm.Hide
    End If
    Exit Sub
NoInetOCX:
    Call AddMessage("The True Random system requires MSINET.ocx; please visit www.netbattle.net to download it.  Switching to Pseudo Random for now.")
    UseTrueRnd = False
    RndTimer.Enabled = False
End Sub

Private Sub SaveDataTimer_Timer()
    SinceLastSave = SinceLastSave + 1
    If SinceLastSave >= 15 Then
        If TodaysDate <> Date Then
            TodaysDate = Date
            Call ServerDB.PurgeUsers(PurgeDays)
        End If
        If ServerDB.WriteDB Then SinceLastSave = 0
        If AutoLogging = 1 Then
            Call StopAutoLog
            Call StartAutoLog
        End If
    End If
End Sub

Private Sub SendAllQueue_Timer()
    Dim X As Integer
    Dim Y As Integer
    
    For X = 1 To MaxUsers
        If IsLoaded(X) Then
            If SendAllData(X, 1).Data <> "" Then
                Call SendData(X, SendAllData(X, 1))
                For Y = 1 To 499
                    SendAllData(X, Y) = SendAllData(X, Y + 1)
                Next
            End If
            If ReceiveQueue(X, 1) <> "" Then
                Call ProcessIncoming(X, ReceiveQueue(X, 1))
                For Y = 1 To 499
                    ReceiveQueue(X, Y) = ReceiveQueue(X, Y + 1)
                Next
            End If
            SendAllData(X, 500) = BlankSend
            ReceiveQueue(X, 500) = ""
        End If
    Next
End Sub

Sub StartAutoLog()
    AutoLogging = 1
    Logging = True
    mnuFileItem(1).Enabled = False
    StatusBar1.Panels(1).Text = "Server Auto-Logging enabled"
    LogFileNum = FreeFile
    LogDate = Date
    If FileExists(SlashPath & Format(Now, "mm-dd-yy") & " Log.txt") Then
        Open SlashPath & Format(Now, "mm-dd-yy") & " Log.txt" For Append As #LogFileNum
    Else
        Open SlashPath & Format(Now, "mm-dd-yy") & " Log.txt" For Output As #LogFileNum
    End If
End Sub

Sub StopAutoLog()
    Close #LogFileNum
    AutoLogging = 0
    Logging = False
    mnuFileItem(1).Enabled = True
End Sub

'Sub SendWrongData(ByVal SendMe As String)
'    Dim Temp As Integer
'
'    On Error GoTo ErrorTrap
'
'    Call AddMessage("Sent " & SendMe & " on the Wrong Mode socket.", , True)
'    If Len(SendMe) < NetChunkSize Then SendMe = SendMe & String(NetChunkSize - Len(SendMe), " ")
'    If Len(SendMe) > NetChunkSize Then SendMe = left(SendMe, NetChunkSize)
'    WrongMode.SendData SendMe
'    Exit Sub
'ErrorTrap:
'    WrongMode.Close
'End Sub

Function IsIgnoring(ByVal PNum As Integer, ByVal Ignoring As Integer) As Boolean
     Dim X As Integer
     Dim Temp As Boolean
     
     On Error Resume Next
     Temp = False
     For X = 0 To UBound(Player(PNum).Ignore) - 1
         If Player(PNum).Ignore(X) = Ignoring Then Temp = True
     Next X
     IsIgnoring = Temp
 End Function
 
 Sub StopIgnore(ByVal PNum As Integer, ByVal Ignoring As Integer)
     Dim X As Integer
     Dim Y As Integer
     
     For X = 0 To UBound(Player(PNum).Ignore) - 1
         If Player(PNum).Ignore(X) = Ignoring Then Exit For
     Next X
     For Y = X + 1 To UBound(Player(PNum).Ignore)
         Player(PNum).Ignore(Y - 1) = Player(PNum).Ignore(Y)
     Next Y
     ReDim Preserve Player(PNum).Ignore(UBound(Player(PNum).Ignore) - 1)
 End Sub
 
Function GetNumber(ByVal PName As String) As Integer
     Dim X As Integer
     Dim Temp As Integer
     
     Temp = 0
     For X = 1 To MaxUsers
         If IsLoaded(X) Then
             If LCase(Player(X).Name) = LCase(PName) Then Temp = X
         End If
     Next X
     GetNumber = Temp
End Function

Private Sub ScriptTimer_Timer()
    Dim X As Integer
    If VarChange Then
        SaveVariables
        VarChange = False
    End If
    For X = TimerLimit To UBound(sEvent)
        sEvent(X).Counter = sEvent(X).Counter + 1
        If sEvent(X).Counter = sEvent(X).Trigger Then
            sEvent(X).Counter = 0
            If ProcessScript Then BlockExec X, False
        End If
    Next X
End Sub


Private Sub tmrKickTimer_Timer()
    Dim X As Integer
    For X = 1 To MaxUsers
        If IsLoaded(X) Then
            If Player(X).KickTimer <> 0 Then
                If Player(X).KickTimer = 1 Then
                    Player(X).KickTimer = 0
                    If Player(X).DCReason = "" Then Player(X).DCReason = "Player forcibly disconnected."
                    Call DisconnectPlayer(X)
                Else
                    Player(X).KickTimer = Player(X).KickTimer - 1
                End If
            ElseIf Player(X).FloodCount <> 0 Then
                Player(X).FloodCount = Player(X).FloodCount - 1
            End If
        End If
    Next X
    
    'Meh, to avoid adding YET ANOTHER timer, I'll just stick this here...
    For X = 1 To UBound(ServerBattle)
        With ServerBattle(X)
            If .Timeout > 0 Then
                If .BattleOver Then
                    .Timeout = 0
                Else
                    .Timeout = .Timeout - 1
                    If .Timeout = 30 Then
                        Call SendAllBattle("HURRY" & CStr(.WaitingFor), X)
                    ElseIf ServerBattle(X).Timeout = 0 Then
                        Call SendAllBattle("TIME:" & CStr(.WaitingFor), X)
                        Call ServerBattle(X).ForceLoss(.WaitingFor)
                        Call BattleProcess(X, True)
                    End If
                End If
            End If
        End With
    Next X
End Sub

Private Sub UserChangeTimer_Timer()
    Dim X As Long
    Dim Y As Long
    Dim C As Long
    'Count up the users and skip the clones
    For X = 1 To MaxUsers
        If Player(X).Active Then
            For Y = 1 To X - 1
                If Player(X).SID = Player(Y).SID _
                Or Player(X).Address = Player(Y).Address Then Exit For
            Next Y
            If Y = X Then C = C + 1
        End If
    Next X
    If C <> UserCount Then
        UserCount = C
        Call SendRegData("USRC:" & CStr(UserCount))
    End If
End Sub

'Private Sub WrongMode_Close()
'    Call AddMessage("Wrong mode socket disconnected.")
'    WrongMode.Close
'    WrongMode.Listen
'End Sub

'Private Sub WrongMode_ConnectionRequest(ByVal requestID As Long)
'    Call AddMessage("Attempt to connect in the wrong mode.")
'    WrongMode.Accept requestID
'    Call SendWrongData("SMSG:This is a server, not a game.  Connect using the proper option.")
'    Call SendWrongData("MSRV")
'End Sub


Function InterpretTeamInfo(ByVal Info As String, ByVal PNum As Integer) As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    Dim T As Single
    Dim BlankPoke As ServerPKMNData
    Dim TempPKMN(1 To 6) As Pokemon
    Dim Backup As String
    't = Timer
    InterpretTeamInfo = False
    On Error GoTo BadInfo
    'Debug.Print "Pokemon Info: "; Len(Info); " characters"
    Player(PNum).ShowTeam = Bin2Bool(ChopString(Info, 1))
    Player(PNum).Rank = Asc(ChopString(Info, 1))
    X = Dec(ChopString(Info, 1))
    If X < 0 Or X > 7 Then X = 5
    Player(PNum).GFXVer = X
    Backup = Info
    For X = 1 To 6
        TempPKMN(X) = Str2PKMN(ChopString(Info, POKELEN))
        'If X = 1 And TempPKMN(1).GameVersion = nbModAdv Then ApplyDBMod
        If TempPKMN(X).No = 0 Then GoTo BadInfo
        Player(PNum).PKMN(X) = TempPKMN(X).No
        Player(PNum).PokeData(X) = BlankPoke
        With Player(PNum).PokeData(X)
            .Item = TempPKMN(X).Item
            .Level = TempPKMN(X).Level
            .Nickname = CorrectText(TempPKMN(X).Nickname)
            .DV_HP = TempPKMN(X).DV_HP
            .DV_Atk = TempPKMN(X).DV_Atk
            .DV_Def = TempPKMN(X).DV_Def
            .DV_Spd = TempPKMN(X).DV_Spd
            .DV_SAtk = TempPKMN(X).DV_SAtk
            .DV_SDef = TempPKMN(X).DV_SDef
            .EV_HP = TempPKMN(X).EV_HP
            .EV_Atk = TempPKMN(X).EV_Atk
            .EV_Def = TempPKMN(X).EV_Def
            .EV_Spd = TempPKMN(X).EV_Spd
            .EV_SAtk = TempPKMN(X).EV_SAtk
            .EV_SDef = TempPKMN(X).EV_SDef
            .AttNum = TempPKMN(X).AttNum
            .Gender = TempPKMN(X).Gender
            .NatureNum = TempPKMN(X).NatureNum
            .UnownLetter = TempPKMN(X).UnownLetter
            .Shiny = TempPKMN(X).Shiny
            Player(PNum).PKMNImage(X) = ChooseImage(TempPKMN(X), Player(PNum).GFXVer)
            For Y = 1 To 4
                '.Move(Y) = Dec(ChopString(Info, 3))
                .Move(Y) = TempPKMN(X).Move(Y)
                If .Move(Y) = 0 And Y = 1 Then
                    Call AddToQueue(PNum, "ILLM:" & BasePKMN(Player(PNum).PKMN(X)).Name & " has no moves.")
                    Exit Function
                ElseIf .Move(Y) < 0 Or .Move(Y) > UBound(Moves) Then
                    GoTo BadInfo
                End If
            Next
            Temp = LegalMove(TempPKMN(X))
            If Temp <> "" Then
                Call AddToQueue(PNum, "ILLM:" & Temp)
                Exit Function
            End If
        End With
    Next
        
    Player(PNum).GameVersion = TempPKMN(1).GameVersion
    Call ReadBinArray(CompatCheck(TempPKMN), Player(PNum).Compatibility)
    If Not Player(PNum).Compatibility(Player(PNum).GameVersion) Then GoTo BadInfo
    
    'Well, now that we've got all this nice data, let's make a string
    'for ServerBattle.SetTeam to use again and again!  This code is
    'basically the same as GetTeam.
    Player(PNum).TeamString = Backup
    Player(PNum).TeamChecksum = MD5(Backup)
    Player(PNum).StadiumOK = True
    
    InterpretTeamInfo = True
    ''Debug.Print Timer - t
    Exit Function
BadInfo:
    'The only way this could happen is if the game was hacked.
    If InVBMode Then
        Stop
        Resume
    End If
    Call AddToQueue(PNum, "ILLM:Your team is invalid.")
End Function

Sub RelayScanner(ByVal Sender As Long, ByVal Receiver As Long, ByVal RelayString As String, ByVal BattleID As Integer)
    Dim DontSend As Boolean
    Dim Command As String * 5
    Dim Data As String
    Dim X As Integer
    Dim Y As Byte
    Dim Z As Byte
    Dim Temp As Boolean
    Dim Team As Byte
    Dim Opponent As Byte
    Dim GoAhead As Boolean
    Dim Watcher As Boolean
    Dim SendToAll As Boolean
    Dim Watch() As String
    Dim TempString As String
    Dim TempEffect(1 To 12) As Byte
    
    Temp = True
    With ServerBattle(BattleID)
        'Check for ghost battles
        If Player(.Player1).BattlingWith <> .Player2 Or Player(.Player2).BattlingWith <> .Player1 Then
            If Not .DisconStall Then Exit Sub
        End If
        
        'Which player is sending?
        If Sender = .Player1 Then
            Team = 1
            Opponent = 2
        ElseIf Sender = .Player2 Then
            Team = 2
            Opponent = 1
        Else
            Watcher = True
        End If
        Call .ClearSelfLoss
        DontSend = False
        SendToAll = False
        Data = RelayString
        Command = ChopString(Data, 5)
        Select Case Command
            'Player's Battle window has loaded.
            Case "READY"
                'Don't pass this on.
                DontSend = True
                If .DisconStall Then
                    If Sender = .Player1 Then
                        TempString = "00000000000001" & Chr$(.Player2) & .GetTeam(2)
                    Else
                        TempString = "00000000000002" & Chr$(.Player1) & .GetTeam(1)
                    End If
                    Call AddToQueue(Sender, "RELAYINFO:1" & FixedHex(.BattleMode, 1) & FixedHex(.ActNum, 1) & FixedHex(.Terrain, 1) & .Rules & TempString)
                    TempString = .SavedSyncStr
                    Do Until Len(TempString) <= 230
                        Call AddToQueue(Sender, "RELAYSYNC:" & ChopString(TempString, 230) & vbNullChar)
                    Loop
                    Call AddToQueue(Sender, "RELAYSYNC:" & TempString)
                    .SavedSyncStr = vbNullString
                    Call AddToQueue(Receiver, "RELAYBACK:")
                    If Sender = .Player1 Then
                        TempString = "BACK:1"
                    Else
                        TempString = "BACK:2"
                    End If
                    Watch = Split(.GetWatchers, ";")
                    For X = 0 To UBound(Watch)
                        If Val(Watch(X)) <> Sender Then
                            Call AddToQueue(Val(Watch(X)), "WATCH" & FixedHex(BattleID, 3) & TempString)
                        End If
                    Next X
                    If .SendLastBStr Then
                        Call AddToQueue(Sender, "RELAY" & .BattleString)
                        .SendLastBStr = False
                    End If
                    .DisconStall = False
                    Exit Sub
                End If
                .PlayerName(Team) = Player(Sender).Name
                'Check in with BattleData object
                GoAhead = .IsReady(Team)
                'GoAhead will be True if both players have checked in - send initial info.
                If GoAhead Then
                    'Raise the BeginBattle event for scripts
                    Call ScriptMod.BlockExec(10, , .Player1, .Player2, StrReverse$(Dec2Bin(Dec(.Rules), UBound(RuleText))))
                    If .RandBat Then
                        'Make the random teams if Challenge Cup Mode
                        'Note to self: Change this!
                        TempString = MakeChallengeTeam(.BattleMode)
                        GoAhead = .SetTeam(1, TempString)
                        Call AddToQueue(.Player1, "RELAYRAND:" & TempString)
                        TempString = MakeChallengeTeam(.BattleMode)
                        GoAhead = .SetTeam(2, TempString)
                        Call AddToQueue(.Player2, "RELAYRAND:" & TempString)
                    Else
                        'Otherwise, since the server has the info already, set the teams now.
                        GoAhead = .SetTeam(1, Player(.Player1).TeamString)
                        GoAhead = .SetTeam(2, Player(.Player2).TeamString)
                    End If
                    
                    'Set and send the berry effects
                    X = GetSmallRnd(15)
                    Call .SetTrace(X)
                    TempString = Hex(X)
                    For X = 1 To 12
                        TempEffect(X) = .GetItemEffect
                    Next X
                    Call .SetItemEffect(TempEffect)
                    For X = 1 To 12
                        TempString = TempString & Chr$(TempEffect(X))
                    Next X
                    
                    'Now send the data.
                    Call AddToQueue(.Player1, "RELAYINFO:0" & FixedHex(.BattleMode, 1) & FixedHex(.ActNum, 1) & FixedHex(.Terrain, 1) & .Rules & TempString & "1" & Chr$(.Player2) & .GetTeam(2))
                    Call AddToQueue(.Player2, "RELAYINFO:0" & FixedHex(.BattleMode, 1) & FixedHex(.ActNum, 1) & FixedHex(.Terrain, 1) & .Rules & TempString & "2" & Chr$(.Player1) & .GetTeam(1))
                    If Not .StadiumMode Then
                        .StartBattle
                        .WatchOK = True
                    End If
                End If
            'Stadium Pokemon
            Case "SPKM:"
                If .NeedsStadiumSelect(Team) Then
                    GoAhead = .SetSPoke(Team, Data)
                    If Not .NeedsStadiumSelect(Opponent) Then
                        GoAhead = .DoThreePKMN
                        Call .StartBattle
                        .WatchOK = True
                    End If
                Else
                    Temp = False
                End If
            Case "CANCL"
                DontSend = True
                For X = Team To .ActNum Step 2
                    If .IsCancellable(X) Then
                        .UnloadMove X
                        .UnloadSwitch X
                        .IsCancellable(X) = False
                        .Timeout = 0
                        Call AddMessage("Move cancelled.")
                        Call AddToQueue(Sender, "RELAYCANCL")
                    End If
                Next X
                
            Case "CMSG:"
                Data = WordFilter(Data)
                If Left$(Data, 4) = "/me " Then
                    RelayString = Command & "*** " & Player(Sender).Name & " " & Right$(Data, Len(Data) - 4)
                Else
                    RelayString = Command & Player(Sender).Name & ": " & Data
                End If
                SendToAll = True
            Case "WMSG:"
                If .NoWatchChat Then Exit Sub
                Data = WordFilter(Data)
                If Left$(Data, 4) = "/me " Then
                    RelayString = Command & "*** " & Player(Sender).Name & " " & Right$(Data, Len(Data) - 4)
                Else
                    RelayString = Command & Player(Sender).Name & ": " & Data
                End If
                SendToAll = True
            Case "UNACC"
                .Unrated = True
            Case "TIACC"
                .BattleOver = True
                Call BattleProcess(BattleID, True)
            Case "MOVE:"
                DontSend = True
                If Len(Data) <> 2 Then
                    Temp = False
                Else
                    Data = Chr2Bin(Data)
                    Call .IsReady(Team)
                    For X = Team To 4 Step 2
                        TempString = ChopString(Data, 6)
                        If Not .Ready(X) Then
                            .IsCancellable(X) = True
                            If ChopString(TempString, 1) = "1" Then
                                Y = Bin2Dec(ChopString(TempString, 3))
                                Z = Bin2Dec(ChopString(TempString, 2))
                                If Z >= X Then Z = Z + 1
                                Temp = .LoadMove(X, Y, Z)
                                If Not Temp Then Exit For
                            Else
                                Y = Bin2Dec(ChopString(TempString, 5))
                                Temp = .LoadSwitch(X, Y)
                                If Not Temp Then Exit For
                            End If
                        End If
                    Next X
                End If
                If Temp Then Call BattleProcess(BattleID)
'            Case "PKSW:"
'                DontSend = True
'                Temp = .LoadSwitch(Team, Val(Data))
'                If Temp Then Call BattleProcess(BattleID)
            Case "RCVB:" 'Confirms that the client received the BattleString.
                'BattleProcess must be called here in case both players
                'have a StuckMove and would be unable to send anything else,
                'in which case the battle would lock up.
                '(Not used anymore)
                DontSend = True
                If Team = 1 Then .Waiting1 = False Else .Waiting2 = False
                If Not (.Waiting1 Or .Waiting2) Then Call BattleProcess(BattleID)
            Case "IGWAT"
                SendToAll = True
                RelayString = Command & Team
                Call AddToQueue(Sender, "RELAY" & RelayString)
                If Team = 1 Then
                    .WatchIgnore1 = Not .WatchIgnore1
                Else
                    .WatchIgnore2 = Not .WatchIgnore2
                End If
        End Select
        If Not Temp Then
            'Buh-bye.
            If InVBMode Then Stop
            DontSend = False
            Watcher = True
            RelayString = "HACK:" & CStr(Team)
            Call .ForceLoss(Team)
            Call BattleProcess(BattleID, True)
        End If
        
        If RelayString <> "READY" And .Timeout = 0 And .UseTimeout Then
            If Not .BattleOver Then
                For X = 1 To 2
                    If .Ready(X) And Not .BattleOver Then
                        'Debug.Print "Starting Timeout"
                        .Timeout = 300
                        .WaitingFor = OtherTeam(X)
                    End If
                Next X
            End If
        End If
        
        If Not DontSend Then
            If Watcher Then
                If Command <> "WMSG:" Or Not .WatchIgnore1 Then Call AddToQueue(.Player1, "RELAY" & RelayString)
                If Command <> "WMSG:" Or Not .WatchIgnore2 Then Call AddToQueue(.Player2, "RELAY" & RelayString)
                Call AddMessage(RelayString & " was sent to battle #" & BattleID, True)
            Else
                Call AddToQueue(Receiver, "RELAY" & RelayString)
                Call AddMessage(RelayString & " was sent from " & Player(Sender).Name & " to " & Player(Receiver).Name, True)
            End If
            If SendToAll Then
                Watch = Split(.GetWatchers, ";")
                For X = 0 To UBound(Watch)
                    If Val(Watch(X)) <> Sender Then
                        Call AddToQueue(Val(Watch(X)), "WATCH" & FixedHex(BattleID, 3) & RelayString)
                    End If
                Next X
            End If
        End If
    End With
End Sub

Sub BattleProcess(ByVal BattleID As Integer, Optional WinCheckOnly As Boolean = False)
    Dim Temp As String
    Dim X As Long
    Dim Watch() As String
    With ServerBattle(BattleID)
        If Not WinCheckOnly Then
            If Not .Ready Then Exit Sub
            For X = 1 To 4
                .IsCancellable(X) = False
            Next X
            .Timeout = 0
            .WaitingFor = 0
            Call .DoBattle
            Call AddToQueue(.Player1, "RELAY" & .BattleString)
            Call AddToQueue(.Player2, "RELAY" & .BattleString)
            .Waiting1 = True
            .Waiting2 = True
            Watch = Split(.GetWatchers, ";")
            For X = 0 To UBound(Watch)
                Call AddToQueue(Val(Watch(X)), "WATCH" & FixedHex(BattleID, 3) & .BattleString)
            Next X
        End If
        If .BattleOver Then
            .Timeout = 0
            .WaitingFor = 0
            Select Case .Winner
                Case 1
                    Call ScriptMod.BlockExec(7, , .Player1, .Player2, "WIN" & IIf(.Unrated, "*", ""))
                    Call SendAll("CMSG:" & .PlayerName(1) & " has beaten " & .PlayerName(2), , , 4)
                    Call AddMessage(.PlayerName(1) & " has beaten " & .PlayerName(2))
                    If Not .Unrated Then
                        Call ChangePlayerStats(.PlayerName(1), 1)
                        Call ChangePlayerStats(.PlayerName(2), 2)
                    End If
                    Call ScriptMod.BlockExec(8, , .Player1, .Player2, "WIN" & IIf(.Unrated, "*", ""))
                Case 2
                    Call ScriptMod.BlockExec(7, , .Player2, .Player1, "WIN" & IIf(.Unrated, "*", ""))
                    Call SendAll("CMSG:" & .PlayerName(2) & " has beaten " & .PlayerName(1), , , 4)
                    Call AddMessage(.PlayerName(2) & " has beaten " & .PlayerName(1))
                    If Not .Unrated Then
                        Call ChangePlayerStats(.PlayerName(1), 2)
                        Call ChangePlayerStats(.PlayerName(2), 1)
                    End If
                    Call ScriptMod.BlockExec(8, , .Player2, .Player1, "WIN" & IIf(.Unrated, "*", ""))
                Case 3
                    Call ScriptMod.BlockExec(7, , .Player1, .Player2, "TIE" & IIf(.Unrated, "*", ""))
                    Call SendAll("CMSG:" & .PlayerName(1) & " and " & .PlayerName(2) & " have tied.", , , 4)
                    Call AddMessage(.PlayerName(1) & " and " & .PlayerName(2) & " have tied")
                    If Not .Unrated Then
                        Call ChangePlayerStats(.PlayerName(1), 3)
                        Call ChangePlayerStats(.PlayerName(2), 3)
                    End If
                    Call ScriptMod.BlockExec(8, , .Player1, .Player2, "TIE" & IIf(.Unrated, "*", ""))
            End Select
        End If
    End With
End Sub

Function MakeList() As String
    Dim TempString As String
    Dim X As Integer
    
    TempString = ""
    For X = 1 To MaxUsers
        If Player(X).Name <> "" And IsLoaded(X) Then
            TempString = TempString & Chr$(X)
        End If
    Next
    TempString = "PLST:" & TempString
    MakeList = TempString
End Function

Private Sub RegSocket_Close()
    If (ConnectedToReg Or ClientSocket(0).State <> sckListening) And PublicServer Then
        Call AddMessage("Lost connection to registry; will retry at 60 second intervals.")
        ConnectedToReg = False
    Else
        Call AddMessage("Registry connection closed.")
    End If
    RegSocket.Close
    If ClientSocket(0).State <> sckListening Then Call ServerStartup
End Sub

Private Sub RegSocket_Connect()
    RegTimer.Enabled = False
    Call AddMessage("Connection to registry established!  Exchanging data...")
End Sub

Private Sub RegSocket_DataArrival(ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(RegSocket, BytesTotal, Packet)
    If Worked Then
        StatusBar1.Panels(2).Text = BytesTotal & " Reg bytes rcd."
        For X = 1 To UBound(Packet)
            Call DoRegIncoming(Packet(X))
        Next X
    Else
        Call AddMessage("Registry Data Arrival error.")
        Call SendRegData("EXIT:")
    End If
End Sub
Public Sub SendRegData(ByVal SendMe As String)
    Dim T As String
    T = Left(SendMe, 5)
    If RegSocket.State <> sckConnected Then Exit Sub
    If T <> "INFO:" And T <> "RLIP:" And T <> "PONG:" And T <> "PASS:" And T <> "EXIT:" And Not ConnectedToReg Then Exit Sub
    RegSocket.SendData FormatPacket(SendMe, True)
End Sub

Private Sub DoRegIncoming(ByVal iData As String)
    Dim Temp As String
    Dim Pref As String
    Temp = XORDecrypt(iData)
    'Debug.Print Temp
    Pref = ChopString(Temp, 5)
    Select Case Pref
        Case "RINF:" 'Request INFo
            If ServerName = "" Or Admin = "" Or ServerDesc = "" Then
                Call AddMessage("Insuficient information; please set registry info in the options screen.")
                PublicServer = False
                Call SendRegData("EXIT:")
                If ClientSocket(0).State <> sckListening Then Call ServerStartup
            End If
            Temp = "INFO:" & Pad(ServerName, 20) & Pad(Admin, 20) & Chr$(ListView1.ListItems.count) & Chr$(MaxUsers) & Pad(You.ProgVersion, 8) & StationID & ServerDesc
            Call SendRegData(Temp)
        Case "OKAY!" 'You're all set
            Call AddMessage("Registry log-on completed.  " & ServerName & " is now officially listed.")
            ConnectedToReg = True
            If ClientSocket(0).State <> sckListening Then Call ServerStartup
            If Len(Temp) > 0 Then SNRegged = True
        Case "RQIP:" 'ReQuest IP
            If RealIP = "" Then
                Temp = InputBox("Your IP has been determined to be a private IP.  This may be because this server is being run on the same computer as the registry, or it may be a bug in the program.  Please enter your real IP Address here (Must be an IP, not a DNS Address or redirect.)  If you do not know your IP Address, please visit www.whatismyip.com to find out.", "Invalid IP")
                If Temp = "" Then
                    Temp = "EXIT:"
                Else
                    SaveSetting "NetBattle", "Master Server", "Real IP", Temp
                    RealIP = Temp
                    Temp = "RLIP:" & Temp
                End If
            Else
                Temp = "RLIP:" & RealIP
            End If
            Call SendRegData(Temp)
        Case "NMIU:" 'NaMe In Use
            Temp = InputBox("The name " & ServerName & " is already in use.  Please choose another.  If you think someone is impersonating you, please email TV's Ian at netbattle@tvsian.com.")
            If Temp = "" Or Temp = ServerName Then
                PublicServer = False
                Call SendRegData("EXIT:")
            Else
                ServerName = Temp
                SaveSetting "NetBattle", "Master Server", "Name", Temp
                Call SendRegData("INFO:" & Pad(ServerName, 20) & Pad(Admin, 20) & FixedHex(ListView1.ListItems.count, 2) & FixedHex(MaxUsers, 2) & ServerDesc)
            End If
        Case "MASS:"
            Call AddMessage("NETWORK-WIDE MESSAGE: " & Temp)
            Call SendAll("MASS:" & Temp)
        Case "TBAN:"
            Call SendRegData("EXIT:")
            Select Case Temp
                Case "0"
                    Temp = "You have been temporarily banned for PONG flooding.  Have a nice day."
                Case "1"
                    Temp = "The server registry has recorded 15 connection attempts from your IP in less than 60 seconds.  For security reasons, you have been temporarily banned from the Registry for 15 minutes.  Please try to space out your connections a bit next time."
                Case "2"
                    Temp = "The server registry has recorded 15 changes to your information in less than 60 seconds.  For security reasons, you have been temporarily banned from the Registry for 15 minutes.  Please try to space out your changes a bit next time."
            End Select
            Call AddMessage(Temp)
        Case "MULTI"
            Call SendRegData("EXIT:")
            Call AddMessage("The Registry has reported that that a connection from this IP already exists.  Only one server per address can be listed.")
        Case "WARN:"
            Call AddMessage("The server registry has recorded 10 changes to your information in less than 60 seconds.  Please wait a while before changing it again or you will be disconnected.")
        Case "OLDVR"
            Call SendRegData("EXIT:")
            Call AddMessage("Your server version is outdated.  You must have the lastest NetBattle version to connect to the Server Registry.")
        Case "PING:"
            Call SendRegData("PONG:")
        Case "REGED"
            Call AddMessage(ServerName & " has been sucessfully registered.")
            SetUsers.lblRegged.Caption = "Registered"
            SNRegged = True
        Case "RIGHT"
            Call AddMessage(ServerName & " has been reregistered.")
        Case "WRONG"
            Call SendRegData("EXIT:")
            Call AddMessage("Password incorrect.")
        Case "NAMPW"
            Temp = InputBox("Your SID does not match the one registered to this Server Name.  If the Name is yours, please enter your password to reregister it.")
            If Len(Temp) = 0 Then
                Call SendRegData("EXIT:")
            Else
                Call SendRegData("PASS:" & MD5(Temp))
            End If
        Case Else
            Call AddMessage("Unknown registry command: " & Temp)
    End Select
End Sub

Private Sub RegSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    StatusBar1.Panels(0).Text = "Registry Error:" & Description
    'Debug.Print Description
    RegTimer.Enabled = False
    If ConnectedToReg Then
        Call AddMessage("Lost connection to registry; will retry at 60 seconds intervals.")
    End If
    ConnectedToReg = False
    RegSocket.Close
    If ClientSocket(0).State <> sckListening Then
        Call AddMessage("Error connnecting to registry; will retry at 60 second intervals.")
        Call ServerStartup
    End If
End Sub

Private Sub RegTimer_Timer()
    If Not PublicServer Then
        RegSocket.Close
        RegTimer.Enabled = False
        Exit Sub
    End If
    If RegSocket.State <> sckConnected Then
        If ClientSocket(0).State <> sckListening Then
            Call AddMessage("Registry connection timed out; will continue to retry at 60 second intervals.")
            Call ServerStartup
        End If
        RegSocket.Close
    End If
    RegTimer.Enabled = False
End Sub

Public Sub Tempban(ByVal Index As Integer)
    Dim X As Integer
    For X = 1 To UBound(TempBanInfo)
        If TempBanInfo(X).TimeLeft <> 0 Then Exit For
    Next X
    If X = UBound(TempBanInfo) + 1 Then ReDim Preserve TempBanInfo(X)
    With TempBanInfo(X)
        '.IP = Player(Index).Address
        '.SID = Player(Index).SID
        .PName = Player(Index).Name
        .TimeLeft = 15
    End With
    Call AddToQueue(Index, "TBAN:15")
End Sub
Public Function IsTempBanned(ByVal Index As Long, rMinutes As Long) As Boolean
    Dim X As Integer
    IsTempBanned = 0
    For X = 1 To UBound(TempBanInfo)
        With TempBanInfo(X)
            If .TimeLeft <> 0 Then
                If .SID = Player(Index).SID Or .PName = Player(Index).Name Then
                    rMinutes = .TimeLeft
                    IsTempBanned = True
                End If
            End If
        End With
    Next X
End Function
Public Function SIDIsTempBanned(ByVal SID As String, ByRef rMinutes As Long) As Boolean
    Dim X As Integer
    SIDIsTempBanned = False
    For X = 1 To UBound(TempBanInfo)
        With TempBanInfo(X)
            If .TimeLeft <> 0 Then
                If .SID = SID Then
                    SIDIsTempBanned = True
                    rMinutes = .TimeLeft
                    Exit Function
                End If
            End If
        End With
    Next X
End Function
Public Function SlashCommand(ByVal Index As Integer, ByRef iData As String) As Boolean
    Dim Command As String
    Dim X As Integer
    iData = Right(iData, Len(iData) - 1)
    Command = Trim(ChopString(iData, InStr(1, iData, " ")))
    SlashCommand = True
    Select Case LCase(Command)
        Case "me"
            If Command = "ME" Then
                iData = "*** " & UCase(Player(Index).Name) & " " & iData
            Else
                iData = "*** " & Player(Index).Name & " " & iData
            End If
        Case "ignore"
            X = GetNumber(iData)
            iData = ""
            If X = 0 Then
                Call AddToQueue(Index, "CMSG:No such player.")
                Exit Function
            ElseIf IsIgnoring(Index, X) Then
                Call AddToQueue(Index, "CMSG:You are already ignoring this player.")
                Exit Function
            ElseIf X = Index Then
                Call AddToQueue(Index, "CMSG:Cannot ignore self!")
                Exit Function
            Else
                Player(Index).Ignore(UBound(Player(Index).Ignore)) = X
                ReDim Preserve Player(Index).Ignore(UBound(Player(Index).Ignore) + 1)
                iData = ""
                Call SendAll(Player(Index).Name & " is ignoring " & Player(X).Name, , , 8)
            End If
        Case "unignore"
            X = GetNumber(iData)
            iData = ""
            If X = 0 Then
                Call AddToQueue(Index, "CMSG:No such player.")
                Exit Function
            ElseIf Not IsIgnoring(Index, X) Then
                Call AddToQueue(Index, "CMSG:Not ignoring player.")
                Exit Function
            Else
                Call StopIgnore(Index, X)
                iData = Player(Index).Name & " has stopped ignoring " & Player(X).Name
            End If
        Case Else
            iData = ""
            SlashCommand = False
    End Select
    If iData <> "" Then iData = "CMSG:" & iData
End Function
Public Sub SendAllBattle(ByVal Data As String, ByVal BattleID As Long)
    Dim Watch() As String
    Dim X As Integer
    With ServerBattle(BattleID)
        If Player(.Player1).BattleID = BattleID Then Call AddToQueue(.Player1, "RELAY" & Data)
        If Player(.Player2).BattleID = BattleID Then Call AddToQueue(.Player2, "RELAY" & Data)
        Watch = Split(.GetWatchers, ";")
        For X = 0 To UBound(Watch)
            Call AddToQueue(Val(Watch(X)), "WATCH" & FixedHex(BattleID, 3) & Data)
        Next X
    End With
End Sub
Public Function BattleOver(ByVal Index As Long, Optional ClearIt As Boolean = False) As Boolean
    Dim X As Byte
    If Player(Index).BattleID = 0 Then
        BattleOver = True
    ElseIf ServerBattle(Player(Index).BattleID).BattleOver Then
        BattleOver = True
    Else
        BattleOver = False
        If ClearIt Then
            With ServerBattle(Player(Index).BattleID)
                Debug.Print "Clearing battle!"
                If .Player1 = Index Then X = 1 Else X = 2
                Call .ForceLoss(X)
            End With
        End If
    End If
End Function

Private Sub SendDBMod(ByVal Index As Long)
    Dim X As Long
    Dim Build As String
    If Len(DBModStr) = 0 Then Exit Sub
    Build = DBModStr
    Do Until Len(Build) < 200
        AddToQueue Index, "DBMD:" & ChopString(Build, 200) & vbNullChar
    Loop
    AddToQueue Index, "DBMD:" & Build
    Player(Index).ModHash = DBModHash
End Sub
Public Sub SetDNSInfo(hMem As Long, MemLen As Long)
    Dim X As Long
    Dim Y As Long
    CopyMemory X, ByVal hMem, 4
    CopyMemory Y, ByVal hMem + 4, 4
    'Debug.Print X & " - " & Y
    With Player(X)
        If inet_addr(.Address) = Y And .DNSAddress = "[DNS Pending]" Then
            .DNSAddress = vbNullString
            MemLen = MemLen - 8
            If MemLen > 0 Then
                .DNSAddress = String$(MemLen, vbNullChar)
                CopyMemory ByVal .DNSAddress, ByVal hMem + 8, MemLen
            End If
        End If
    End With
    GlobalFree hMem
    X = UBound(DNSQueue) - X
    If X >= 1 Then
        CopyMemory DNSQueue(1), DNSQueue(2), X * Len(DNSQueue(0))
        ReDim Preserve DNSQueue(X)
        SetDNSIP DNSQueue(1).IP, DNSQueue(1).PNum
        Call CreateThread(ByVal 0&, ByVal 0&, DNSPtr, ByVal 0&, 0, ByVal 0)
    Else
        ReDim DNSQueue(0)
    End If
End Sub


