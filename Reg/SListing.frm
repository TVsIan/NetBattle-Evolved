VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MSListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetBattle Server List"
   ClientHeight    =   4335
   ClientLeft      =   180
   ClientTop       =   555
   ClientWidth     =   7470
   Icon            =   "SListing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer KickTimer 
      Interval        =   1000
      Left            =   4680
      Top             =   0
   End
   Begin VB.CommandButton cmdMassMsg 
      Caption         =   "Send Mass Message"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Timer ChannelScanner 
      Interval        =   60000
      Left            =   5640
      Top             =   0
   End
   Begin VB.Timer QueueTimer 
      Interval        =   5
      Left            =   5160
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ServerSocket 
      Index           =   0
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   30001
   End
   Begin MSWinsockLib.Winsock ClientSocket 
      Index           =   0
      Left            =   6120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   30002
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect Server"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListDisplay 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address / IP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Main Admin"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Users/Max"
         Object.Width           =   1746
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4065
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7990
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   5205
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1035
      Left            =   165
      Top             =   2955
      Width           =   5340
   End
End
Attribute VB_Name = "MSListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const VERSION = "0.9.4"
Const MAXSERVERS As Integer = 50
Const MAXCLIENTS As Integer = 25
Private Type PingType
    Chances As Integer
    Pongs As Integer
    SentPing As Boolean
End Type
Private Type QueueType2
    iData() As String
End Type
Private Type QueueType1
    Num() As QueueType2
End Type
Private Type ServerType
    Address As String
    Admin As String
    ServerName As String
    Users As Integer
    MaxUsers As Integer
    Active As Boolean
    Description As String
    SentPing As Boolean
    Pongs As Integer
    Info As String
    IPChangeable As Boolean
    ClientsKnow As Boolean
    InfoChanges As Integer
    UserChanges As Integer
    DisconnectMe As Integer
    SID As String
    Regged As Boolean
End Type
Private Type ServerDBEntry
    ServerName As String
    SID As String
    Pass As String
End Type
    
Private Queue(1 To 4) As QueueType1
'Usage:
'1=SvrSnd, 2=SvrRcv, 3=ClntSnd, 4=ClntRcv
'Queue(1).Num(3).iData(1) stores the
'next thing to be sent to Server #3.
'Queue(4).Num(1).iData(2) is the second
'thing in line to be recieved from Client #1.
Private SDiscon(MAXSERVERS) As Boolean
Private CDiscon(MAXCLIENTS) As Boolean
Private Server(MAXSERVERS) As ServerType
Private CPing(MAXCLIENTS) As PingType
Private IPBan() As String
Private AttemptLog() As String
Private SlashPath As String
Private Entry() As ServerDBEntry


Private Sub CSendAll(ByVal SendMe As String)
    Dim X As Integer
    For X = 1 To ClientSocket.UBound
        If ClientSocket(X).State = sckConnected Then
            Call AddToQueue(3, X, SendMe)
        End If
    Next X
End Sub

Private Sub ChannelScanner_Timer()
    Dim X As Integer
    Dim Temp As String
    On Error Resume Next
    For X = 1 To MAXSERVERS
        If ServerSocket(X).State = sckConnected Then
            If Server(X).SentPing Then Call DisconnectPlayer(True, X)
            Call AddToQueue(1, X, "PING:")
            Server(X).SentPing = True
            Server(X).Pongs = 0
            Server(X).InfoChanges = 0
            Server(X).UserChanges = 0
        Else
            If Server(X).Active Then Call DisconnectPlayer(True, X)
        End If
    Next X
    For X = 1 To MAXCLIENTS
        If ClientSocket(X).State = sckConnected Then
            If CPing(X).SentPing Then CPing(X).Chances = CPing(X).Chances + 1
            If CPing(X).Chances = 3 Then Call DisconnectPlayer(False, X)
            Call AddToQueue(3, X, "PING:")
            CPing(X).SentPing = True
            CPing(X).Pongs = 0
        End If
    Next X
    ReDim AttemptLog(0)
    For X = 1 To UBound(IPBan)
        Temp = Mid(IPBan(X), 1, 1)
        If Temp = "F" Then
            IPBan(X) = ""
        Else
            Mid(IPBan(X), 1, 1) = Hex(Val("&H" & Temp) + 1)
        End If
    Next X
    For X = UBound(IPBan) To 1 Step -1
        If IPBan(X) = "" Then ReDim Preserve IPBan(X - 1) Else Exit For
    Next X
    
    On Error GoTo FileOpen
    Kill SlashPath & "servers.csv"
    Open SlashPath & "servers.csv" For Output As #1
    For X = 1 To UBound(Entry)
        With Entry(X)
            Write #1, .ServerName, .SID, .Pass
        End With
    Next X
    Close #1
FileOpen:
End Sub

Private Sub ClientSocket_Close(Index As Integer)
    Call DisconnectPlayer(False, Index)
End Sub

Private Sub ClientSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    On Error Resume Next
    For X = 1 To ClientSocket.UBound
        If ClientSocket(0).RemoteHostIP = ClientSocket(X).RemoteHostIP And ClientSocket(X).State = sckConnected Then Temp = "Repeat"
    Next X
    If Temp = "" Then
        Temp = ShortenIP(ClientSocket(0).RemoteHostIP)
        For X = 1 To UBound(IPBan)
            If Mid(IPBan(X), 2, 8) = Temp Then Exit Sub
        Next X
        For X = 1 To UBound(AttemptLog)
            If Mid(AttemptLog(X), 2, 8) = Temp Then
                If Mid(AttemptLog(X), 1, 1) = "F" Then
                    Call TempBan(Temp)
                    Temp = "Ban"
                Else
                    Mid(AttemptLog(X), 1, 1) = Hex(Val("&H" & Mid(AttemptLog(X), 1, 1)) + 1)
                End If
                Exit For
            End If
        Next X
    End If
    If X = UBound(AttemptLog) + 1 Then
        ReDim Preserve AttemptLog(X)
        AttemptLog(X) = "0" & Temp
    End If
    For X = 1 To MAXCLIENTS
        If ClientSocket(X).State <> sckConnected Then Exit For
    Next X
    If X = MAXCLIENTS + 1 Then Exit Sub
    CPing(X) = CPing(0)
    CDiscon(X) = False
    ReDim Queue(3).Num(X).iData(0)
    ReDim Queue(4).Num(X).iData(0)
    If Temp = "Ban" Then
        Call TempBan(ShortenIP(ClientSocket(X).RemoteHostIP))
    ElseIf Temp = "Repeat" Then
        Exit Sub
    Else
        ClientSocket(X).Close
        ClientSocket(X).Accept requestID
        For Y = 1 To MAXSERVERS
            If Server(Y).Active Then Call AddToQueue(3, X, "SERV:" & Server(Y).Info)
        Next Y
    End If
    Call UpdateClientCount
End Sub


Private Sub ClientSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DisconnectPlayer(False, Index)
End Sub

Private Sub cmdDisconnect_Click()
    Dim Temp As String
    If ListDisplay.ListItems.Count = 0 Then Exit Sub
    Temp = ListDisplay.SelectedItem.Key
    Call DisconnectPlayer(True, Val(Right(Temp, Len(Temp) - 5)))
End Sub

Private Sub cmdMassMsg_Click()
    Dim Temp As String
    Dim X As Integer
    Temp = InputBox("Please enter a message.  This message will be sent to all players on all connected servers.", "Mass Message")
    If Temp = "" Then Exit Sub
    For X = 1 To ServerSocket.UBound
        If ServerSocket(X).State = sckConnected Then
            Call AddToQueue(1, X, "MASS:" & Temp)
        End If
    Next X
End Sub

Private Sub KickTimer_Timer()
    Dim X As Integer
    For X = 1 To MAXSERVERS
        If Server(X).DisconnectMe <> 0 Then
            Server(X).DisconnectMe = Server(X).DisconnectMe - 1
            If Server(X).DisconnectMe = 0 Then Call DisconnectPlayer(True, X)
        End If
    Next X
End Sub

Private Sub ListDisplay_Click()
    Dim X As Integer
    On Error Resume Next
    X = Val(Right(ListDisplay.SelectedItem.Key, Len(ListDisplay.SelectedItem.Key) - 5))
    Label1.Caption = Server(X).Description
End Sub

Private Sub ListDisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label1.Caption = Server(Val(Right(Item.Key, Len(Item.Key) - 5))).Description
End Sub

Private Sub ServerSocket_Close(Index As Integer)
    Call DisconnectPlayer(True, Index)
End Sub

Private Sub ServerSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim X As Integer
    Dim Y As Integer
    Dim Temp As String
    On Error Resume Next
    For X = 1 To MAXSERVERS
        If ServerSocket(0).RemoteHostIP = Server(X).Address Then Temp = "Repeat"
    Next X
    If Temp = "" Then
        Temp = ShortenIP(ServerSocket(0).RemoteHostIP)
        For X = 1 To UBound(IPBan)
            If IPBan(X) = Temp Then Exit Sub
        Next X
        For X = 1 To UBound(AttemptLog)
            If Mid(AttemptLog(X), 2, 8) = Temp Then
                If Mid(AttemptLog(X), 1, 1) = "F" Then
                    Call TempBan(Temp)
                    Temp = "Ban"
                Else
                    Mid(AttemptLog(X), 1, 1) = Hex(Val("&H" & Mid(AttemptLog(X), 1, 1)) + 1)
                End If
            End If
        Next X
    End If
    If X = UBound(AttemptLog) + 1 Then
        ReDim Preserve AttemptLog(X)
        AttemptLog(X) = "0" & Temp
    End If
    For X = 1 To MAXSERVERS
        If ServerSocket(X).State <> sckConnected Then Exit For
    Next X
    If X = MAXSERVERS + 1 Then Exit Sub
    Server(X) = Server(0)
    ReDim Queue(1).Num(X).iData(0)
    ReDim Queue(2).Num(X).iData(0)
    ServerSocket(X).Close
    ServerSocket(X).Accept requestID
    If Temp = "Ban" Then
        Call TempBan(ShortenIP(ServerSocket(0).RemoteHostIP))
        Call AddToQueue(1, X, "TBAN:2")
    ElseIf Temp = "Repeat" Then
        Call AddToQueue(1, X, "MULTI")
    Else
        Server(X).Active = True
        Call AddToQueue(1, X, "RINF:")
    End If
End Sub

Private Sub ServerSocket_DataArrival(Index As Integer, ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(ServerSocket(Index), BytesTotal, Packet)
    If Worked Then
        For X = 1 To UBound(Packet)
            Call AddToQueue(2, Index, Packet(X))
        Next X
    Else
        Call DisconnectPlayer(True, Index)
    End If
End Sub
Private Sub ClientSocket_DataArrival(Index As Integer, ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(ClientSocket(Index), BytesTotal, Packet)
    If Worked Then
        For X = 0 To UBound(Packet)
            Call AddToQueue(4, Index, Packet(X))
        Next X
    Else
        Call DisconnectPlayer(False, Index)
    End If
End Sub


Private Sub Form_Load()
    Dim TempVar As String
    Dim X As Integer
    Dim Y As Integer
    ReDim IPBan(0)
    ReDim AttemptLog(0)
    For X = 1 To 2
        ReDim Queue(X).Num(MAXSERVERS)
        For Y = 0 To MAXSERVERS
            ReDim Queue(X).Num(Y).iData(0)
        Next Y
    Next X
    For X = 3 To 4
        ReDim Queue(X).Num(MAXCLIENTS)
        For Y = 0 To MAXCLIENTS
            ReDim Queue(X).Num(Y).iData(0)
        Next Y
    Next X
    For X = 1 To MAXSERVERS
        Load ServerSocket(X)
    Next X
    For X = 1 To MAXCLIENTS
        Load ClientSocket(X)
    Next X
    ServerSocket(0).Listen
    ClientSocket(0).Listen
    
    SlashPath = App.Path
    ReDim Entry(0)
    If Right$(SlashPath, 1) <> "\" Then SlashPath = SlashPath & "\"
    If Len(Dir(SlashPath & "servers.csv")) = 0 Then
        Open SlashPath & "servers.csv" For Output As #1
        Close #1
    Else
        X = 0
        Open SlashPath & "servers.csv" For Input As #1
        Do Until EOF(1)
            X = X + 1
            ReDim Preserve Entry(X)
            With Entry(X)
                Input #1, .ServerName, .SID, .Pass
            End With
        Loop
        Close #1
    End If
End Sub

Private Sub AddToQueue(ByVal QNum As Integer, ByVal Number As Integer, ByVal QData As String)
    Dim X As Integer
    On Error GoTo BadNum
    If QNum Mod 2 = 0 Then QData = XORDecrypt(QData)
    X = UBound(Queue(QNum).Num(Number).iData) + 1
    ReDim Preserve Queue(QNum).Num(Number).iData(X)
    Queue(QNum).Num(Number).iData(X) = QData
    Exit Sub
BadNum:
    Call DisconnectPlayer(True, Number)
End Sub

Private Sub QueueTimer_Timer()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Temp As String
    For X = 1 To 4
        For Y = 1 To UBound(Queue(X).Num)
            If UBound(Queue(X).Num(Y).iData) <> 0 Then
                Temp = Queue(X).Num(Y).iData(1)
                For Z = 2 To UBound(Queue(X).Num(Y).iData)
                    Queue(X).Num(Y).iData(Z - 1) = Queue(X).Num(Y).iData(Z)
                Next Z
                ReDim Preserve Queue(X).Num(Y).iData(Z - 2)
                Select Case X
                    Case 1: Call SendData(True, Y, Temp)
                    Case 2: Call DoIncoming(True, Y, Temp)
                    Case 3: Call SendData(False, Y, Temp)
                    Case 4: Call DoIncoming(False, Y, Temp)
                End Select
            End If
        Next Y
    Next X
End Sub
Private Sub SendData(ByVal ToServer As Boolean, ByVal Index As Integer, ByVal SendMe As String)
    On Error GoTo ErrorTrap
    Dim XORSendMe As String
    XORSendMe = FormatPacket(SendMe, True)
    If ToServer Then
        If Index > ServerSocket.UBound Then Exit Sub
        If ServerSocket(Index).State <> sckConnected Then Exit Sub
        ServerSocket(Index).SendData XORSendMe
        If SendMe = "TERR:" Or Left(SendMe, 5) = "TBAN:" Or Left(SendMe, 5) = "MULTI" Or Left(SendMe, 5) = "OLDVR" Or SendMe = "WRONG" Then Server(Index).DisconnectMe = 5
    Else
        If Index > ClientSocket.UBound Then Exit Sub
        If ClientSocket(Index).State <> sckConnected Then Exit Sub
        ClientSocket(Index).SendData XORSendMe
    End If
    Exit Sub
ErrorTrap:
    Call DisconnectPlayer(ToServer, Index)
End Sub
Private Sub DoIncoming(ByVal FromServer As Boolean, ByVal Index As Integer, ByVal Info As String)
    Dim X As Integer
    Dim P1 As Integer
    Dim P2 As Integer
    Dim Temp As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim TempVar As Variant
    Temp = Right(Info, Len(Info) - 5)
    If FromServer Then
        Server(Index).Active = True
        Select Case Left(Info, 5)
            Case "INFO:" 'new server INFO
                If Len(Temp) < 44 Or Server(Index).ClientsKnow Then Call DisconnectPlayer(True, Index)
                For X = 1 To ServerSocket.UBound
                    If ServerSocket(X).State = sckConnected And Server(X).ClientsKnow Then
                        If Server(X).ServerName = Trim$(Mid$(Temp, 1, 20)) Then
                            Call AddToQueue(1, Index, "NMIU:")
                            Exit Sub
                        End If
                    End If
                Next X
                Server(Index).ServerName = Trim$(Mid$(Temp, 1, 20))
                Server(Index).Admin = Trim$(Mid$(Temp, 21, 20))
                Server(Index).Users = Asc(Mid$(Temp, 41, 1))
                Server(Index).MaxUsers = Asc(Mid$(Temp, 42, 1))
                Server(Index).SID = DecompressSID(Mid$(Temp, 51, 13)) '& "T"
                Server(Index).Address = ServerSocket(Index).RemoteHostIP
                Server(Index).Description = Right$(Temp, Len(Temp) - 63)
                If Trim$(Mid$(Temp, 43, 8)) <> VERSION Then
                    Call AddToQueue(1, Index, "OLDVR")
                    Exit Sub
                End If
                For X = 1 To UBound(Entry)
                    If Entry(X).ServerName = Server(Index).ServerName _
                    And Entry(X).SID <> Server(Index).SID Then
                        Call AddToQueue(1, Index, "NAMPW")
                        Exit Sub
                    End If
                Next X
                For X = 1 To UBound(Entry)
                    If Entry(X).SID = Server(Index).SID Then
                        Entry(X).ServerName = Server(Index).ServerName
                        Server(Index).Regged = True
                    End If
                Next X
                
                Call SetInfo(Index)
                Call RefreshListing
'                If Mid(Server(Index).Address, 1, 7) = "192.168" Then
'                    Server(Index).IPChangeable = True
'                    Call AddToQueue(1, Index, "RQIP:")
'                    Exit Sub
'                Else
                    Server(Index).IPChangeable = False
                'End If
                Call CSendAll("SERV:" & Server(Index).Info)
                Server(Index).ClientsKnow = True
                If Server(Index).Regged Then
                    Call AddToQueue(1, Index, "OKAY!R")
                Else
                    Call AddToQueue(1, Index, "OKAY!")
                End If
            Case "RLIP:"
                If Not Server(Index).IPChangeable Then
                    Call DisconnectPlayer(True, Index)
                    Exit Sub
                End If
                Call ChangeVar(Server(Index).Address, Index, Temp)
                Server(Index).ClientsKnow = True
                Call AddToQueue(1, Index, "OKAY!")
            Case "NAMC:" 'NAMe Change
                Call ChangeVar(Server(Index).ServerName, Index, Temp)
            Case "ADMC:" 'ADMin Change
                Call ChangeVar(Server(Index).Admin, Index, Temp)
            Case "DESC:" 'DEScription Change
                Call ChangeVar(Server(Index).Description, Index, Temp)
            Case "MAXC:" 'MAX user Change
                Call ChangeVar(Server(Index).MaxUsers, Index, Temp)
            Case "USRC:" 'USeR Change
                If Server(Index).Users = Temp Or Server(Index).ServerName = "" Then Exit Sub
                Server(Index).Users = Temp
                Call SetInfo(Index)
                Call RefreshListing
                Call CSendAll("SERV:" & Server(Index).Info)
                Server(Index).UserChanges = Server(Index).UserChanges + 1
                If Server(Index).UserChanges > 13 Then
                    Call TempBan(ShortenIP(Server(Index).Address))
                    Call AddToQueue(1, Index, "TBAN:0")
                End If
            Case "PASS:" 'Password
                For X = 1 To ServerSocket.UBound
                    If ServerSocket(X).State = sckConnected And Server(X).ClientsKnow Then
                        If Server(X).ServerName = Server(Index).ServerName Then
                            Call DisconnectPlayer(True, Index)
                            Exit Sub
                        End If
                    End If
                Next X
                
                For X = 1 To UBound(Entry)
                    If Entry(X).ServerName = Server(Index).ServerName Then
                        If Temp = Entry(X).Pass Then
                            Entry(X).SID = Server(Index).SID
                            Call AddToQueue(1, Index, "RIGHT")
                            Server(Index).Regged = True
                            If Not Server(Index).ClientsKnow Then
                                Call SetInfo(Index)
                                Call RefreshListing
                                Server(Index).IPChangeable = False
                                Call CSendAll("SERV:" & Server(Index).Info)
                                Server(Index).ClientsKnow = True
                                Call AddToQueue(1, Index, "OKAY!R")
                            End If
                        Else
                            Call AddToQueue(1, Index, "WRONG")
                        End If
                        Exit For
                    End If
                Next X
                If X > UBound(Entry) Then
                    ReDim Preserve Entry(X)
                    With Entry(X)
                        .Pass = Temp
                        .ServerName = Server(Index).ServerName
                        .SID = Server(Index).SID
                    End With
                    Call AddToQueue(1, Index, "REGED")
                End If
            Case "PONG:"
                Server(Index).SentPing = False
                Server(Index).Pongs = Server(Index).Pongs + 1
                If Server(Index).Pongs > 4 Then
                    Call TempBan(ShortenIP(Server(Index).Address))
                    Call AddToQueue(1, Index, "TBAN:0")
                End If
            Case Else
                Call DisconnectPlayer(True, Index)
        End Select
    Else
        Select Case Left(Info, 5)
            Case "PONG:"
                CPing(Index).Pongs = CPing(Index).Pongs + 1
                If CPing(Index).Pongs > 4 Then Call DisconnectPlayer(False, Index)
                CPing(Index).SentPing = False
                CPing(Index).Chances = 0
            Case Else
                Call DisconnectPlayer(False, Index)
        End Select
    End If
End Sub
Private Sub DisconnectPlayer(ByVal IsServer As Boolean, ByVal Number As Integer)
    Dim X As Integer
    'On Error Resume Next
    If IsServer Then
        If SDiscon(Number) Then Exit Sub
        SDiscon(Number) = True
        If Server(Number).ClientsKnow Then
            Call CSendAll("DISC:" & ShortenIP(Server(Number).Address))
        End If
        Server(Number) = Server(0)
        ReDim Queue(1).Num(Number).iData(0)
        ReDim Queue(2).Num(Number).iData(0)
        ServerSocket(Number).Close
        SDiscon(Number) = False
        Call RefreshListing
    Else
        If ClientSocket.UBound < Number Then Exit Sub
        If CDiscon(Number) Then Exit Sub
        CDiscon(Number) = True
        CPing(Number) = CPing(0)
        ReDim Queue(3).Num(Number).iData(0)
        ReDim Queue(4).Num(Number).iData(0)
        ClientSocket(Number).Close
        CDiscon(Number) = False
    End If
    Call UpdateClientCount
End Sub
Private Sub RefreshListing()
    Dim X As Integer
    Dim TempItem As ListItem
    ListDisplay.ListItems.Clear
    ListDisplay.Sorted = False
    For X = 1 To UBound(Server)
        If Server(X).Active Then
            Set TempItem = ListDisplay.ListItems.Add(, "SRVR:" & X, X)
            TempItem.SubItems(1) = Server(X).ServerName
            TempItem.SubItems(2) = Server(X).Address
            TempItem.SubItems(3) = Server(X).Admin
            TempItem.SubItems(4) = Server(X).Users & "/" & Server(X).MaxUsers
        End If
    Next X
    ListDisplay.Sorted = True
    Call ListDisplay_Click
End Sub
Private Function ShortenIP(ByVal IP As String) As String
    Dim Temp As String
    Dim Temp2 As String
    Dim X As Integer
    On Error GoTo NoIP
    Temp = ""
    Temp2 = ""
    For X = 1 To Len(IP)
        If Mid(IP, X, 1) = "." Then
            Temp2 = Temp2 & FHex(Temp)
            Temp = ""
        Else
            Temp = Temp & Mid(IP, X, 1)
        End If
    Next X
    Temp2 = Temp2 & FHex(Temp)
    If Len(Temp2) <> 8 Then GoTo NoIP
    ShortenIP = Temp2
    Exit Function
NoIP:
    ShortenIP = "00000000"
End Function
Private Function FHex(ByVal Number As Integer, Optional ByVal Digits As Integer = 2)
    Dim Temp As String
    Temp = Hex(Number)
    If Len(Temp) >= Digits Then
        FHex = Temp
    Else
        FHex = String(Digits - Len(Temp), "0") & Temp
    End If
End Function
Private Sub SetInfo(ByVal Index As Integer)
    With Server(Index)
        .Info = FHex(Index) & _
        .ServerName & String(20 - Len(.ServerName), " ") & _
        .Admin & String(20 - Len(.Admin), " ") & _
        FHex(.Users) & FHex(.MaxUsers) & _
        ShortenIP(.Address) & .Description
    End With
End Sub
Private Sub ChangeVar(ByRef Var As Variant, Index As Integer, NewVal As Variant)
    If Var = NewVal Or Server(Index).ServerName = "" Then Exit Sub
    Var = NewVal
    Call SetInfo(Index)
    Call RefreshListing
    Call CSendAll("SERV:" & Server(Index).Info)
    Server(Index).InfoChanges = Server(Index).InfoChanges + 1
    If Server(Index).InfoChanges = 10 Then Call AddToQueue(1, Index, "WARN:")
    If Server(Index).InfoChanges = 15 Then
        Call TempBan(ShortenIP(Server(Index).Address))
        Call AddToQueue(1, Index, "TBAN:2")
    End If
End Sub
Private Sub TempBan(ByVal iAddress As String)
    Dim X As Integer
    X = UBound(IPBan) + 1
    ReDim Preserve IPBan(X)
    IPBan(X) = "0" & iAddress
End Sub
Private Sub UpdateClientCount()
    Dim X As Integer
    Dim Y As Integer
    For X = 0 To MAXCLIENTS
        If ClientSocket(X).State = sckConnected Then Y = Y + 1
    Next X
    StatusBar1.Panels(3).Text = "Clients: " & Y
End Sub

Private Sub ServerSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DisconnectPlayer True, Index
End Sub

Public Function ChopString(ByRef Source As String, ByVal Count As Integer) As String
    Dim Temp As String
    If Source = "" Then Exit Function
    Temp = Left$(Source, Count)
    Source = Right$(Source, Len(Source) - Count)
    ChopString = Temp
End Function

Public Function DecompressSID(SID As String, Optional Fake As Boolean = False) As String
    Dim Build As String
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Build = Left$(Chr2Bin(SID), 100)
    For X = 1 To 5
        Temp = Temp & Mid$(Build, X * 20, 1)
    Next X
    Build = Temp & Build
    For X = 1 To 21
        Y = Bin2Dec(Mid$(Build, X * 5 - 4, 5))
        Y = Y + IIf(Y > 8, 56, 49)
        Mid$(Build, X, 1) = Chr$(Y)
    Next X
    If Fake Then Mid(Build, 1, 1) = "Y"
    DecompressSID = Left$(Build, 21)
End Function
Public Function Chr2Bin(ByVal ChrString As String) As String
    Dim Build As String
    Dim X As Long
    'Reverse of the above
    Build = String$(Len(ChrString) * 8, vbNullChar)
    For X = 1 To Len(ChrString)
        Mid(Build, X * 8 - 7) = Dec2Bin(Asc(Mid(ChrString, X, 1)), 8)
    Next X
    Chr2Bin = Build
End Function
Public Function Bin2Dec(ByVal BitString As String) As Long
    Dim X As Long
    Static T() As Integer
    If BitString = vbNullString Then Exit Function
    ReDim T(0 To Len(BitString) - 1)
    CopyMemory T(0), ByVal StrPtr(BitString), LenB(BitString)
    Bin2Dec = T(0) - vbKey0
    For X = 1 To UBound(T)
        Bin2Dec = Bin2Dec + Bin2Dec + T(X) - vbKey0
    Next X
End Function
Public Function Dec2Bin(ByVal X As Long, ByVal Fixed As Integer) As String
    Static lDone As Long
    Static sByte(0 To 255) As String
    Dim sNibble(0 To 15) As String
    Dim Y As Long
    'If Sgn(X) = -1 And InVBMode Then Stop
    If lDone = 0 Then
        sNibble(0) = "0000"
        sNibble(1) = "0001"
        sNibble(2) = "0010"
        sNibble(3) = "0011"
        sNibble(4) = "0100"
        sNibble(5) = "0101"
        sNibble(6) = "0110"
        sNibble(7) = "0111"
        sNibble(8) = "1000"
        sNibble(9) = "1001"
        sNibble(10) = "1010"
        sNibble(11) = "1011"
        sNibble(12) = "1100"
        sNibble(13) = "1101"
        sNibble(14) = "1110"
        sNibble(15) = "1111"
        For lDone = 0 To 255
            sByte(lDone) = sNibble(lDone \ &H10) & sNibble(lDone And &HF)
        Next
    End If
    
    If X < &H100 Then
        Dec2Bin = Right$(sByte(X), Fixed)
    ElseIf X < &H10000 Then
        Dec2Bin = Right$( _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    ElseIf X < &H1000000 Then
        Dec2Bin = Right$( _
                  sByte(X \ &H10000 And &HFF) & _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    Else
        Dec2Bin = Right$( _
                  sByte(X \ &H1000000 And &HFF) & _
                  sByte(X \ &H10000 And &HFF) & _
                  sByte(X \ &H100 And &HFF) & _
                  sByte(X And &HFF), Fixed)
    End If
    Y = Len(Dec2Bin)
    If Y < Fixed Then Dec2Bin = String(Fixed - Y, "0") & Dec2Bin
    
'    If InVBMode Then
'        If X <> Bin2Dec(Dec2Bin) Then Err.Raise 6
'    End If

End Function

