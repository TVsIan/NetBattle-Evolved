VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form PlayerInfoAdv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Info"
   ClientHeight    =   3030
   ClientLeft      =   5535
   ClientTop       =   4770
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "User Lookup"
      TabPicture(0)   =   "PlayerInfoAdv.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtInfo(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtInfo(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtInfo(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtInfo(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtInfo(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdLookup"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Ban List"
      TabPicture(1)   =   "PlayerInfoAdv.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRefresh"
      Tab(1).Control(1)=   "cmdUnban"
      Tab(1).Control(2)=   "BanList"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Alias Lookup"
      TabPicture(2)   =   "PlayerInfoAdv.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "NickList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdALook"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdALook 
         Caption         =   "Lookup"
         Height          =   315
         Left            =   -72240
         TabIndex        =   22
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   -74880
         TabIndex        =   21
         Top             =   420
         Width           =   2535
      End
      Begin MSComctlLib.ListView NickList 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   20
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2990
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
            Object.Width           =   3360
         EndProperty
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnban 
         Caption         =   "Unban"
         Height          =   375
         Left            =   -73080
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin MSComctlLib.ListView BanList 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   17
         Top             =   420
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Name/IP"
            Object.Width           =   3360
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ban Type"
            Object.Width           =   2434
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   2535
      End
      Begin VB.CommandButton cmdLookup 
         Caption         =   "Lookup"
         Default         =   -1  'True
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   420
         Width           =   735
      End
      Begin VB.CommandButton Command 
         Caption         =   "TempBan"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "[No User Selected]"
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1140
         Width           =   2415
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1380
         Width           =   2415
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1860
         Width           =   2415
      End
      Begin VB.CommandButton Command 
         Caption         =   "IP Ban"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command 
         Caption         =   "SID Ban"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   4
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1620
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "IP Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Station ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Authority:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   900
         Width           =   1395
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2775
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PlayerInfoAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PlayerOnline As Boolean
Public iBanList As String
Dim WithEvents TempBanInput As TextBox
Attribute TempBanInput.VB_VarHelpID = -1

Private Sub BanList_Click()
    If StatusBar.SimpleText <> "Requesting Ban List..." Then
        If BanList.SelectedItem.Tag = "X" Then
            cmdUnban.Enabled = False
        ElseIf Player(YourNumber).Authority = 1 And Left$(BanList.SelectedItem.SubItems(1), 4) <> "Temp" Then
            cmdUnban.Enabled = False
        Else
            cmdUnban.Enabled = True
        End If
    End If
End Sub

Private Sub cmdALook_Click()
    Dim X As Integer
    If cmdLookup.Enabled = False Then Exit Sub
    cmdLookup.Enabled = False
    cmdALook.Enabled = False
    For X = 0 To 2
        Command(X).Enabled = False
    Next X
    NickList.ListItems.Clear
    Call MasterServer.SendData("ALIA:" & Text2.Text)
    StatusBar.SimpleText = "Querying, Please Wait..."
End Sub

Public Sub cmdLookup_Click()
    Dim X As Integer
    If cmdLookup.Enabled = False Then Exit Sub
    cmdLookup.Enabled = False
    For X = 0 To 2
        Command(X).Enabled = False
    Next X
    txtInfo(0).Text = "[No User Selected]"
    For X = 1 To 4
        txtInfo(X).Text = ""
    Next X
    MasterServer.LookingUp = Text1.Text
    Call MasterServer.SendData("LOOK:" & Text1.Text)
    PlayerOnline = False
    For X = 1 To MaxUsers
         If LCase(Player(X).Name) = LCase(Text1.Text) Then
            PlayerOnline = True
            Command(1).Enabled = (Player(YourNumber).Authority > Player(X).Authority)
            Exit For
        End If
    Next X
    StatusBar.SimpleText = "Querying, Please Wait..."
End Sub

Private Sub cmdRefresh_Click()
    cmdRefresh.Enabled = False
    cmdUnban.Enabled = False
    StatusBar.SimpleText = "Requesting Ban List..."
    Call MasterServer.SendData("UBAN:")
End Sub

Private Sub cmdUnban_Click()
    cmdRefresh.Enabled = False
    cmdUnban.Enabled = False
    StatusBar.SimpleText = "Unban command sent."
    With BanList.SelectedItem
        Call MasterServer.SendData("UBAN:" & .Tag & .Text)
    End With
End Sub

Private Sub Command_Click(Index As Integer)
    Select Case Index
    Case 0
        TempbanSet.Show vbModal, MainContainer
        If TempbanDuration = 0 Then Exit Sub
        With MasterServer
            Call .SendData("TMPB:" & Format(TempbanDuration, "0000") & .LookingUp)
        End With
        StatusBar.SimpleText = "TempBan command sent."
    Case 1
        With MasterServer
            Call .SendData("IPBN:" & .LookingUp)
        End With
        StatusBar.SimpleText = "IP Ban command sent."
    Case 2
        With MasterServer
            Call .SendData("IDBN:" & .LookingUp)
        End With
        StatusBar.SimpleText = "SID Ban command sent."
    End Select
End Sub

Private Sub Form_Load()
    Command(1).Visible = (Player(YourNumber).Authority = 3)
    Command(2).Visible = (Player(YourNumber).Authority = 3)
    BanList.ListItems.Add , , "[Please click Refresh]"
    BanList.ListItems(1).Tag = "X"
    Set BanList.SelectedItem = BanList.ListItems(1)
    iBanList = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MasterServer.LookingUp = ""
End Sub

Public Sub ProcessBanList()
    Dim A1() As String
    Dim A2() As String
    Dim A3() As String
    Dim X As Integer
    Dim i As ListItem
    With BanList
        .ListItems.Clear
        A1 = Split(iBanList, "|")
        iBanList = ""
        A2 = Split(A1(0), ",")
        For X = 1 To UBound(A2)
            A3 = Split(A2(X), ":")
            Set i = .ListItems.Add(, , A3(0))
            i.Tag = "T"
            i.SubItems(1) = "Temp: " & A3(1) & " min"
        Next X
        A2 = Split(A1(1), ",")
        For X = 1 To UBound(A2)
            Set i = .ListItems.Add(, , A2(X))
            i.Tag = "S"
            i.SubItems(1) = "SID Ban"
        Next X
        A2 = Split(A1(2), ",")
        For X = 1 To UBound(A2)
            Set i = .ListItems.Add(, , A2(X))
            i.Tag = "I"
            i.SubItems(1) = "IP Ban"
        Next X
        If .ListItems.count = 0 Then
            .ListItems.Add , , "[No current bans]"
            .ListItems(1).Tag = "X"
        End If
        StatusBar.SimpleText = "Ban List Received."
        Set .SelectedItem = .ListItems(1)
        Call BanList_Click
        cmdRefresh.Enabled = True
    End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
    Case 0
        cmdLookup.Default = True
    Case 1
        cmdRefresh.Default = True
    Case 2
        cmdALook.Default = True
    End Select
End Sub

Private Sub Text1_Change()
    Text2.Text = Text1.Text
End Sub

Private Sub Text2_Change()
    Text1.Text = Text2.Text
End Sub
