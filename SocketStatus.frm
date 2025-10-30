VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SocketStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socket Status"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   Icon            =   "SocketStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Disconnect"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Timer RefreshTimer 
      Interval        =   1000
      Left            =   240
      Top             =   5520
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SocketStatus.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SocketStatus.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SocketStatus.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SocketStatus.frx":1658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView SocketList 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7435
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Player"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Sockets"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "SocketStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Player(Index).DCReason = "Disconnected from Socket Window."
    Call ServerWindow.DisconnectPlayer(SocketList.SelectedItem.Index)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    For X = 1 To MaxUsers
        SocketList.ListItems.Add X, , Format(X, String$(Len(CStr(MaxUsers)), "0")), , 0
        SocketList.ListItems(X).SubItems(2) = "Loading..."
    Next
    Call RefreshTimer_Timer
End Sub

Private Sub RefreshTimer_Timer()
    Dim X As Integer
    
    For X = 1 To MaxUsers
        If Not IsLoaded(X) Then
            SocketList.ListItems(X).SmallIcon = 1
            SocketList.ListItems(X).SubItems(1) = ""
            SocketList.ListItems(X).SubItems(2) = "Not Loaded"
        ElseIf Disconnecting(X) Then
            SocketList.ListItems(X).SmallIcon = 3
            SocketList.ListItems(X).SubItems(1) = Player(X).Name
            SocketList.ListItems(X).SubItems(2) = "D: " & Player(X).Address
        ElseIf Player(X).Name = "" Then
            SocketList.ListItems(X).SmallIcon = 3
            SocketList.ListItems(X).SubItems(1) = ""
            SocketList.ListItems(X).SubItems(2) = "C: " & Player(X).Address
        ElseIf Chances(X) > 2 Then
            SocketList.ListItems(X).SmallIcon = 1
            SocketList.ListItems(X).SubItems(1) = Player(X).Name
            SocketList.ListItems(X).SubItems(2) = "I: " & Player(X).Address
        Else
            SocketList.ListItems(X).SmallIcon = 2
            SocketList.ListItems(X).SubItems(1) = Player(X).Name
            SocketList.ListItems(X).SubItems(2) = "N: " & Player(X).Address
        End If
    Next
End Sub

