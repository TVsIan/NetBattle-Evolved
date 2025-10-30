VERSION 5.00
Begin VB.Form NetSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Connect"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Clear &History"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton Mode 
      Caption         =   "S&ingle Player (vs. AI)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton Mode 
      Caption         =   "&Host"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton Mode 
      Caption         =   "&Client"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox Address 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Mode 
      Caption         =   "S&pecific Server"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Mode 
      Caption         =   "&Server Browser"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Direct Connect:"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Connect to:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server (Name or IP Address)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "NetSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim X As Integer
    If Mode(1).Value = True Then
        ServerAddress = Address.Text
        If ServerAddress = "GoTo Listing" Then Exit Sub
        For X = 0 To 4
            If Mode(X).Value Then GameType = X
        Next X
        IsServer = False
        SaveSetting "NetBattle", "Networking", "Was Server", GameType
        SaveSetting "NetBattle", "Networking", "Last Server", ServerAddress
        SaveSetting "NetBattle", "Networking", "Last Name", ""
        Call UpdateServerList(ServerAddress)
    Else
        ServerAddress = "GoTo Listing"
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    ServerAddress = "Cancelled"
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim Answer As Integer
    Dim X As Byte
    
    Answer = MsgBox("Are you sure you want to delete your server history?", vbYesNo + vbQuestion, "Confirm")
    If Answer = vbNo Then Exit Sub
    If FileExists(SlashPath & "servlist.txt") Then Kill (SlashPath & "servlist.txt")
    For X = 2 To 100
        RecentServer(X) = ""
    Next
    RecentServer(1) = "server.netbattle.net"
    Address.Clear
    Address.AddItem RecentServer(1)
End Sub

Private Sub Form_Load()
    Dim Temp As Integer
    Dim X As Integer
    
    Call GetRecentServers
    For X = 1 To 100
        If RecentServer(X) <> "" Then Address.AddItem RecentServer(X)
    Next
    Address.Text = GetSetting("NetBattle", "Networking", "Last Server", "server.netbattle.net")
    Temp = GetSetting("NetBattle", "Networking", "Was Server", 1)
    If Temp > 1 Then Temp = 1
    If Temp < 0 Then Temp = 0
    Mode(Temp).Value = True
End Sub

Private Sub Mode_Click(Index As Integer)
    If Mode(0).Value Or Mode(3).Value Then Address.Enabled = False Else Address.Enabled = True
End Sub

Private Sub Mode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Mode(0).Value Or Mode(3).Value Then Address.Enabled = False Else Address.Enabled = True
End Sub

Private Sub Mode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mode(0).Value Or Mode(3).Value Then Address.Enabled = False Else Address.Enabled = True
End Sub
