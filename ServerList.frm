VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Listing"
   ClientHeight    =   5235
   ClientLeft      =   5235
   ClientTop       =   3045
   ClientWidth     =   5115
   Icon            =   "ServerList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "A&dvanced..."
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4980
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1852
            MinWidth        =   1852
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   4680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect!"
      Height          =   375
      Left            =   1770
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListDisplay 
      Height          =   2895
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Main Admin"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Users/Max"
         Object.Width           =   1746
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose a server."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   20
      Width           =   4935
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   5
      Top             =   3525
      Width           =   4905
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1035
      Left            =   45
      Top             =   3480
      Width           =   5025
   End
End
Attribute VB_Name = "ServerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ServerType
    Address As String
    ServerName As String
    Admin As String
    Users As String
    MaxUsers As String
    ServerDesc As String
End Type
Dim ServerInfo(6) As String
Dim Server() As ServerType

Private Sub cmdCancel_Click()
    ServerAddress = "Cancelled"
    If Socket.State = sckConnected Then Socket.SendData FormatPacket("EXIT:", True)
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    Dim X As Integer
    On Error GoTo ErrorExit
    X = Val(Right(ListDisplay.SelectedItem.Key, Len(ListDisplay.SelectedItem.Key) - 5))
    If Server(X).MaxUsers = Server(X).Users Then
        MsgBox "This server is currently full.  Please try again later.", vbInformation
        Exit Sub
    End If
    ServerAddress = Server(X).Address
    ServerRegName = Server(X).ServerName
    Unload Me
ErrorExit:
End Sub

Private Sub Command1_Click()
    ServerAddress = "Error"
    Unload Me
End Sub

Private Sub Form_Load()
    ReDim Server(0)
    StatusBar1.Panels(1).Text = "Connecting to Registry at " & RegAddress & "..."
    Socket.RemoteHost = RegAddress
    Socket.RemotePort = RegPortC
    Socket.Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ServerAddress = "" Then ServerAddress = "Cancelled"
End Sub

Private Sub ListDisplay_Click()
'    Dim Temp As String
'    On Error Resume Next
'    Label2.Caption = ""
'    Temp = ListDisplay.SelectedItem.Key
'    Label2.Caption = Server(Val(Right(Temp, Len(Temp) - 5))).ServerDesc
End Sub

Private Sub ListDisplay_DblClick()
    Call cmdConnect_Click
End Sub

Private Sub ListDisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim Temp As String
    On Error Resume Next
    Label2.Caption = ""
    Temp = Item.Key
    Label2.Caption = Server(Val(Right(Temp, Len(Temp) - 5))).ServerDesc
End Sub

Private Sub Socket_Close()
    StatusBar1.Panels(1).Text = "Disconnected from Registry."
End Sub

Private Sub Socket_Connect()
    StatusBar1.Panels(1).Text = "Connected!"
End Sub

Private Sub Socket_DataArrival(ByVal BytesTotal As Long)
    Dim Worked As Boolean
    Dim Packet() As String
    Dim X As Integer
    Worked = GetPacket(Socket, BytesTotal, Packet)
    If Worked Then
        For X = 1 To UBound(Packet)
            Call DoIncoming(Packet(X))
        Next X
    Else
        On Error Resume Next
        Socket.SendData FormatPacket("EXIT:", True)
    End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "The following error occured while trying to connect to the Registry:" & vbNewLine & Number & ": " & Description & vbNewLine & vbNewLine & "This error may mean that the Server Registry is temporarily disabled.  If this is the case, it should be back up shortly.  If you know your server's address or IP, you may enter it at the following screen.  If not, please visit the NetBattle forums at www.tvsian.com for a list of commonly active servers.", vbCritical, "Registry Error"
    ServerAddress = "Error"
    Unload Me
End Sub
Private Sub DoIncoming(ByVal iData As String)
    Dim Prefix As String
    Dim Data As String
    Dim Temp As String
    Dim X As Integer
    iData = XORDecrypt(iData)
    Prefix = Left(iData, 5)
    Data = Right(iData, Len(iData) - 5)
    Select Case Prefix
    Case "SERV:"
        If Data = "" Then Exit Sub
        X = Val("&H" & Left(Data, 2))
        If X > UBound(Server) Then ReDim Preserve Server(X)
        Server(X).ServerName = Trim(Mid(Data, 3, 20))
        Server(X).Admin = Trim(Mid(Data, 23, 20))
        Server(X).Users = Trim(Str(Val("&H" & Mid(Data, 43, 2))))
        Server(X).MaxUsers = Trim(Str(Val("&H" & Mid(Data, 45, 2))))
        Server(X).Address = LengthenIP(Mid(Data, 47, 8))
        Server(X).ServerDesc = Right(Data, Len(Data) - 54)
        For X = UBound(Server) To 1 Step -1
            If Server(X).ServerName = "" Then
                ReDim Preserve Server(X - 1)
            Else
                Exit For
            End If
        Next X
        Call RefreshListing
    Case "TBAN:"
        If Data = "0" Then
            MsgBox "You have been temporarily banned for PONG flooding.  Have a nice day.", vbInformation, "Temp Ban"
        Else
            MsgBox "The server registry has recorded 15 connection attempts from your IP in less than 60 seconds.  For security reasons, you have been temporarily banned from the Registry for 15 minutes.  Please try to space out your connections a bit next time.", vbInformation, "Temp Ban"
        End If
    Case "MULTI"
        MsgBox "The Registry has reported that that a connection from this IP already exists.", vbInformation
    Case "PING:"
        Socket.SendData FormatPacket("PONG:", True)
    Case "DISC:"
        Temp = LengthenIP(Data)
        For X = 1 To UBound(Server)
            If Server(X).Address = Temp Then
                Server(X) = Server(0)
                Call RefreshListing
            End If
        Next X
    End Select
End Sub
Private Function LengthenIP(ByVal IP As String) As String
    Dim Build As String
    Build = Trim(Str(Val("&H" & Mid(IP, 1, 2))))
    Build = Build & "." & Trim(Str(Val("&H" & Mid(IP, 3, 2))))
    Build = Build & "." & Trim(Str(Val("&H" & Mid(IP, 5, 2))))
    Build = Build & "." & Trim(Str(Val("&H" & Mid(IP, 7, 2))))
    LengthenIP = Build
End Function
Private Sub RefreshListing()
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim TempItem As ListItem
    Dim TempServer() As ServerType
    Dim TSNum() As Integer
    Dim Temp As String
    Dim CurrentKey As String
    If UBound(Server) = 0 Then
        ListDisplay.ListItems.Clear
        ListDisplay.ColumnHeaders(1).Width = 2000
        ListDisplay.ColumnHeaders(2).Width = 2000
        Exit Sub
    End If
    SetRedraw ListDisplay.hWnd, False
    ReDim TempServer(0)
    ReDim TSNum(0)
    For X = 1 To UBound(Server)
        If Server(X).ServerName <> "" Then
            Y = UBound(TempServer)
            For Z = 1 To Y
                If Val(TempServer(Z).Users) < Val(Server(X).Users) Then Exit For
            Next Z
            ReDim Preserve TempServer(Y + 1)
            ReDim Preserve TSNum(Y + 1)
            For Z = Y + 1 To Z + 1 Step -1
                TempServer(Z) = TempServer(Z - 1)
                TSNum(Z) = TSNum(Z - 1)
            Next Z
            TempServer(Z) = Server(X)
            TSNum(Z) = X
        End If
    Next X
    Z = ListDisplay.ListItems.count
    If Z <> 0 Then CurrentKey = ListDisplay.SelectedItem.Key Else CurrentKey = ""
    'ListDisplay.ListItems.Clear
    If UBound(TempServer) > 12 Then
        ListDisplay.ColumnHeaders(1).Width = 1880
        ListDisplay.ColumnHeaders(2).Width = 1880
    Else
        ListDisplay.ColumnHeaders(1).Width = 2000
        ListDisplay.ColumnHeaders(2).Width = 2000
    End If
    ListDisplay.Sorted = False
    For X = 1 To UBound(TempServer)
        If X > ListDisplay.ListItems.count Then
            With TempServer(X)
                Set TempItem = ListDisplay.ListItems.Add(X, "SRVR:" & TSNum(X), .ServerName)
                TempItem.SubItems(1) = .Admin
                TempItem.SubItems(2) = .Users & "/" & .MaxUsers
            End With
        ElseIf ListDisplay.ListItems(X).Key <> "SRVR:" & TSNum(X) Then
            If Listed("SRVR:" & TSNum(X)) Then
                ListDisplay.ListItems.Remove X
                X = X - 1
            Else
                With TempServer(X)
                    Set TempItem = ListDisplay.ListItems.Add(X, "SRVR:" & TSNum(X), .ServerName)
                    TempItem.SubItems(1) = .Admin
                    TempItem.SubItems(2) = .Users & "/" & .MaxUsers
                End With
            End If
        Else
            With ListDisplay.ListItems(X)
                If .Text <> TempServer(X).ServerName Then .Text = TempServer(X).ServerName
                If .SubItems(1) <> TempServer(X).Admin Then .SubItems(1) = TempServer(X).Admin
                Temp = TempServer(X).Users & "/" & TempServer(X).MaxUsers
                If .SubItems(2) <> Temp Then .SubItems(2) = Temp
            End With
        End If
    Next X
    
    'ListDisplay.Sorted = True
    If ListDisplay.ListItems.count <> 0 Then
        On Error Resume Next
        ListDisplay.ListItems(1).Selected = True
        ListDisplay.ListItems(CurrentKey).Selected = True
        Call ListDisplay_ItemClick(ListDisplay.SelectedItem)
    End If
    SetRedraw ListDisplay.hWnd, True
End Sub
Private Function Listed(Key As String) As Boolean
    Listed = False
    On Error Resume Next
    Listed = (ListDisplay.ListItems(Key).Key <> "")
End Function
