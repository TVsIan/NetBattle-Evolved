VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form UserEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Data Manager"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCloser 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -240
      Top             =   4260
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   14
      Top             =   4200
      Width           =   4815
      Begin CCRProgressBar6.ccrpProgressBar ProgBar 
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         Caption         =   " "
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Populating User List..."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3840
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "User List"
      TabPicture(0)   =   "UserEdit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "NewPW"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChangePW"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DeleteUser"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "UserList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "AuthChange"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "AuthSet(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "AuthSet(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "AuthSet(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "NewUser"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtUserJump"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Word Filter"
      TabPicture(1)   =   "UserEdit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command7"
      Tab(1).Control(1)=   "List4"
      Tab(1).Control(2)=   "Command6"
      Tab(1).Control(3)=   "Text3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "IP Ban"
      TabPicture(2)   =   "UserEdit.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "Line3"
      Tab(2).Control(2)=   "cmdChange2"
      Tab(2).Control(3)=   "txtBan2"
      Tab(2).Control(4)=   "Text1"
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(6)=   "List2"
      Tab(2).Control(7)=   "Command2"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "SID Ban"
      TabPicture(3)   =   "UserEdit.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command10"
      Tab(3).Control(1)=   "List1"
      Tab(3).Control(2)=   "Command11"
      Tab(3).Control(3)=   "Text5"
      Tab(3).Control(4)=   "txtBan1"
      Tab(3).Control(5)=   "cmdChange1"
      Tab(3).Control(6)=   "Line1"
      Tab(3).Control(7)=   "Label5"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "ISP Ban"
      TabPicture(4)   =   "UserEdit.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line2"
      Tab(4).Control(1)=   "Label4"
      Tab(4).Control(2)=   "cmdChange3"
      Tab(4).Control(3)=   "txtBan3"
      Tab(4).Control(4)=   "Command5"
      Tab(4).Control(5)=   "List3"
      Tab(4).Control(6)=   "Command4"
      Tab(4).Control(7)=   "Text2"
      Tab(4).ControlCount=   8
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   41
         Top             =   3645
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Add"
         Height          =   372
         Left            =   -72360
         TabIndex        =   40
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   2985
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Remove"
         Height          =   372
         Left            =   -71160
         TabIndex        =   38
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Unban"
         Height          =   372
         Left            =   -71160
         TabIndex        =   36
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Manual Ban"
         Height          =   372
         Left            =   -72360
         TabIndex        =   34
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   33
         Top             =   3645
         Width           =   2415
      End
      Begin VB.TextBox txtBan2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   32
         Top             =   3045
         Width           =   3615
      End
      Begin VB.CommandButton cmdChange2 
         Caption         =   "Change"
         Height          =   375
         Left            =   -71160
         TabIndex        =   31
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   29
         Top             =   3645
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Manual Ban"
         Height          =   372
         Left            =   -72360
         TabIndex        =   28
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Unban"
         Height          =   372
         Left            =   -71160
         TabIndex        =   26
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtBan3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   25
         Top             =   3045
         Width           =   3615
      End
      Begin VB.CommandButton cmdChange3 
         Caption         =   "Change"
         Height          =   375
         Left            =   -71160
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Unban"
         Height          =   372
         Left            =   -71160
         TabIndex        =   22
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Ban"
         Height          =   372
         Left            =   -72360
         TabIndex        =   20
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   19
         Top             =   3645
         Width           =   2415
      End
      Begin VB.TextBox txtBan1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74880
         TabIndex        =   18
         Top             =   3045
         Width           =   3615
      End
      Begin VB.CommandButton cmdChange1 
         Caption         =   "Change"
         Height          =   375
         Left            =   -71160
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtUserJump 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "SID &Ban"
         Height          =   372
         Left            =   1380
         TabIndex        =   13
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox NewUser 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         TabIndex        =   11
         Top             =   2880
         Width           =   2415
      End
      Begin VB.OptionButton AuthSet 
         Caption         =   "&Admin"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   855
      End
      Begin VB.OptionButton AuthSet 
         Caption         =   "&Mod"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   735
      End
      Begin VB.OptionButton AuthSet 
         Caption         =   "&User"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton AuthChange 
         Caption         =   "&Authority"
         Height          =   372
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Change user authority"
         Top             =   3600
         Width           =   1095
      End
      Begin MSComctlLib.ListView UserList 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3413
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
            Text            =   "Player Name"
            Object.Width           =   5101
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Authority"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.CommandButton DeleteUser 
         Caption         =   "&Delete"
         Height          =   372
         Left            =   3780
         TabIndex        =   3
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton ChangePW 
         Caption         =   "&Password"
         Height          =   372
         Left            =   2580
         TabIndex        =   2
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox NewPW 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Line Line3 
         X1              =   -75000
         X2              =   -69960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label3 
         Caption         =   "Ban Message:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2820
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Ban Message:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   2820
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   -75000
         X2              =   -69960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         X1              =   -75000
         X2              =   -69960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label5 
         Caption         =   "Ban Message:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   2820
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   3240
         Width           =   855
      End
   End
End
Attribute VB_Name = "UserEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Closing As Boolean
Private Loading As Boolean

Private Sub AuthChange_Click()
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Index As Long
    
    For Y = 1 To UserList.ListItems.count
        If UserList.ListItems(Y).Selected Then
            For X = 0 To 2
                If AuthSet(X).Value Then Index = X + 1
            Next
            
            Call ServerDB.ChgAuth(UCase(UserList.ListItems(Y).Text), Index)
            'Call RefreshUsers
            Select Case Index
            Case 1
                UserList.ListItems(Y).SubItems(1) = "User"
            Case 2
                UserList.ListItems(Y).SubItems(1) = "Moderator"
            Case 3
                UserList.ListItems(Y).SubItems(1) = "Administrator"
            End Select
            
            For X = 1 To MaxUsers
                If Player(X).Active = True And UCase(Player(X).Name) = UCase(UserList.ListItems(Y).Text) Then
                    Call ServerWindow.SendAll("AUTH:" & FixedHex(X, 3) & Index)
                    Player(X).Authority = Index
                    ServerWindow.RefreshListing
                End If
            Next
        End If
    Next Y
End Sub

Private Sub ChangePW_Click()
    Dim X As Long
    For X = 1 To UserList.ListItems.count
        If UserList.ListItems(X).Selected Then
            If ServerDB.VIP(UserList.ListItems(X).Text) Then
                MsgBox "Unable to modify this user!", , "Error"
                Exit Sub
            End If
            Call ServerDB.ChgPwd(UCase(UserList.ListItems(X).Text), MD5(NewPW.Text))
        End If
    Next X
    NewPW.Text = ""
End Sub

Private Sub cmdChange1_Click()
    If List1.ListIndex = -1 Then Exit Sub
    Call SetSIDMessage(List1.List(List1.ListIndex), txtBan1.Text)
End Sub

Private Sub cmdChange2_Click()
    If List2.ListIndex = -1 Then Exit Sub
    Call SetIPMessage(List2.List(List2.ListIndex), txtBan2.Text)
End Sub

Private Sub cmdChange3_Click()
    If List3.ListIndex = -1 Then Exit Sub
    Call SetISPMessage(List3.List(List3.ListIndex), txtBan3.Text)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command10_Click()
    Call ServerDB.DelSIDBan(List1.List(List1.ListIndex))
    Call RefreshUBan
    Call List1_Click
End Sub

Private Sub Command11_Click()
    If Not ServerDB.AddSIDBan(Text5.Text) Then Beep
    Text5.Text = ""
    Call RefreshUBan
End Sub

Private Sub Command2_Click()
    Call ServerDB.DelIPBan(List2.List(List2.ListIndex))
    Call RefreshIPs
    Call List2_Click
End Sub

Private Sub Command3_Click()
    Call ServerDB.AddIPBan(Text1.Text)
    Call RefreshIPs
End Sub

Private Sub Command4_Click()
    Call ServerDB.AddISPBan(Text2.Text)
    Call RefreshISPs
End Sub

Private Sub Command5_Click()
    Call ServerDB.DelISPBan(List3.List(List3.ListIndex))
    Call RefreshISPs
    Call List3_Click
End Sub

Private Sub Command6_Click()
    Call ServerDB.AddWord(Text3.Text)
    Call RefreshFilter
End Sub

Private Sub Command7_Click()
    Call ServerDB.DelWord(List4.List(List4.ListIndex))
    Call RefreshFilter
End Sub

Private Sub Command8_Click()
    Call ServerDB.ProcessLogon(NewUser.Text)
    Call ServerDB.ChgPwd(NewUser.Text, MD5(NewPW.Text))
    NewUser.Text = ""
    NewPW.Text = ""
    Call RefreshUsers
End Sub

Private Sub Command9_Click()
    Dim X As Long
    For X = 1 To UserList.ListItems.count
        If UserList.ListItems(X).Selected Then
            If ServerDB.AddSIDBan(UserList.ListItems(X).Text) Then
                Call RefreshUBan
            End If
        End If
    Next X
End Sub

Private Sub DeleteUser_Click()
    Dim Y As Long
    For Y = 1 To UserList.ListItems.count
        If Y > UserList.ListItems.count Then Exit For
        If UserList.ListItems(Y).Selected Then
            If ServerDB.VIP(UserList.ListItems(Y).Text) Then
                MsgBox "Unable to delete this user!", , "Error"
            Else
                Call ServerDB.DelUser(UCase(UserList.ListItems(Y).Text))
                UserList.ListItems.Remove Y
                Y = Y - 1
            End If
        End If
    Next Y
    UserList.SelectedItem.Selected = True
    'Call RefreshUsers
End Sub

Private Sub Form_Load()
    Closing = False
    Loading = True
    Picture1.Top = 2880
    SSTab1.Tab = 0
    SSTab1.Enabled = False
    Me.Show
    Call RefreshUsers
    If Closing Then
        tmrCloser.Enabled = True
    Else
        Call RefreshIPs
        Call RefreshISPs
        Call RefreshFilter
        Call RefreshUBan
        Picture1.Visible = False
        SSTab1.Enabled = True
    End If
    Loading = False
End Sub

Private Sub RefreshUsers()
    Dim X As Long
    Dim T As Single
    T = Timer
    RedrawWindow SSTab1.hWnd, ByVal 0&, 0, RDW_ALLCHILDREN Or RDW_UPDATENOW Or RDW_INVALIDATE
    DoEvents
    UserList.ListItems.Clear
    UserList.Sorted = False
    SetRedraw SSTab1.hWnd, False
    ProgBar.Max = ServerDB.GetUserMax \ 50 + 1
    ProgBar.Value = 0
    For X = 1 To ServerDB.GetUserMax
        'SetRedraw UserList.hWnd, False
        
        UserList.ListItems.Add X, , ServerDB.GetNameByNum(X)
        Select Case ServerDB.GetAuthByNum(X)
            Case 0, 1
                UserList.ListItems(X).SubItems(1) = "User"
            Case 2
                UserList.ListItems(X).SubItems(1) = "Moderator"
            Case 3
                UserList.ListItems(X).SubItems(1) = "Administrator"
        End Select
        If X Mod 50 = 0 Then
            ProgBar.Value = ProgBar.Value + 1
            DoEvents
            If Closing Then GoTo ExitThis
        End If
    Next
    ProgBar.Value = ProgBar.Max
    DoEvents
    'Debug.Print Timer - T
    UserList.Sorted = True
ExitThis:
    SetRedraw SSTab1.hWnd, True, Not Closing
    'Debug.Print Timer - T
End Sub

Private Sub RefreshIPs()
    Dim X As Long
    
    List2.Clear
    If ServerDB.GetIPBanMax = 0 Then Exit Sub
    For X = 1 To ServerDB.GetIPBanMax
        List2.AddItem ServerDB.GetIPByNum(X)
    Next
End Sub

Private Sub RefreshISPs()
    Dim X As Long
    
    List3.Clear
    If ServerDB.GetISPBanMax = 0 Then Exit Sub
    For X = 1 To ServerDB.GetISPBanMax
        List3.AddItem ServerDB.GetISPByNum(X)
    Next
End Sub

Private Sub RefreshFilter()
    Dim X As Long
    
    List4.Clear
    If ServerDB.GetWordFilterMax = 0 Then Exit Sub
    For X = 1 To ServerDB.GetWordFilterMax
        List4.AddItem ServerDB.GetFilterByNum(X)
    Next
End Sub

Private Sub RefreshUBan()
    Dim X As Long
    
    List1.Clear
    If ServerDB.GetSIDBanMax = 0 Then Exit Sub
    For X = 1 To ServerDB.GetSIDBanMax
        List1.AddItem ServerDB.GetSIDNameByNum(X)
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Loading Then
        Closing = True
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ServerDB.WriteDB
End Sub

Private Sub List1_Click()
    If List1.ListIndex = -1 Then
        txtBan1.Text = ""
    Else
        txtBan1.Text = GetSIDMessage(List1.List(List1.ListIndex))
    End If
End Sub

Private Sub List2_Click()
    If List2.ListIndex = -1 Then
        txtBan2.Text = ""
    Else
        txtBan2.Text = GetIPMessage(List2.List(List2.ListIndex))
    End If
End Sub

Private Sub List3_Click()
    If List3.ListIndex = -1 Then
        txtBan3.Text = ""
    Else
        txtBan3.Text = GetISPMessage(List3.List(List3.ListIndex))
    End If
End Sub

Private Sub tmrCloser_Timer()
    tmrCloser.Enabled = False
    Unload Me
End Sub

Private Sub txtUserJump_GotFocus()
    Dim X As Long
    For X = 1 To UserList.ListItems.count
        UserList.ListItems(X).Selected = False
    Next X
    UserList.SelectedItem.Selected = True
End Sub

Private Sub txtUserJump_KeyPress(KeyAscii As Integer)
    Dim X As Long
    Dim Y As Long
    Dim B As Boolean
    Dim Temp As String
    If KeyAscii = 8 Then Exit Sub
    Temp = FutureText(txtUserJump, KeyAscii)
    If Temp = "" Then Exit Sub
    KeyAscii = 0
    B = False
    With UserList
        Y = Len(Temp)
        For X = 1 To .ListItems.count
            If LCase(Left(.ListItems(X).Text, Y)) = LCase(Temp) Then
                '.ListItems(X).Selected = True
                .SelectedItem.Selected = False
                Set .SelectedItem = .ListItems(X)
                .ListItems(X).EnsureVisible
                'Call UserList_ItemClick(.ListItems(X))
                txtUserJump.Text = .ListItems(X).Text
                txtUserJump.SelStart = Y
                txtUserJump.SelLength = Len(txtUserJump.Text) - Y
                B = True
                Exit For
            End If
        Next X
        If Not B Then
            X = txtUserJump.SelStart + 1
            txtUserJump.Text = Temp
            txtUserJump.SelStart = X
        End If
    End With
End Sub

Private Sub UserList_Click()
    On Error Resume Next
    txtUserJump.Text = UserList.SelectedItem.Text
End Sub

Private Sub UserList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    txtUserJump.Text = Item.Text
End Sub
