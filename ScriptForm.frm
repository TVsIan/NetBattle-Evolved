VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ScriptForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Script"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   5520
   End
   Begin VB.Timer tmrVar 
      Interval        =   100
      Left            =   3960
      Top             =   5520
   End
   Begin VB.CheckBox Placeholder 
      Caption         =   "Check1"
      Height          =   195
      Left            =   -1000
      TabIndex        =   1
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Script"
      TabPicture(0)   =   "ScriptForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLineNum"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtScript"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Messages"
      TabPicture(1)   =   "ScriptForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMsg(9)"
      Tab(1).Control(1)=   "txtMsg(8)"
      Tab(1).Control(2)=   "txtMsg(7)"
      Tab(1).Control(3)=   "txtMsg(6)"
      Tab(1).Control(4)=   "txtMsg(5)"
      Tab(1).Control(5)=   "txtMsg(4)"
      Tab(1).Control(6)=   "txtMsg(3)"
      Tab(1).Control(7)=   "txtMsg(2)"
      Tab(1).Control(8)=   "txtMsg(1)"
      Tab(1).Control(9)=   "VScroll1"
      Tab(1).Control(10)=   "cmdAdd"
      Tab(1).Control(11)=   "txtMsg(10)"
      Tab(1).Control(12)=   "cmdX"
      Tab(1).Control(13)=   "cmdI"
      Tab(1).Control(14)=   "Line1"
      Tab(1).Control(15)=   "lblMsg(10)"
      Tab(1).Control(16)=   "lblMsg(1)"
      Tab(1).Control(17)=   "lblMsg(2)"
      Tab(1).Control(18)=   "lblMsg(3)"
      Tab(1).Control(19)=   "lblMsg(4)"
      Tab(1).Control(20)=   "lblMsg(5)"
      Tab(1).Control(21)=   "lblMsg(6)"
      Tab(1).Control(22)=   "lblMsg(7)"
      Tab(1).Control(23)=   "lblMsg(8)"
      Tab(1).Control(24)=   "lblMsg(9)"
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "Variables"
      TabPicture(2)   =   "ScriptForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Player Arrays"
      TabPicture(3)   =   "ScriptForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TreeView"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.TreeView TreeView 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   9
         Left            =   -74160
         TabIndex        =   20
         Top             =   3540
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   8
         Left            =   -74160
         TabIndex        =   19
         Top             =   3180
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   7
         Left            =   -74160
         TabIndex        =   18
         Top             =   2820
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   6
         Left            =   -74160
         TabIndex        =   17
         Top             =   2460
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   5
         Left            =   -74160
         TabIndex        =   16
         Top             =   2100
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   4
         Left            =   -74160
         TabIndex        =   15
         Top             =   1740
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   3
         Left            =   -74160
         TabIndex        =   14
         Top             =   1380
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   2
         Left            =   -74160
         TabIndex        =   13
         Top             =   1020
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   1
         Left            =   -74160
         TabIndex        =   12
         Top             =   660
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3735
         LargeChange     =   10
         Left            =   -68760
         Max             =   0
         TabIndex        =   11
         Top             =   540
         Width           =   255
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add New"
         Height          =   375
         Left            =   -69480
         TabIndex        =   10
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Index           =   10
         Left            =   -74160
         TabIndex        =   9
         Top             =   3885
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CommandButton cmdX 
         Caption         =   "×"
         Height          =   255
         Left            =   -69120
         TabIndex        =   8
         ToolTipText     =   "Delete"
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdI 
         Caption         =   "+"
         Height          =   255
         Left            =   -69360
         TabIndex        =   7
         ToolTipText     =   "Insert New"
         Top             =   660
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Line Number"
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   5040
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin RichTextLib.RichTextBox txtScript 
         Height          =   4575
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8070
         _Version        =   393217
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         RightMargin     =   1e7
         TextRTF         =   $"ScriptForm.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8281
         View            =   3
         Arrange         =   1
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
            Text            =   "Variable"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   8246
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   -68400
         X2              =   -75000
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   10
         Left            =   -75360
         TabIndex        =   30
         Top             =   3900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   29
         Top             =   660
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   2
         Left            =   -75360
         TabIndex        =   28
         Top             =   1020
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   3
         Left            =   -75360
         TabIndex        =   27
         Top             =   1380
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   4
         Left            =   -75360
         TabIndex        =   26
         Top             =   1740
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   5
         Left            =   -75360
         TabIndex        =   25
         Top             =   2100
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   6
         Left            =   -75360
         TabIndex        =   24
         Top             =   2460
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   7
         Left            =   -75360
         TabIndex        =   23
         Top             =   2820
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   8
         Left            =   -75360
         TabIndex        =   22
         Top             =   3180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0  - "
         Height          =   255
         Index           =   9
         Left            =   -75360
         TabIndex        =   21
         Top             =   3540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblLineNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Number: "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   2295
      End
   End
End
Attribute VB_Name = "ScriptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EnterLoop As Boolean
Dim BtnScr As Integer
Dim TempPDM() As String
Dim UpdateLN As Boolean
Dim LastChar As Long
Dim ShiftKey As Boolean


Private Sub Check1_Click()
    lblLineNum.Visible = (Check1.Value = 1)
    Check1.Refresh
End Sub

Private Sub cmdAdd_Click()
    Dim Y As Long
    Y = UBound(TempPDM) + 1
    ReDim Preserve TempPDM(Y)
    UpdateScrollBar
    VScroll1.Value = VScroll1.Max
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Disregard any and all changes?", vbQuestion Or vbYesNo, "Cancel") = vbYes Then Unload Me
End Sub


Private Sub cmdI_Click()
    Dim X As Integer
    Dim Y As Integer
    Y = UBound(TempPDM) + 1
    ReDim Preserve TempPDM(Y)
    Y = cmdI.Top / 360 + VScroll1.Value - 1
    For X = UBound(TempPDM) - 1 To Y Step -1
        TempPDM(X + 1) = TempPDM(X)
    Next X
    TempPDM(Y) = ""
    UpdateScrollBar
    VScroll1_Change
End Sub

Private Sub cmdOK_Click()
    Dim Temp As String
    Dim X As Long
    Dim Y As Integer
    'On Error Resume Next
    X = FreeFile
    MainScript = txtScript.Text
    Temp = Reread(MainScript)
    If Temp <> "" Then ServerWindow.AddMessage Temp
    Open AppPath & "Script.ini" For Output As #X
    Print #X, MainScript
    Close #X
    PDM = TempPDM
    If FileExists(SlashPath & "Messages.ini") Then Kill SlashPath & "Messages.ini"
    Open AppPath & "Messages.ini" For Output As #X
    For Y = 1 To UBound(PDM)
        Print #X, PDM(Y)
    Next Y
    Close #X
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdX_Click()
    Dim X As Integer
    Dim Y As Integer
    Y = cmdX.Top / 360 + VScroll1.Value - 1
    For X = Y To UBound(TempPDM) - 1
        TempPDM(X) = TempPDM(X + 1)
    Next X
    ReDim Preserve TempPDM(X - 1)
    If cmdX.Top <> 720 Then
        cmdX.Top = cmdX.Top - 360
    Else
        BtnScr = BtnScr - 1
    End If
    UpdateScrollBar
    VScroll1_Change
End Sub

Private Sub Form_Load()
    Dim Y As Integer
    txtScript.Text = MainScript
    TempPDM = PDM
    UpdateScrollBar
    VScroll1_Change
    VScroll1.Value = 0
    SSTab1.Tab = 2
    RefreshVars
    SSTab1.Tab = 3
    RefreshVars
    SSTab1.Tab = 0
End Sub
Private Sub UpdateLineNum()
    Dim Y As Long
    Y = txtScript.SelStart
    If Y <> LastChar Then
        LastChar = Y
        Y = txtScript.GetLineFromChar(Y) + 1
        lblLineNum.Caption = "Line Number: " & CStr(Y)
    End If
End Sub
Private Sub RefreshVars()
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim T1 As String
    Dim T2 As String
    Dim Sort As Boolean
    Dim mLen As String
    Dim TempItem As ListItem
    Dim TempNode As Node
    mLen = String(Len(CStr(MaxUsers)), "0")
    If SSTab1.Tab = 2 Then
        With ListView1
            For X = 1 To .ListItems.count
                If Left(.ListItems(X).Key, 1) = "N" Then Y = Y + 1
            Next X
            For X = 1 To .ListItems.count
                If Left(.ListItems(X).Key, 1) = "T" Then Z = Z + 1
            Next X
        End With
        If UBound(NumVar) <> Y Or UBound(TxtVar) <> Z Then
            ListView1.ListItems.Clear
            For X = 1 To UBound(NumVar)
                Set TempItem = ListView1.ListItems.Add(, "N" & CStr(X), NumVar(X).vName)
                TempItem.SubItems(1) = NumVar(X).Value
            Next X
            For X = 1 To UBound(TxtVar)
                Set TempItem = ListView1.ListItems.Add(, "T" & CStr(X), TxtVar(X).vName)
                TempItem.SubItems(1) = TxtVar(X).Value
            Next X
            Sort = True
        Else
            For X = 1 To ListView1.ListItems.count
                T1 = ListView1.ListItems(X).Key
                T2 = ChopString(T1, 1)
                Y = Val(T1)
                If T2 = "N" Then
                    If ListView1.ListItems(X).Text <> NumVar(Y).vName Then
                        ListView1.ListItems(X).Text = NumVar(Y).vName
                        Sort = True
                    End If
                    If ListView1.ListItems(X).SubItems(1) <> NumVar(Y).Value Then
                        ListView1.ListItems(X).SubItems(1) = NumVar(Y).Value
                    End If
                Else
                    If ListView1.ListItems(X).Text <> TxtVar(Y).vName Then
                        ListView1.ListItems(X).Text = TxtVar(Y).vName
                        Sort = True
                    End If
                    If ListView1.ListItems(X).SubItems(1) <> TxtVar(Y).Value Then
                        ListView1.ListItems(X).SubItems(1) = TxtVar(Y).Value
                    End If
                End If
            Next X
        End If
        If Sort Then
            ListView1.Sorted = True
            ListView1.Sorted = False
        End If
    ElseIf SSTab1.Tab = 3 Then
        On Error Resume Next
        Y = 0
        Z = 0
        With TreeView
            For X = 1 To .Nodes.count
                If Left(.Nodes(X).Key, 1) = "N" Then
                    Y = Y + 1
                End If
            Next X
            For X = 1 To .Nodes.count
                If Left(.Nodes(X).Key, 1) = "T" Then
                    Z = Z + 1
                End If
            Next X
            X = 0
            If Y + Z <> 0 Then
                If .Nodes(1).Children <> ServerWindow.ListView1.ListItems.count Then X = 1
            End If
            Sort = False
            If UBound(PANum) = Y And UBound(PATxt) = Z And X = 0 Then
                X = 1
                For V = 1 To Y + Z
                    T2 = .Nodes(X).Key
                    T1 = ChopString(T2, 1)
                    Y = Val(T2)
                    If T1 = "N" Then
                        If .Nodes(X).Text <> PANum(Y).vName Then
                            .Nodes(X).Text = PANum(Y).vName
                            Sort = True
                        End If
                        W = .Nodes(X).Child.Index
                        For Z = 1 To MaxUsers
                            If IsLoaded(Z) Then
                                T1 = Format(Z, mLen) & ": " & CStr(PANum(Y).Value(Z))
                                If .Nodes(W).Text <> T1 Then
                                    .Nodes(W).Text = T1
                                    Sort = True
                                End If
                                W = .Nodes(W).Next.Index
                            End If
                        Next Z
                    Else
                        If .Nodes(X).Text <> PATxt(Y).vName Then
                            .Nodes(X).Text = PATxt(Y).vName
                            Sort = True
                        End If
                        W = .Nodes(X).Child.Index
                        For Z = 1 To MaxUsers
                            If IsLoaded(Z) Then
                                T1 = Format(Z, mLen) & ": " & PATxt(Y).Value(Z)
                                If .Nodes(W).Text <> T1 Then
                                    .Nodes(W).Text = T1
                                    Sort = True
                                End If
                                W = .Nodes(W).Next.Index
                            End If
                        Next Z
                    End If
                    X = .Nodes(X).Next.Index
                Next V
            Else
                Sort = True
                .Nodes.Clear
                For X = 1 To UBound(PANum)
                    Set TempNode = .Nodes.Add(, , "N" & CStr(X), PANum(X).vName)
                    Y = TempNode.Index
                    For W = 1 To MaxUsers
                        If IsLoaded(W) Then
                            .Nodes.Add Y, tvwChild, CStr(X) & "NV" & CStr(W), Format(W, mLen) & ": " & CStr(PANum(X).Value(W))
                        End If
                    Next W
                Next X
                For X = 1 To UBound(PATxt)
                    Set TempNode = .Nodes.Add(, , "T" & CStr(X), PATxt(X).vName)
                    Y = TempNode.Index
                    For W = 1 To MaxUsers
                        If IsLoaded(W) Then
                            .Nodes.Add Y, tvwChild, CStr(X) & "TV" & CStr(W), Format(W, mLen) & ": " & PATxt(X).Value(W)
                        End If
                    Next W
                Next X
            End If
            If Sort Then
                .Sorted = True
                .Sorted = False
            End If
        End With
    End If
End Sub
Private Sub Placeholder_GotFocus()
    txtScript.SetFocus
    SendKeys "  "
End Sub

Private Sub sBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
    If Check1.Value = 1 Then UpdateLineNum
End Sub

Private Sub tmrVar_Timer()
    RefreshVars
End Sub

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    On Error GoTo tExit
    X = 1
    Do Until 1 = 2
        If X <> Node.Index Then TreeView.Nodes(X).Expanded = False
        X = TreeView.Nodes(X).Next.Index
    Loop
tExit:
End Sub

Private Sub txtMsg_GotFocus(Index As Integer)
    cmdX.Top = 360 * (Index + 1)
    cmdI.Top = cmdX.Top
End Sub

Private Sub txtMsg_LostFocus(Index As Integer)
    Dim X As Integer
    Dim Y As Integer
    On Error Resume Next
    X = InStr(1, txtMsg(Index).Text, Chr(34))
    Do Until X = 0
        X = InStr(X + 1, txtMsg(Index).Text, Chr(34))
        If X = 0 Then
            txtMsg(Index).Text = txtMsg(Index).Text & Chr(34)
        Else
            X = InStr(X + 1, txtMsg(Index).Text, Chr(34))
        End If
    Loop
    txtMsg(Index).Text = LineCheck(txtMsg(Index).Text, "")
    TempPDM(Index + VScroll1.Value) = txtMsg(Index).Text

End Sub

Private Sub txtMsg_Validate(Index As Integer, Cancel As Boolean)
    Call txtMsg_LostFocus(Index)
End Sub

Private Sub txtScript_Click()
    UpdateLineNum
End Sub

Private Sub txtScript_GotFocus()
    UpdateLineNum
End Sub

Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then UpdateLN = True
    If Shift <> 0 Then ShiftKey = True
End Sub

Private Sub txtScript_KeyPress(KeyAscii As Integer)
    Dim Temp As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim Text As String
    Dim X As Long
    Dim Y As Long
    Dim E As String
    If EnterLoop Then
        EnterLoop = False
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Text = txtScript.Text
        Y = txtScript.SelStart
        If Y = 0 Then Exit Sub
        X = InStrRev(Text, vbNewLine, Y)
        If X = 0 Then X = -1
        X = X + 2
        Temp3 = Mid(Text, X, Y - X + 1)
        If Trim(Temp3) <> "" Then Temp3 = RTrim(Temp3)
        Temp = LTrim(Temp3)
        Temp2 = LineCheck(Temp, E, True)
        If E = "" Then
            SetRedraw txtScript.hWnd, False
            txtScript.SelStart = X - 1
            X = Len(Temp3) - Len(Temp)
            txtScript.SelLength = Len(Temp3)
            txtScript.SelText = String(X, " ") & Temp2
            SetRedraw txtScript.hWnd, True
            'txtScript.Text = Left(Text, X - 1) & Replace(Text, Temp, Temp2, X, 1)
            'txtScript.SelStart = Y - Len(Temp) + Len(Temp2)
        End If
        Y = Len(Temp3) - Len(Temp)
        KeyAscii = 0
        If Not ShiftKey Then
            EnterLoop = True
            SendKeys "{ENTER}" & String(Y, " ")
        End If
    End If
End Sub

Private Sub txtScript_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftKey = False
End Sub

Private Sub txtScript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateLineNum
End Sub

Private Sub VScroll1_Change()
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Y = UBound(TempPDM)
    Z = IIf(Y > 9, 10, Y)
    cmdX.Visible = (Z <> 0)
    cmdI.Visible = (Z <> 0)
    For X = 1 To Z
        txtMsg(X).Visible = True
        txtMsg(X).Text = TempPDM(X + VScroll1.Value)
        lblMsg(X).Visible = True
        lblMsg(X).Caption = CStr(X + VScroll1.Value) & "  - "
    Next X
    For X = X To 10
        txtMsg(X).Visible = False
        lblMsg(X).Visible = False
    Next X
    Z = 0
    X = (BtnScr - VScroll1.Value)
    BtnScr = VScroll1.Value
    X = X + cmdX.Top \ 360
    If X < 2 Then X = 2
    If X > 11 Then X = 11
    cmdX.Top = X * 360 - 30
    cmdI.Top = X * 360 - 30
End Sub


Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub
Private Sub UpdateScrollBar()
    Dim Y As Integer
    Y = UBound(TempPDM)
    If Y < 11 Then
        VScroll1.Max = 0
        VScroll1.Enabled = False
        VScroll1.Value = 0
        Call VScroll1_Change
    Else
        VScroll1.Enabled = True
        VScroll1.Max = Y - 10
    End If
End Sub

