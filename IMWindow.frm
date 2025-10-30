VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form IMWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Private Message"
   ClientHeight    =   3855
   ClientLeft      =   1200
   ClientTop       =   -9705
   ClientWidth     =   3750
   Icon            =   "IMWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChall 
      Caption         =   "Challenge..."
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton SendButton 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox ChatBox 
      Height          =   495
      Left            =   0
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   3735
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"IMWindow.frx":1272
   End
   Begin VB.PictureBox picResizer 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   375
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
End
Attribute VB_Name = "IMWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ThisPlayer As Integer
Public LongMsgBuffer As String
Public RTB As RTBClass
  
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
        
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const EM_LINESCROLL = &HB6
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Sizing As Boolean
Private SizeY As Single

Private Sub cmdChall_Click()
    Dim Temp As String
    Dim X As Integer
    Dim Y As Integer
    With MasterServer.UserList
        For X = 1 To .ListItems.count
            Temp = .ListItems(X).Key
            Y = Val(Right(Temp, Len(Temp) - 5))
            If Y = ThisPlayer Then
                .ListItems(X).Selected = True
                Call MasterServer.OpenChallenge
                Exit For
            End If
        Next X
    End With
End Sub

Private Sub Form_Activate()
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    If ThisPlayer > 0 Then IMWindowFlash(IMWindowID(ThisPlayer)) = False
End Sub

Private Sub Form_GotFocus()
    IMWindowFlash(IMWindowID(ThisPlayer)) = False
End Sub

Private Sub Form_Load()
    Me.Top = -10000
    Set RTB = New RTBClass
    RTB.SetRTBHook Messages, ChatBox, 2115, 2235
    RTB.UseTimestamp = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ThisPlayer > 0 Then
        Cancel = True
        Call MasterServer.AddToIMQueue("KILL:" & CStr(ThisPlayer))
        'SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Form_Resize()
    Dim X As Single
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
        Call MainContainer.PopButton(ThisPlayer)
    Else
        Me.Visible = True
        If Me.Width < 2115 Then Me.Width = 2115
        If Me.Height < 2235 Then Me.Height = 2235
        X = Me.Height - ChatBox.Height - 1095
        If X < 285 Then
            Call DoSize(X - 285)
        End If
        SendButton.Top = Me.Height - 885
        SendButton.Left = Me.Width - 965
        cmdChall.Top = Me.Height - 885
        ChatBox.Top = Me.Height - ChatBox.Height - 990
        ChatBox.Width = Me.Width - 120
        Messages.Width = Me.Width - 120
        Messages.Height = ChatBox.Top - 105
        picResizer.Top = Messages.Height
        picResizer.Width = Me.Width
    End If
End Sub

Sub AddMessage(ByVal Message As String, Optional ByVal DebugMessage As Boolean = False, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    If DebugMessage And Not DebugMode Then Exit Sub
    Call RTB.AddMessage(Message, BreakChar, Color, Bold, Italic)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RTB.UnsetRTBHook
End Sub

Private Sub picResizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Sizing = True
        SizeY = Y
    End If
End Sub

Private Sub picResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Z As Single
    If Button = 1 And Sizing Then
        Z = SizeY - Y
        If ChatBox.Height + Z < 285 Then
            Z = 285 - ChatBox.Height
        End If
        If Messages.Height - Z < 285 Then
            Z = Messages.Height - 285
        End If
        If Z = 0 Then Exit Sub
        SetRedraw Me.hWnd, False
        Call DoSize(Z)
        SetRedraw Me.hWnd, True
    End If
End Sub
Private Sub DoSize(Change As Single)
    picResizer.Top = picResizer.Top - Change
    ChatBox.Top = ChatBox.Top - Change
    ChatBox.Height = ChatBox.Height + Change
    Messages.Height = Messages.Height - Change
End Sub
Private Sub picResizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sizing = False
End Sub

Private Sub SendButton_Click()
    Dim X As Integer
    Dim Temp As String
    Temp = RTrim$(FilterIllegalChars(ChatBox.Text, True))
    ChatBox.Text = ""
    If Len(Temp) = 0 Then Exit Sub
    If Left$(Temp, 4) = "/me " Then
        Call AddMessage("*** " & Player(YourNumber).Name & " " & Right$(Temp, Len(Temp) - 4), False, , &HC000C0)
    Else
        Call AddMessage(Player(YourNumber).Name & ": " & Temp, False, ":", vbRed, True, False)
    End If
    While Len(Temp) > 200
        Call MasterServer.SendData("IMCH:" & Chr$(ThisPlayer) & ChopString(Temp, 200) & Chr(0))
    Wend
    Call MasterServer.SendData("IMCH:" & Chr$(ThisPlayer) & Temp)
End Sub

