VERSION 5.00
Begin VB.Form TaskBarIcon 
   Caption         =   "TrayIcon"
   ClientHeight    =   615
   ClientLeft      =   1515
   ClientTop       =   4560
   ClientWidth     =   2235
   Icon            =   "TBIcon.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   615
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   120
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuBar 
      Caption         =   "PopupMenu"
      Begin VB.Menu mnuMain 
         Caption         =   "&Restore Window"
         Index           =   0
      End
      Begin VB.Menu mnuMain 
         Caption         =   "S&cript Window"
         Index           =   1
      End
      Begin VB.Menu mnuMain 
         Caption         =   "&Set Options"
         Index           =   2
      End
      Begin VB.Menu mnuMain 
         Caption         =   "&Data Manager"
         Index           =   3
      End
      Begin VB.Menu mnuMain 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "TaskBarIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim T As NOTIFYICONDATA

Private Sub Form_Load()
    T.cbSize = Len(T)
    T.hWnd = pichook.hWnd
    T.uId = 1&
    T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    T.ucallbackMessage = WM_MOUSEMOVE
    T.hIcon = Me.Icon
    T.szTip = "NetBattle Server" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, T
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    T.cbSize = Len(T)
    T.hWnd = pichook.hWnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
End Sub

Private Sub mnuMain_Click(Index As Integer)
    Select Case Index
        Case 0
            MainContainer.WindowState = vbNormal
            MainContainer.Visible = True
        Case 1
            ScriptForm.Show
        Case 2
            SetUsers.Show
        Case 3
            UserEdit.Show
        Case 4
            Unload MainContainer
    End Select
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                If MainContainer.WindowState = vbMinimized Then
                    MainContainer.WindowState = vbNormal
                    MainContainer.Visible = True
                Else
                    MainContainer.WindowState = vbMinimized
                    MainContainer.Visible = False
                End If
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                Me.PopupMenu mnuBar
        End Select
        rec = False
    End If
End Sub

Private Sub Timer1_Timer()
    'Refresh the icon in case Explorer.exe crashed.
    Dim B As Boolean
    B = Shell_NotifyIcon(NIM_MODIFY, T)
    If B = False Then '"If Not B Then" doesn't work here for some bizarre reason...
        T.cbSize = Len(T)
        T.hWnd = pichook.hWnd
        T.uId = 1&
        T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        T.ucallbackMessage = WM_MOUSEMOVE
        T.hIcon = Me.Icon
        T.szTip = "NetBattle Server" & Chr$(0)
        Shell_NotifyIcon NIM_DELETE, T
        Shell_NotifyIcon NIM_ADD, T
    End If
End Sub
