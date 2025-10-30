VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form PokeDex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pokédex"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   4695
   Icon            =   "PokeDex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList EvoImages 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":492A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":5BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":5D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":5E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":5FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":6554
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":6AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7088
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7622
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":777C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":78D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView EvoTree 
      Height          =   3495
      Left            =   4680
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      _Version        =   393217
      Indentation     =   353
      Style           =   1
      ImageList       =   "EvoImages"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox PokedexText 
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin CCRProgressBar6.ccrpProgressBar HP 
      Height          =   255
      Left            =   120
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1185
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   480
      Width           =   1095
      Begin VB.Label NumberBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image PKMNPic 
         Height          =   960
         Left            =   0
         Top             =   0
         Width           =   1080
      End
   End
   Begin CCRProgressBar6.ccrpProgressBar Attack 
      Height          =   255
      Left            =   120
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin CCRProgressBar6.ccrpProgressBar Defense 
      Height          =   255
      Left            =   120
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin CCRProgressBar6.ccrpProgressBar Speed 
      Height          =   255
      Left            =   120
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin CCRProgressBar6.ccrpProgressBar SAttack 
      Height          =   255
      Left            =   120
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin CCRProgressBar6.ccrpProgressBar SDefense 
      Height          =   255
      Left            =   120
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   0
      Caption         =   " "
      FillColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Smooth          =   -1  'True
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Evolution"
            Object.ToolTipText     =   "Evolution"
            ImageKey        =   "Macro"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Camera"
            Object.ToolTipText     =   "Pictures"
            ImageKey        =   "Camera"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   12
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Red/Blue"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Red/Blue (Back)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Gold"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Gold (Shiny)"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Silver"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Silver (Shiny)"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Gold/Silver (Back)"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Gold/Silver (Shiny, Back)"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Advance"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Advance (Shiny)"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Advance (Back)"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Advance (Shiny, Back)"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6960
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":7F98
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":80AA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":81BC
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":82CE
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PokeDex.frx":83E0
            Key             =   "Camera"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView MoveList 
      Height          =   2295
      Left            =   1560
      TabIndex        =   13
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Types"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Move"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Learned By"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label EvoDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4680
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Type1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label MoveDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      TabIndex        =   7
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sp.Def"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sp.Att"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Defense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close Pokédex"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   3
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
Attribute VB_Name = "PokeDex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CurrentView As Integer
Dim ViewVer As Byte
Dim BackView As Boolean
Dim Shiny As Boolean
'Declarations required for the scrollable text window
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
  
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
 
Private Sub EvoTree_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo TopLevel
    EvoDesc.Caption = Node.Text & " evolves from " & Node.Parent.Text & " by " & EvoMethod(Node.Image - 2)
    Exit Sub
TopLevel:
    EvoDesc.Caption = Node.Text & " is the least evolved in it's chain."
End Sub

Private Sub PokedexText_GotFocus()
    HideCaret PokedexText.hwnd
End Sub
 
 
Private Sub PokedexText_LostFocus()
    ShowCaret PokedexText.hwnd
End Sub
 
 
Private Sub PokedexText_KeyDown(KeyCode As Integer, Shift As Integer)
 
    'Scroll the TextBox if appropriate
    Select Case KeyCode
        Case vbKeyDown
            'Scroll the text up
            VScrollTextBox PokedexText, True, False
        Case vbKeyUp
            'Scroll the text down
            VScrollTextBox PokedexText, False, False
        Case vbKeyPageDown
            'Scroll the text up
            VScrollTextBox PokedexText, True, True
        Case vbKeyPageUp
            'Scroll the text down
            VScrollTextBox PokedexText, False, True
    End Select
 
End Sub
 
 
Public Sub VScrollTextBox(ByRef TBox As TextBox, _
        ByVal ScrollDown As Boolean, ByVal PageMode As Boolean)
 
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
    Call SendMessage(TBox.hwnd, WM_VSCROLL, lParam, 0)
 
End Sub

Private Sub Form_Load()
    Dim X As Integer

    MoveList.Icons = MainContainer.Types
    MoveList.SmallIcons = MainContainer.Types
    BackView = False
    Shiny = False
    ViewVer = 3
    Toolbar1.Buttons(7).ToolTipText = "Advance"
    For X = 1 To UBound(BasePKMN)
        If BasePKMN(X).MaxHP > HP.Max Then HP.Max = BasePKMN(X).MaxHP
        If BasePKMN(X).Attack > Attack.Max Then Attack.Max = BasePKMN(X).Attack
        If BasePKMN(X).Defense > Defense.Max Then Defense.Max = BasePKMN(X).Defense
        If BasePKMN(X).Speed > Speed.Max Then Speed.Max = BasePKMN(X).Speed
        If BasePKMN(X).SpecialAttack > SAttack.Max Then SAttack.Max = BasePKMN(X).SpecialAttack
        If BasePKMN(X).SpecialDefense > SDefense.Max Then SDefense.Max = BasePKMN(X).SpecialDefense
    Next X
    CurrentView = 1
    Call LoadPokeData(CurrentView)
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0
            'Insert print code here
        Case 2
            Unload Me
        Case 3
            End
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            MainContainer.GoExplorer "http://www.tvsian.com/netbattle"
        Case 2
            frmAbout.Show 1
    End Select
End Sub

Private Sub MoveList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    Dim MoveNumber As Integer
    Dim ThisString As String

    For X = 1 To UBound(Moves)
        If Moves(X).Name = Item Then MoveNumber = X
    Next X
    If Moves(MoveNumber).Power > 0 Then ThisString = ThisString & "Power: " & Moves(MoveNumber).Power & vbNewLine
    If Moves(MoveNumber).Accuracy > 0 Then ThisString = ThisString & "Accuracy: " & Moves(MoveNumber).Accuracy & vbNewLine
    ThisString = ThisString & "PP: " & Moves(MoveNumber).PP & vbNewLine & Moves(MoveNumber).Text
    MoveDesc.Caption = ThisString
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim X As Byte
    On Error Resume Next
    Select Case Button.Key
        Case "Back"
            CurrentView = CurrentView - 1
            If CurrentView = 0 Then CurrentView = UBound(BasePKMN)
            Call LoadPokeData(CurrentView)
        Case "Find"
            Search.Show
        Case "Forward"
            CurrentView = CurrentView + 1
            If CurrentView = 252 Then CurrentView = 1
            Call LoadPokeData(CurrentView)
        Case "Evolution"
            If EvoTree.Visible Then
                PokeDex.Width = 4785
                EvoTree.Visible = False
                EvoDesc.Visible = False
                Button.Value = tbrUnpressed
            Else
                PokeDex.Width = 7695
                EvoTree.Visible = True
                EvoDesc.Visible = True
                Button.Value = tbrPressed
            End If
        Case "Camera"
            Call RandomizeView
            Call LoadPokeData(CurrentView)
    End Select
End Sub

Public Sub LoadPokeData(ByVal Number As Integer)
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    Dim X As Integer
    Dim Index As Long
    
    MoveList.ListItems.Clear
    PokeDex.Caption = "PokéDex - " & BasePKMN(Number).Name
    If BasePKMN(Number).Type2 = 0 Then
        Type1.Caption = "Type: " & Element(BasePKMN(Number).Type1)
    Else
        Type1.Caption = "Type: " & Element(BasePKMN(Number).Type1) & "/" & Element(BasePKMN(Number).Type2)
    End If
    NumberBox.Caption = "No. " & Number
    HP.Value = BasePKMN(Number).MaxHP
    Attack.Value = BasePKMN(Number).Attack
    Defense.Value = BasePKMN(Number).Defense
    Speed.Value = BasePKMN(Number).Speed
    SAttack.Value = BasePKMN(Number).SpecialAttack
    SDefense.Value = BasePKMN(Number).SpecialDefense
    If HP.Value > HP.Max / 2 Then HP.FillColor = vbGreen
    If HP.Value <= HP.Max / 2 And HP.Value > HP.Max / 4 Then HP.FillColor = vbYellow
    If HP.Value <= HP.Max / 4 Then HP.FillColor = vbRed
    Label1(0).Caption = "HP: " & HP.Value
    If Attack.Value > Attack.Max / 2 Then Attack.FillColor = vbGreen
    If Attack.Value <= Attack.Max / 2 And Attack.Value > Attack.Max / 4 Then Attack.FillColor = vbYellow
    If Attack.Value <= Attack.Max / 4 Then Attack.FillColor = vbRed
    Label1(1).Caption = "Attack: " & Attack.Value
    If Defense.Value > Defense.Max / 2 Then Defense.FillColor = vbGreen
    If Defense.Value <= Defense.Max / 2 And Defense.Value > Defense.Max / 4 Then Defense.FillColor = vbYellow
    If Defense.Value <= Defense.Max / 4 Then Defense.FillColor = vbRed
    Label1(2).Caption = "Defense: " & Defense.Value
    If Speed.Value > Speed.Max / 2 Then Speed.FillColor = vbGreen
    If Speed.Value <= Speed.Max / 2 And Speed.Value > Speed.Max / 4 Then Speed.FillColor = vbYellow
    If Speed.Value <= Speed.Max / 4 Then Speed.FillColor = vbRed
    Label1(3).Caption = "Speed: " & Speed.Value
    If SAttack.Value > SAttack.Max / 2 Then SAttack.FillColor = vbGreen
    If SAttack.Value <= SAttack.Max / 2 And SAttack.Value > SAttack.Max / 4 Then SAttack.FillColor = vbYellow
    If SAttack.Value <= SAttack.Max / 4 Then SAttack.FillColor = vbRed
    Label1(4).Caption = "Sp.Atk: " & SAttack.Value
    If SDefense.Value > SDefense.Max / 2 Then SDefense.FillColor = vbGreen
    If SDefense.Value <= SDefense.Max / 2 And SDefense.Value > SDefense.Max / 4 Then SDefense.FillColor = vbYellow
    If SDefense.Value <= SDefense.Max / 4 Then SDefense.FillColor = vbRed
    Label1(5).Caption = "Sp.Def: " & SDefense.Value
'    If Shiny Then
'        Call MainContainer.DoPicture(ChooseImage(Number, 10, 10, 10, 10, ViewVer, BackView))
'    Else
'        Call MainContainer.DoPicture(ChooseImage(Number, 15, 15, 15, 15, ViewVer, BackView))
'    End If
    PKMNPic.Picture = MainContainer.SwapSpace.Picture
    TempVar = (Picture1.Width - PKMNPic.Width) / 2
    TempVar2 = 960 - PKMNPic.Height
    PKMNPic.Left = TempVar
    PKMNPic.Top = TempVar2
    MoveList.Sorted = False
    Index = 1
    For X = 1 To UBound(BasePKMN(Number).BaseMoves)
        If Abs(BasePKMN(Number).BaseMoves(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).BaseMoves(X))).Name, , Moves(Abs(BasePKMN(Number).BaseMoves(X))).Type
            If BasePKMN(Number).Name = "Smeargle" And Moves(Abs(BasePKMN(Number).BaseMoves(X))).Name <> "Sketch" Then
                MoveList.ListItems(Index).SubItems(1) = "Sketch"
            Else
                MoveList.ListItems(Index).SubItems(1) = "Level"
            End If
            Index = Index + 1
        End If
    Next X
    For X = 1 To UBound(BasePKMN(Number).MachineMoves)
        If Abs(BasePKMN(Number).MachineMoves(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).MachineMoves(X))).Name, , Moves(Abs(BasePKMN(Number).MachineMoves(X))).Type
            MoveList.ListItems(Index).SubItems(1) = Moves(Abs(BasePKMN(Number).MachineMoves(X))).NewTM
            Index = Index + 1
        End If
    Next X
    For X = 1 To UBound(BasePKMN(Number).BreedingMoves)
        If Abs(BasePKMN(Number).BreedingMoves(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).BreedingMoves(X))).Name, , Moves(Abs(BasePKMN(Number).BreedingMoves(X))).Type
            MoveList.ListItems(Index).SubItems(1) = "Breeding"
            Index = Index + 1
        End If
    Next X
    For X = 1 To UBound(BasePKMN(Number).RBYMoves)
        If Abs(BasePKMN(Number).RBYMoves(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).RBYMoves(X))).Name, , Moves(Abs(BasePKMN(Number).RBYMoves(X))).Type
            MoveList.ListItems(Index).SubItems(1) = "R/B/Y"
            Index = Index + 1
        End If
    Next X
    For X = 1 To UBound(BasePKMN(Number).RBYTM)
        If Abs(BasePKMN(Number).RBYTM(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).RBYTM(X))).Name, , Moves(Abs(BasePKMN(Number).RBYTM(X))).Type
            MoveList.ListItems(Index).SubItems(1) = "R/B/Y " & Moves(Abs(BasePKMN(Number).RBYTM(X))).OldTM
            Index = Index + 1
        End If
    Next X
    For X = 1 To UBound(BasePKMN(Number).SpecialMoves)
        If Abs(BasePKMN(Number).SpecialMoves(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).SpecialMoves(X))).Name, , Moves(Abs(BasePKMN(Number).SpecialMoves(X))).Type
            MoveList.ListItems(Index).SubItems(1) = "Special"
            Index = Index + 1
        End If
    Next X
        For X = 1 To UBound(BasePKMN(Number).MoveTutor)
        If Abs(BasePKMN(Number).MoveTutor(X)) > 0 Then
            MoveList.ListItems.Add Index, , Moves(Abs(BasePKMN(Number).MoveTutor(X))).Name, , Moves(Abs(BasePKMN(Number).MoveTutor(X))).Type
            MoveList.ListItems(Index).SubItems(1) = "Tutor"
            Index = Index + 1
        End If
    Next X
    MoveList.Sorted = True
    MoveDesc.Caption = ""
    PokedexText.Text = ""
    If PDexText(Number).RedBlue <> "" Then
        PokedexText.Text = PokedexText.Text & "Red/Blue: " & PDexText(Number).RedBlue & vbCrLf
    End If
    If PDexText(Number).Yellow <> "" Then
        PokedexText.Text = PokedexText.Text & "Yellow: " & PDexText(Number).Yellow & vbCrLf
    End If
    If PDexText(Number).Gold <> "" Then
        PokedexText.Text = PokedexText.Text & "Gold: " & PDexText(Number).Gold & vbCrLf
    End If
    If PDexText(Number).Silver <> "" Then
        PokedexText.Text = PokedexText.Text & "Silver: " & PDexText(Number).Silver & vbCrLf
    End If
    If PDexText(Number).Crystal <> "" Then
        PokedexText.Text = PokedexText.Text & "Crystal: " & PDexText(Number).Crystal
    End If
    If PDexText(Number).Ruby <> "" Then
        PokedexText.Text = PokedexText.Text & "Ruby: " & PDexText(Number).Ruby
    End If
    If PDexText(Number).Sapphire <> "" Then
        PokedexText.Text = PokedexText.Text & "Sapphire: " & PDexText(Number).Sapphire
    End If
    Call DoEvoChart(Number)
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1
            BackView = False
            Shiny = False
            ViewVer = 0
        Case 2
            BackView = True
            Shiny = False
            ViewVer = 0
        Case 3
            BackView = False
            Shiny = False
            ViewVer = 1
        Case 4
            BackView = False
            Shiny = True
            ViewVer = 1
        Case 5
            BackView = False
            Shiny = False
            ViewVer = 2
        Case 6
            BackView = False
            Shiny = True
            ViewVer = 2
        Case 7
            BackView = True
            Shiny = False
            ViewVer = 1
        Case 8
            BackView = True
            Shiny = True
            ViewVer = 1
        Case 9
            BackView = False
            Shiny = False
            ViewVer = 3
        Case 10
            BackView = False
            Shiny = True
            ViewVer = 3
        Case 11
            BackView = True
            Shiny = False
            ViewVer = 3
        Case 12
            BackView = True
            Shiny = True
            ViewVer = 3
        Case 14
            Call RandomizeView
    End Select
    Select Case ViewVer
        Case 0
            Toolbar1.Buttons(7).ToolTipText = "Red/Blue"
        Case 1
            If Not BackView Then
                Toolbar1.Buttons(7).ToolTipText = "Gold"
            Else
                Toolbar1.Buttons(7).ToolTipText = "G/S"
            End If
        Case 2
            If Not BackView Then
                Toolbar1.Buttons(7).ToolTipText = "Silver"
            Else
                Toolbar1.Buttons(7).ToolTipText = "G/S"
            End If
        Case 3
            Toolbar1.Buttons(7).ToolTipText = "Advance"
    End Select
    If ViewVer > 0 And Shiny Then
        Toolbar1.Buttons(7).ToolTipText = Toolbar1.Buttons(7).ToolTipText & " - Shiny"
    End If
    If BackView Then
        Toolbar1.Buttons(7).ToolTipText = Toolbar1.Buttons(7).ToolTipText & " (Back)"
    End If
    Call LoadPokeData(CurrentView)
End Sub

Sub RandomizeView()
    Dim X As Integer
    
    X = Int(Rnd * 2) + 1
    If X = 1 Then BackView = True Else BackView = False
    X = Int(Rnd * 2) + 1
    If X = 1 Then Shiny = True Else Shiny = False
    X = Int(Rnd * 4)
    ViewVer = X
    Select Case ViewVer
        Case 0
            Toolbar1.Buttons(7).ToolTipText = "Red/Blue"
        Case 1
            If Not BackView Then
                Toolbar1.Buttons(7).ToolTipText = "Gold"
            Else
                Toolbar1.Buttons(7).ToolTipText = "G/S"
            End If
        Case 2
            If Not BackView Then
                Toolbar1.Buttons(7).ToolTipText = "Silver"
            Else
                Toolbar1.Buttons(7).ToolTipText = "G/S"
            End If
        Case 3
            Toolbar1.Buttons(7).ToolTipText = "Advance"
    End Select
    If ViewVer > 0 And Shiny Then
        Toolbar1.Buttons(7).ToolTipText = Toolbar1.Buttons(7).ToolTipText & " - Shiny"
    End If
    If BackView Then
        Toolbar1.Buttons(7).ToolTipText = Toolbar1.Buttons(7).ToolTipText & " (Back)"
    End If
End Sub

Sub DoEvoChart(ByVal Number As Integer)
    Dim TopLevel(6) As Integer
    Dim MidLevel(6) As Integer
    Dim LstLevel(6) As Integer
    Dim TopLevelKey As String
    Dim MidLevelKey As String
    Dim X As Byte
    
    EvoTree.Nodes.Clear
    EvoDesc.Caption = ""
    With BasePKMN(Number)
        Select Case .MyStage
            Case 0
                EvoTree.Nodes.Add , , .Name, .Name, 1, 1
            Case 1
                TopLevel(0) = .No
            Case 2
                MidLevel(0) = .No
            Case 3
                LstLevel(0) = .No
        End Select
        If .MyStage > 0 Then
            For X = 1 To 5
                If .Evo(X) > 0 Then
                    Select Case BasePKMN(.Evo(X)).MyStage
                        Case 1
                            TopLevel(X) = .Evo(X)
                        Case 2
                            MidLevel(X) = .Evo(X)
                        Case 3
                            LstLevel(X) = .Evo(X)
                    End Select
                End If
            Next
            If TopLevel(0) > 0 Then
                EvoTree.Nodes.Add , , .Name, .Name, 1, 1
                TopLevelKey = .Name
            End If
            For X = 1 To 6
                If TopLevel(X) > 0 Then
                    EvoTree.Nodes.Add , , BasePKMN(TopLevel(X)).Name, BasePKMN(TopLevel(X)).Name, 1, 1
                    TopLevelKey = BasePKMN(TopLevel(X)).Name
                End If
            Next
            If MidLevel(0) > 0 Then
                EvoTree.Nodes.Add TopLevelKey, tvwChild, .Name, .Name, .MyMethod + 2, .MyMethod + 2
                MidLevelKey = .Name
            End If
            For X = 1 To 6
                If MidLevel(X) > 0 Then
                    EvoTree.Nodes.Add TopLevelKey, tvwChild, BasePKMN(MidLevel(X)).Name, BasePKMN(MidLevel(X)).Name, .EvoM(X) + 2, .EvoM(X) + 2
                    MidLevelKey = BasePKMN(MidLevel(X)).Name
                End If
            Next
            If LstLevel(0) > 0 Then
                EvoTree.Nodes.Add MidLevelKey, tvwChild, .Name, .Name, .MyMethod + 2, .MyMethod + 2
            End If
            For X = 1 To 6
                If LstLevel(X) > 0 Then
                    EvoTree.Nodes.Add MidLevelKey, tvwChild, BasePKMN(LstLevel(X)).Name, BasePKMN(LstLevel(X)).Name, .EvoM(X) + 2, .EvoM(X) + 2
                End If
            Next
            EvoTree.Nodes(TopLevelKey).Expanded = True
            EvoTree.Nodes(MidLevelKey).Expanded = True
        End If
    End With
End Sub

