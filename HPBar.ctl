VERSION 5.00
Begin VB.UserControl ColorProgress 
   BackColor       =   &H8000000C&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   Begin VB.Label BarCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image FillBitmap 
      Height          =   285
      Index           =   2
      Left            =   0
      Picture         =   "HPBar.ctx":0000
      Top             =   2040
      Width           =   2685
   End
   Begin VB.Image FillBitmap 
      Height          =   285
      Index           =   1
      Left            =   0
      Picture         =   "HPBar.ctx":041A
      Top             =   1680
      Width           =   2685
   End
   Begin VB.Image FillBitmap 
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "HPBar.ctx":0834
      Top             =   1320
      Width           =   2685
   End
   Begin VB.Image imgLTop 
      Height          =   45
      Left            =   0
      Picture         =   "HPBar.ctx":0C4E
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imgLSide 
      Height          =   30
      Left            =   0
      Picture         =   "HPBar.ctx":0F80
      Stretch         =   -1  'True
      Top             =   45
      Width           =   45
   End
   Begin VB.Image imgTop 
      Height          =   45
      Left            =   45
      Picture         =   "HPBar.ctx":12AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Left            =   45
      Picture         =   "HPBar.ctx":15DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgLBottom 
      Height          =   45
      Left            =   0
      Picture         =   "HPBar.ctx":190D
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imgRSide 
      Height          =   30
      Left            =   0
      Picture         =   "HPBar.ctx":1C3F
      Stretch         =   -1  'True
      Top             =   45
      Width           =   45
   End
   Begin VB.Image imgRBottom 
      Height          =   45
      Left            =   0
      Picture         =   "HPBar.ctx":1F6E
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imgRTop 
      Height          =   45
      Left            =   0
      Picture         =   "HPBar.ctx":22A0
      Top             =   0
      Width           =   45
   End
   Begin VB.Image BarFill 
      Height          =   285
      Left            =   45
      Picture         =   "HPBar.ctx":25D2
      Top             =   45
      Width           =   2685
   End
   Begin VB.Image Background 
      Height          =   285
      Left            =   45
      Picture         =   "HPBar.ctx":29EC
      Top             =   45
      Width           =   2685
   End
End
Attribute VB_Name = "ColorProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CaptionStyleType
    nbNoCaption
    nbExact
    nbPercent
End Enum
Private BarMax As Single
Private BarValue As Single
Private CaptionStyle As CaptionStyleType

Event Click()
Event DblClick()

Public Property Let Value(NValue As Single)
    BarValue = NValue
    Call RefreshBar
End Property

Public Property Get Value() As Single
    Value = BarValue
End Property

Public Property Let Max(NMax As Single)
    BarMax = NMax
End Property

Public Property Get Max() As Single
    Max = BarMax
End Property

Public Property Let Caption(Style As CaptionStyleType)
    If Style >= 0 And Style <= 2 Then CaptionStyle = Style
End Property

Public Property Get Caption() As CaptionStyleType
    Caption = CaptionStyle
End Property

Sub RefreshBar()
    Dim Percent As Single
    
    If BarMax = 0 Then
        BarFill.Picture = Nothing
        BarCaption.Caption = ""
        BarFill.Picture = FillBitmap(0).Picture
        BarFill.Width = Background.Width
        Exit Sub
    End If
    If BarValue < 0 Then BarValue = 0
    If BarValue > BarMax Then BarValue = BarMax
    Percent = (BarValue * 100) / BarMax
    If Percent < 25 Then
        If BarFill.Picture <> FillBitmap(2).Picture Then BarFill.Picture = FillBitmap(2).Picture
    ElseIf Percent < 50 Then
        If BarFill.Picture <> FillBitmap(1).Picture Then BarFill.Picture = FillBitmap(1).Picture
    Else
        If BarFill.Picture <> FillBitmap(0).Picture Then BarFill.Picture = FillBitmap(0).Picture
    End If
    If BarFill.Width <> Int((Background.Width * Percent) / 100) Then BarFill.Width = Int((Background.Width * Percent) / 100)
    If BarFill.Visible <> (Percent <> 0) Then BarFill.Visible = (Percent <> 0)
    Select Case CaptionStyle
        Case 0
            BarCaption.Caption = ""
        Case 1
            BarCaption.Caption = BarValue & "/" & BarMax
        Case 2
            Percent = Round(Percent, 0)
            If BarValue > 0 And Percent = 0 Then Percent = 1
            BarCaption.Caption = Percent & "%"
    End Select
End Sub

Private Sub UserControl_Resize()
    Dim X As Integer
    Dim Y As Integer
    X = UserControl.ScaleWidth
    Y = UserControl.ScaleHeight
    imgBottom.Top = Y - 3
    imgBottom.Width = X - 6
    imgLBottom.Top = imgBottom.Top
    imgRBottom.Left = X - 3
    imgRBottom.Top = imgBottom.Top
    imgTop.Width = imgBottom.Width
    imgRTop.Left = imgRBottom.Left
    imgLSide.Height = Y - 6
    imgRSide.Left = X - 3
    imgRSide.Height = imgLSide.Height
    BarCaption.Width = X
    BarCaption.Top = (Y - BarCaption.Height) \ 2
End Sub

Private Sub BarFill_Click()
    RaiseEvent Click
End Sub
Private Sub BarFill_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub BarCaption_Click()
    RaiseEvent Click
End Sub
Private Sub BarCaption_DblClick()
    RaiseEvent DblClick
End Sub
