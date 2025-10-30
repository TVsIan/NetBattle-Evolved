VERSION 5.00
Begin VB.Form Rearrange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rearrange Team"
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton MoveDown 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton MoveUp 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Rearrange.frx":0000
      TabIndex        =   3
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "Rearrange.frx":008A
      Left            =   720
      List            =   "Rearrange.frx":008C
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Order"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Rearrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PositionChange(0 To 5) As Byte
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Byte
    For X = 0 To 5
        PositionChange(X) = X + 1
    Next
    Call RefreshList(0)
End Sub

Private Sub MoveDown_Click()
    Dim Temp As Byte
    Dim X As Byte
    
    If List1.ListIndex = 5 Then Exit Sub
    Temp = PositionChange(List1.ListIndex)
    PositionChange(List1.ListIndex) = PositionChange(List1.ListIndex + 1)
    PositionChange(List1.ListIndex + 1) = Temp
    Call RefreshList(List1.ListIndex + 1)
End Sub

Private Sub MoveUp_Click()
    Dim Temp As Byte
    Dim X As Byte
    
    If List1.ListIndex = 0 Then Exit Sub
    Temp = PositionChange(List1.ListIndex)
    PositionChange(List1.ListIndex) = PositionChange(List1.ListIndex - 1)
    PositionChange(List1.ListIndex - 1) = Temp
    Call RefreshList(List1.ListIndex - 1)
End Sub

Sub RefreshList(ByVal SelectMe As Byte)
    Dim X As Byte
    
    List1.Clear
    For X = 0 To 5
        If PKMN(PositionChange(X)).No > 0 Then
            List1.AddItem PKMN(PositionChange(X)).Nickname & " (" & PKMN(PositionChange(X)).Name & ")", X
        Else
            List1.AddItem "{Blank}"
        End If
    Next
    List1.ListIndex = SelectMe
End Sub

Private Sub OKButton_Click()
    Dim TempPKMN(0 To 5) As Pokemon
    Dim X As Byte
    Dim Y As Byte
    
    Y = 0
    For X = 0 To 5
        If PositionChange(X) <> X + 1 Then Y = Y + 1
    Next X
    If Y = 0 Then
        Unload Me
        Exit Sub
    End If
    
    For X = 0 To 5
        TempPKMN(X) = PKMN(PositionChange(X))
    Next
    For X = 1 To 6
        PKMN(X) = TempPKMN(X - 1)
    Next
    If TeamChangeFromMS Then MasterServer.TeamChanged = True
    Unload Me
End Sub
