VERSION 5.00
Begin VB.Form PInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Info"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   ClipControls    =   0   'False
   Icon            =   "PInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   120
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   1320
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   120
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   2
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   2
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   3
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   3
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   4
      Left            =   1320
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   4
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   5
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   840
         Index           =   5
         Left            =   120
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3495
   End
End
Attribute VB_Name = "PInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim TempVar As Integer
    Dim TempVar2 As Integer
    
    On Error GoTo BadPlayer
    PInfo.Icon = MainContainer.Trainers.ListImages(Player(ChallengeNumber).Picture).Picture
    PInfo.Caption = Player(ChallengeNumber).Name
    Label1.Caption = Player(ChallengeNumber).Extra & vbNewLine & "Rating: " & Player(ChallengeNumber).Rank & vbNewLine & "Won " & Player(ChallengeNumber).Wins & " - Lost " & Player(ChallengeNumber).Losses & " - Tied " & Player(ChallengeNumber).Ties & " - Quit " & Player(ChallengeNumber).Disconnect
    Label2.Caption = "Address: " & Player(ChallengeNumber).DNSAddress & " (" & Player(ChallengeNumber).Address & ")"
    Label3.Caption = "Version: " & Player(ChallengeNumber).Version
    Label4.Caption = "Speed: " & Player(ChallengeNumber).Speed
    For X = 1 To 6
        Call MainContainer.DoPicture(Player(ChallengeNumber).PKMNImage(X))
        Image1(X - 1).Picture = MainContainer.SwapSpace.Picture
        Image1(X - 1).ToolTipText = BasePKMN(Player(ChallengeNumber).PKMN(X)).Name
        TempVar = (Picture1(X - 1).Width - Image1(X - 1).Width) / 2
        TempVar2 = (Picture1(X - 1).Height - Image1(X - 1).Height) / 2
        Image1(X - 1).Left = TempVar
        Image1(X - 1).Top = TempVar2
    Next X
    Exit Sub
BadPlayer:
    Unload Me
End Sub

