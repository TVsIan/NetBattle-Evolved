VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "v0.0.0"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   3345
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SplashScreenUp = True
    If BetaRel = "" Then
        lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    Else
        lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & BetaRel
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SplashScreenUp = False
    If UCase(Command$) <> "SERVER" Then Loader.Visible = True
End Sub

