VERSION 5.00
Begin VB.Form PWWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Password"
   ClientHeight    =   1590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4005
   Icon            =   "ServerPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PWBox 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This server is password protected.  Please enter the password."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "PWWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ServerPassword = ""
    PWWindow.Caption = PasswordBoxTitle
    Label1.Caption = PasswordBoxCaption
    'PWBox.SetFocus
End Sub

Private Sub OKButton_Click()
    ServerPassword = PWBox.Text
    Unload Me
End Sub
