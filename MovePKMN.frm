VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form MovePKMN 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move To Box"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CopyCheck 
      Caption         =   "&Copy Instead of Move"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin MSComctlLib.ImageList Balls 
      Left            =   120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711680
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MovePKMN.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MovePKMN.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip BoxTabs 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1296
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   1
      TabFixedWidth   =   1759
      TabFixedHeight  =   582
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   0
      ImageList       =   "Balls"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 1"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 2"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 3"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 4"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 5"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 6"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 7"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 8"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 9"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Box 10"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label MoveCap 
      BackStyle       =   0  'Transparent
      Caption         =   "Moving:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
End
Attribute VB_Name = "MovePKMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    ToBox = -1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Byte
            
    For X = 1 To 10
        BoxTabs.Tabs(X).Image = 2
    Next X
    For X = 1 To UBound(BoxPKMN)
        BoxTabs.Tabs(BoxPKMN(X).InBox).Image = 1
    Next X
    BoxTabs.SelectedItem = BoxTabs.Tabs(FromBox)
    If CopyFlag = True Then CopyCheck.Value = 1 Else CopyCheck.Value = 0
    With BoxPKMN(MoveBoxNum)
        MoveCap.Caption = "Moving: " & .Nickname & "(Lv. " & .Level & " " & .Name & ")"
    End With
End Sub

Private Sub OKButton_Click()
    If BoxTabs.SelectedItem.Index = FromBox And CBool(CopyCheck.Value) = False Then
        MsgBox "Can't Move to the same box!", vbExclamation, "Error"
        Exit Sub
    End If
    ToBox = BoxTabs.SelectedItem.Index
    If CopyCheck.Value = 1 Then CopyFlag = True Else CopyFlag = False
    Unload Me
End Sub
