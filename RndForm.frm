VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form RndForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RndForm"
   ClientHeight    =   525
   ClientLeft      =   -3300
   ClientTop       =   -2355
   ClientWidth     =   1020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   1020
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   0
   End
End
Attribute VB_Name = "RndForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strike As Byte
Private Sub Form_Load()
    'Me.Hide
    '>>> Call WriteDebugLog("RndForm Loaded.")
    Strike = 0
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If RndState = rQuerying Then Inet.Cancel
End Sub

Private Sub Timer1_Timer()
    On Error GoTo Timeout
    Timer1.Enabled = False
    Inet.RequestTimeout = 120
    '>>> Call WriteDebugLog("Beginning Random Number download.")
    TmpByte = Inet.OpenURL("http://www.random.org/cgi-bin/randbyte?nbytes=" & CStr(RndGroup) & "&format=f", icByteArray)
    DoEvents
    '>>> Call WriteDebugLog("Random Number download complete.")
    If RndGroup <> UBound(TmpByte) + 1 Then
        ServerWindow.AddMessage "Error downloading random numbers.  Switching to Pseudo-Random mode."
        UseTrueRnd = False
    Else
        RndCache = RndGroup
        RndByte = TmpByte
        Erase TmpByte
        ServerWindow.AddMessage "Random number download successful.", True, False
    End If
    RndState = rReady
    '>>> Call WriteDebugLog("Unloading RndForm")
    Unload Me
    Exit Sub
Timeout:
    If Err.Number <> icTimeout Then
        If InVBMode Then Stop
        ServerWindow.AddMessage "Error downloading random numbers.  Switching to Pseudo-Random mode."
        UseTrueRnd = False
        RndState = rReady
        Unload Me
        Exit Sub
    End If
    Strike = Strike + 1
    If Strike = 3 Then
        ServerWindow.AddMessage "Random.org is inaccessable.  Switching to Pseudo-Random mode."
        UseTrueRnd = False
        Unload Me
    Else
        Timer1.Enabled = True
    End If
End Sub
