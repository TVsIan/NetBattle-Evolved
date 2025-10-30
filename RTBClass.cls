VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This handles the stuff for all the RichTextBoxes throughout the program.
'----
Private Type QueueType
    Message As String
    BreakChar As String
    Color As Long
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    SkipNewline As Boolean
End Type
'Various system stuff
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const GWL_WNDPROC = (-4)
Const EM_SETEVENTMASK = &H445
Const EM_GETEVENTMASK = &H43B
Const EM_GETTEXTRANGE = &H44B
Const EM_AUTOURLDETECT = &H45B
Const EN_LINK = &H70B
Const ENM_LINK = &H4000000
Const EM_GETLINECOUNT = &HBA
Const EM_LINEINDEX = &HBB
Const EM_LINELENGTH = &HC1
Const EM_SCROLL = &HB5
Const SB_LINEDOWN = 1
'The RichTextBox attached to the class
Public WithEvents RTB As RichTextBox
Attribute RTB.VB_VarHelpID = -1
'The input box to pass the focus to
Public InputBox As TextBox
'The original Window Procedure, to be restored at closing time
Public OrigWndProc As Long
'Whether or not a RTB has been attached yet
Public RTBSet As Boolean
'For storing queued messages
Private MessageQueue() As QueueType
'Whether or not user is selecting stuff
Private Paused As Boolean
Private CurrentOffset As Long
Private RTBIndex As Long
Public UseTimestamp As Boolean
Public LimitText As Boolean
Public LineLimit As Long

Public Sub SetRTBHook(iRTB As RichTextBox, iInputBox As TextBox, Optional ByVal MinParentWidth As Single = -1, Optional ByVal MinParentHeight As Single = -1)
    Dim X As Long
    LineLimit = DisplayLines
    iRTB.HideSelection = True
    X = NewRTB(iRTB.hWnd)
    RTBIndex = X
    If X > 0 Then
        RTBInfo(X).hWndParent = iRTB.Parent.hWnd
        Set InputBox = iInputBox 'Subclassing in VB can lead to some freaky
        Set RTB = iRTB           'results, including crashing VB all together.
        If Not InVBMode Then     'Only enable it if you have to.
            RTBInfo(X).OrigWndProc = SetWindowLong(RTB.hWnd, GWL_WNDPROC, AddressOf RTBWndProc)
            RTBInfo(X).ParentOrigWnd = SetWindowLong(RTB.Parent.hWnd, GWL_WNDPROC, AddressOf RTBParentWndProc)
        End If
        RTBInfo(X).BackupWndProc = RTBInfo(X).OrigWndProc
        RTBSet = True
        'Tell the RTB to start automatically detecting URLs
        SendMessage RTB.hWnd, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(RTB.hWnd, EM_GETEVENTMASK, 0, 0)
        SendMessage RTB.hWnd, EM_AUTOURLDETECT, ByVal IIf(ParseURLs, 1, 0), ByVal 0
        RTBInfo(X).MinX = MinParentWidth / Screen.TwipsPerPixelX
        RTBInfo(X).MinY = MinParentHeight / Screen.TwipsPerPixelY
    End If
End Sub
Public Sub UnsetRTBHook()
    Dim X As Integer
    Dim R As Long
    If Not RTBSet Then Exit Sub
    RTBSet = False
    X = GetRTB(RTB.hWnd)
    If X > 0 Then
        If RTBInfo(X).OrigWndProc > 0 Then
            R = SetWindowLong(RTB.hWnd, GWL_WNDPROC, RTBInfo(X).OrigWndProc)
        End If
        If RTBInfo(X).ParentOrigWnd > 0 Then
            R = SetWindowLong(RTB.Parent.hWnd, GWL_WNDPROC, RTBInfo(X).ParentOrigWnd)
        End If
        DelRTB RTB.hWnd
    End If
    Set RTB = Nothing
End Sub
Public Sub AddMessage(ByVal Message As String, Optional ByVal BreakChar As String = "", Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal SkipNewline As Boolean = False, Optional ByVal Underline As Boolean = False)
    Dim LeftHalf As String
    Dim RightHalf As String
    Dim L As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim B As Boolean
    If Not RTBSet Then Exit Sub
    On Error Resume Next
    If Paused Then
        ReDim Preserve MessageQueue(UBound(MessageQueue) + 1)
        With MessageQueue(UBound(MessageQueue))
            .Message = Message
            .BreakChar = BreakChar
            .Color = Color
            .Bold = Bold
            .Italic = Italic
            .Underline = Underline
            .SkipNewline = SkipNewline
        End With
        Exit Sub
    End If
    SetRedraw RTB.hWnd, False
    'Debug.Print RTBInfo(RTBIndex).NotBottom;
    
    L = SendMessage(RTB.hWnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    If LimitText And Not RTBInfo(RTBIndex).NotBottom Then
        Do Until L < LineLimit
            Y = RTB.SelStart
            RTB.SelStart = 0
            X = SendMessage(RTB.hWnd, EM_LINELENGTH, ByVal 0&, ByVal 0&)
            RTB.SelLength = X + 2
            RTB.SelText = vbNullString
            RTB.SelStart = Y - X - 2
            L = L - 1
        Loop
        RTBInfo(RTBIndex).NotBottom = False
    End If
    B = RTBInfo(RTBIndex).NotBottom
    'Debug.Print B;
    
    If RTB.Text <> "" And Not SkipNewline Then Call RTBAdd(vbNewLine)
    'Add the timestamp
    If UseTS And TSFormat <> "" And UseTimestamp Then
        RTBAdd (Replace(Format(Now, Replace(TSFormat, "[", Chr(1))), Chr(1), "[") & " ")
    End If
    If FancyText Then
        RightHalf = Message
        If BreakChar <> "" Then
            LeftHalf = ChopString(RightHalf, InStr(1, RightHalf, BreakChar))
        End If
        If LeftHalf <> "" Then
            Call RTBAdd(LeftHalf, Color, Bold, Italic, Underline)
            Call RTBAdd(RightHalf)
        Else
            Call RTBAdd(RightHalf, Color, Bold, Italic, Underline)
        End If
    Else
        RTB.Text = RTB.Text & Message
    End If
    
    'Debug.Print RTBInfo(RTBIndex).NotBottom;

    If Not B Then
        SendMessage RTB.hWnd, EM_SCROLL, SB_LINEDOWN, 0
        RTBInfo(RTBIndex).NotBottom = False
    End If
    
    SetRedraw RTB.hWnd, True
    'Debug.Print RTBInfo(RTBIndex).NotBottom

End Sub
Private Sub RTBAdd(ByVal Message As String, Optional ByVal Color As Long = vbBlack, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional Underline As Boolean = False)
    With RTB
        .SelStart = Len(RTB.Text)
        .SelLength = 0
        .SelText = Message
        .SelStart = .SelStart - Len(Message)
        .SelLength = Len(Message)
        .SelBold = Bold
        .SelItalic = Italic
        .SelUnderline = Underline
        .SelColor = Color
        .SelHangingIndent = 200
        .SelLength = 0
    End With
End Sub
Private Sub Class_Initialize()
    ReDim MessageQueue(0)
End Sub

Private Sub RTB_GotFocus()
    If Not Paused Then InputBox.SetFocus
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
    InputBox.SetFocus
End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'User is selecting text, pause the incoming messages
    Paused = True
    RTBInfo(GetRTB(RTB.hWnd)).SelectedLink = ""
End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'User is done.  Copy the selected text to the clipboard and add any messages he missed.
    Dim Z As Long
    If RTB.SelText <> "" Then
        On Error GoTo ClipboardError
        Clipboard.Clear
        Clipboard.SetText RTB.SelText
        On Error GoTo 0
        RTB.SelLength = 0
    End If
    InputBox.SetFocus
    Paused = False
    For Z = 1 To UBound(MessageQueue)
        With MessageQueue(Z)
            Call AddMessage(.Message, .BreakChar, .Color, .Bold, .Italic, .SkipNewline, .Underline)
        End With
    Next Z
    ReDim MessageQueue(0)
    Exit Sub
ClipboardError:
    Sleep 1
    Resume
End Sub
