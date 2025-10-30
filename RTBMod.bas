Attribute VB_Name = "RTBMod"
Option Explicit
'This is all for the RTBs
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    NMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    Code As Long
End Type
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
Private Type RTBInfoType
    hWnd As Long
    hWndParent As Long
    ParentOrigWnd As Long
    OrigWndProc As Long
    BackupWndProc As Long
    NotBottom As Boolean
    SelectedLink As String
    MinX As Single
    MinY As Single
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal N As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwprocessid As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)
Const GWL_WNDPROC = (-4)
Const WM_USER = &H400
Const WM_DESTROY = &H2
Const WM_NOTIFY = &H4E
Const WM_PARENTNOTIFY = &H210
Const EM_SCROLLCARET = &HB7
Const EM_REPLACESEL = &HC2
Const EM_SETCHARFORMAT = (WM_USER + 68)
Const WM_SETFOCUS = &H7
Const WM_VSCROLL = &H115
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_SETCURSOR = &H20
Const WM_GETMINMAXINFO = &H24
Const EN_LINK = &H70B
Const EM_SETEVENTMASK = &H445
Const EM_GETEVENTMASK = &H43B
Const EM_GETTEXTRANGE = &H44B
Const SB_VERT = 1
Const SW_SHOW = 5
Const SIF_RANGE = &H1
Const SIF_PAGE = &H2
Const SIF_POS = &H4
Const SIF_DISABLENOSCROLL = &H8
Const SIF_TRACKPOS = &H10
Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Public NB_DNSADDR As Long
Public RTBInfo() As RTBInfoType

Public Function RTBWndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Index As Integer
    Dim s As SCROLLINFO
    Index = GetRTB(hWnd)
    If RTBInfo(Index).OrigWndProc = 0 Then GoTo DestroyIt
    Select Case msg
        Case WM_VSCROLL, EM_REPLACESEL
            RTBInfo(Index).NotBottom = False
            s.cbSize = Len(s)
            s.fMask = SIF_ALL
            If GetScrollInfo(hWnd, SB_VERT, s) Then
                If s.NMax Then
                    If s.nPos < s.NMax - (s.nPage - 1) Then
                        RTBInfo(Index).NotBottom = True
                        'Debug.Print "VSCROLL CANCEL"
                    End If
                End If
            End If
        Case EM_SCROLLCARET ', WM_SETFOCUS
            If RTBInfo(Index).NotBottom Then
                'Debug.Print "EM_SCROLLCARET CANCEL"
                RTBWndProc = True
                Exit Function
            End If
        Case EM_SETCHARFORMAT
            If RTBInfo(Index).NotBottom Then
                'LockWindowUpdate hWnd
                'Debug.Print "EM_SETCHARFORMAT CANCEL"
                RTBWndProc = CallWindowProc(RTBInfo(Index).OrigWndProc, hWnd, msg, wParam, lParam)
                'LockWindowUpdate 0
                Exit Function
            End If
        Case WM_PARENTNOTIFY
            If (wParam And &HFF) = WM_DESTROY Then GoTo DestroyIt
        Case WM_DESTROY
DestroyIt:
            SetWindowLong hWnd, GWL_WNDPROC, RTBInfo(Index).BackupWndProc
            RTBWndProc = CallWindowProc(RTBInfo(Index).BackupWndProc, hWnd, msg, wParam, lParam)
            Exit Function
    End Select
    RTBWndProc = CallWindowProc(RTBInfo(Index).OrigWndProc, hWnd, msg, wParam, lParam)
End Function

Public Function RTBParentWndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim MinMax As MINMAXINFO
    Dim sText As String
    Dim lLen As Long
    Dim Index As Integer
    Index = GetRTBParent(hWnd)
    Select Case msg
    Case WM_NOTIFY
        CopyMemory uHead, ByVal lParam, Len(uHead)
        If (uHead.hWndFrom = RTBInfo(Index).hWnd) And (uHead.Code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
            Select Case eLink.msg
            Case WM_LBUTTONDOWN
                'User clicked a link.  Store it in memory until ButtonUp triggers
                eText.chrg.cpMin = eLink.chrg.cpMin
                eText.chrg.cpMax = eLink.chrg.cpMax
                eText.lpstrText = String$(1024, " ")
                lLen = SendMessage(RTBInfo(Index).hWnd, EM_GETTEXTRANGE, 0, eText)
                sText = Left$(eText.lpstrText, lLen)
                RTBInfo(Index).SelectedLink = sText
            Case WM_LBUTTONUP
                'ButtonUp triggered over a link.  Make sure it's the same link
                'ButtonDown got, then open it.
                eText.chrg.cpMin = eLink.chrg.cpMin
                eText.chrg.cpMax = eLink.chrg.cpMax
                eText.lpstrText = String$(1024, " ")
                lLen = SendMessage(RTBInfo(Index).hWnd, EM_GETTEXTRANGE, 0, eText)
                sText = Left$(eText.lpstrText, lLen)
                If RTBInfo(Index).SelectedLink = sText Then
                    'Launch the browser
                    ShellExecute hWnd, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
                End If
                RTBInfo(Index).SelectedLink = ""
                
            'Other miscellaneous messages
            Case WM_LBUTTONDBLCLK
            Case WM_RBUTTONDBLCLK
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
            Case WM_SETCURSOR
            End Select
        End If
    Case WM_GETMINMAXINFO
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.X = RTBInfo(Index).MinX
        MinMax.ptMinTrackSize.Y = RTBInfo(Index).MinY
        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
        RTBParentWndProc = DefWindowProc(RTBInfo(Index).ParentOrigWnd, msg, wParam, lParam)
        Exit Function
    Case NB_DNSADDR
        WriteDebugLog "DNSADDR received"
        ServerWindow.SetDNSInfo wParam, lParam
    Case WM_DESTROY
DestroyIt:
        SetWindowLong hWnd, GWL_WNDPROC, RTBInfo(Index).ParentOrigWnd
        RTBParentWndProc = CallWindowProc(RTBInfo(Index).ParentOrigWnd, hWnd, msg, wParam, lParam)
        RTBInfo(Index).ParentOrigWnd = 0
        Exit Function
    End Select
    RTBParentWndProc = CallWindowProc(RTBInfo(Index).ParentOrigWnd, hWnd, msg, wParam, lParam)
End Function




Public Function GetRTB(hWnd As Long) As Integer
    Dim X As Integer
    For X = 1 To UBound(RTBInfo)
        If RTBInfo(X).hWnd = hWnd Then Exit For
    Next X
    If X > UBound(RTBInfo) Then X = 0
    GetRTB = X
End Function
Public Function GetRTBParent(hWnd As Long) As Integer
    Dim X As Integer
    For X = 1 To UBound(RTBInfo)
        If RTBInfo(X).hWndParent = hWnd Then Exit For
    Next X
    If X > UBound(RTBInfo) Then X = 0
    GetRTBParent = X
End Function
Public Function NewRTB(hWnd As Long) As Integer
    Dim X As Integer
    For X = 1 To UBound(RTBInfo)
        If RTBInfo(X).hWnd = hWnd Then Exit For
    Next X
    If X > UBound(RTBInfo) Then
        ReDim Preserve RTBInfo(X)
        RTBInfo(X).hWnd = hWnd
        NewRTB = X
    Else
        NewRTB = 0
    End If
End Function
Public Function DelRTB(hWnd As Long) As Boolean
    Dim X As Integer
    Dim Y As Boolean
    For X = 1 To UBound(RTBInfo)
        If RTBInfo(X).hWnd = hWnd Then Y = True
        If Y And X < UBound(RTBInfo) Then RTBInfo(X) = RTBInfo(X + 1)
    Next X
    If Y Then ReDim Preserve RTBInfo(X - 2)
    DelRTB = Y
End Function


Function ProcIDFromWnd(ByVal hWnd As Long) As Long
   Dim idProc As Long
   GetWindowThreadProcessId hWnd, idProc
   ProcIDFromWnd = idProc
End Function
      
Function GetWinHandle(hInstance As Long) As Long
   Const GW_HWNDNEXT = 2
   Dim tempHwnd As Long
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
End Function


