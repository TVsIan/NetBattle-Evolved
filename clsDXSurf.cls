VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDXSurf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const OffscreenPixels As Long = 256
Public Left As Long, Top As Long, Width As Long, Height As Long
Public OffsetX As Long
Public OffsetY As Long
Public DDS7 As DirectDrawSurface7
Public AngleOfRotation As Single
Public AlphaBlend As Single
Public Visible As Boolean
Public TintRed As Single
Public TintBlue As Single
Public TintGreen As Single
Public MinFilter As CONST_D3DTEXTUREMINFILTER
Public MagFilter As CONST_D3DTEXTUREMAGFILTER
Public WhitespaceTop As Long
Public WhitespaceLeft As Long
Public WhitespaceRight As Long
Public WhitespaceBottom As Long

Private DD As DirectDraw7
Private RealHeight As Long
Private RealWidth As Long
Private scr As RECT
Private ClipX As Long
Private ClipY As Long
Private Clipped As Boolean
Private mClipSurf As DirectDrawSurface7
Private mOriginal As DirectDrawSurface7
Private mIndex As Long
Private mTrans As Long
Private mFramed As Boolean
Private mDistort As Double
Private Desc As DDSURFACEDESC2

'These have no DX purpose and are for external use
Public Tag As String
Public RealVis As Boolean
Public sngX As Single
Public sngY As Single
Public sngZ As Single

Public Property Get scrTop() As Long
    scrTop = scr.Top
End Property
Public Property Let scrTop(N As Long)
    Dim H As Long
    H = scr.Bottom - scr.Top
    scr.Top = N
    scr.Bottom = N + H
    UpdateFrame
End Property
Public Property Get scrLeft() As Long
    scrLeft = scr.Left
End Property
Public Property Let scrLeft(N As Long)
    Dim H As Long
    H = scr.Right - scr.Left
    scr.Left = N
    scr.Right = N + H
    UpdateFrame
End Property

Public Property Get scrHeight() As Long
    scrHeight = scr.Bottom - scr.Top
End Property
Public Property Let scrHeight(N As Long)
    If scrHeight = N Then Exit Property
    scr.Bottom = scr.Top + N
    Height = N
    CreateFrame
End Property
Public Property Get scrWidth() As Long
    scrWidth = scr.Right - scr.Left
End Property
Public Property Let scrWidth(N As Long)
    If scrWidth = N Then Exit Property
    scr.Right = scr.Left + N
    Width = N
    CreateFrame
End Property



Public Property Get Framed() As Boolean
    Framed = mFramed
End Property
Public Property Let Framed(F As Boolean)
    Dim r1 As RECT
    Dim r2 As RECT
    If F And Not mFramed Then
        Set mOriginal = DDS7
        mFramed = F
        scrHeight = Width
        CreateFrame
    ElseIf mFramed And Not F Then
        If Clipped Then Exit Property
        Set DDS7 = Nothing
        Set DDS7 = mOriginal
        Set mOriginal = Nothing
        mFramed = F
    End If
End Property


Public Property Let ClipSurface(C As clsDXSurf)
    Dim DC1 As Long
    Dim DC2 As Long
    Dim Pict() As Byte
    Dim r1 As RECT
    Dim r2 As RECT
    Dim ret As Long
    Dim X As Long
    Dim Y As Long
    Dim T(0 To 3) As Byte
    Dim TempDesc As DDSURFACEDESC2
    Dim EmptyDesc As DDSURFACEDESC2
    Dim Key As DDCOLORKEY

    X = C.Trans
    If X = -1 Then Exit Property
    

    C.DDS7.GetSurfaceDesc TempDesc
    EmptyDesc.lHeight = TempDesc.lHeight
    EmptyDesc.lWidth = TempDesc.lWidth
    EmptyDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT 'Or DDSD_CKSRCBLT
    EmptyDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'TempDesc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    Set mClipSurf = DD.CreateSurface(EmptyDesc)
    DC1 = C.DDS7.GetDC
    DC2 = mClipSurf.GetDC
    BitBlt DC2, 0, 0, TempDesc.lWidth, TempDesc.lHeight, DC1, 0, 0, vbSrcCopy
    C.DDS7.ReleaseDC DC1
    mClipSurf.ReleaseDC DC2
    
    CopyMemory ByVal VarPtr(T(0)), ByVal VarPtr(X), 4
    
    mClipSurf.Lock r1, TempDesc, DDLOCK_WAIT, 0
    mClipSurf.GetLockedArray Pict
    ClipX = TempDesc.lWidth
    ClipY = TempDesc.lHeight
    For Y = 0 To TempDesc.lHeight - 1
        For X = 0 To TempDesc.lWidth * 4 - 1 Step 4
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                Pict(X, Y) = 0
                Pict(X + 1, Y) = 0
                Pict(X + 2, Y) = 0
            Else
                Pict(X, Y) = 255
                Pict(X + 1, Y) = 255
                Pict(X + 2, Y) = 255
            End If
        Next X
    Next Y
    mClipSurf.Unlock r1
    
    
    Key.High = 0
    Key.Low = 0
    mClipSurf.SetColorKey DDCKEY_SRCBLT, Key
    scr.Bottom = TempDesc.lHeight
    scr.Right = TempDesc.lWidth
    Top = C.Top
    Left = C.Left
    Height = scr.Bottom
    Width = scr.Right
    'Key.high = 65535
    'Key.low = 65535
    'DDS7.SetColorKey DDCKEY_SRCBLT, Key
    'mTrans = 65535
    Clipped = True
    Framed = True
    Trans = 16777215 ' vbWhite
    
End Property
Private Sub CreateFrame()
    Dim NewDesc As DDSURFACEDESC2
    If Not mFramed Then Exit Sub
    Set DDS7 = Nothing
    NewDesc.lWidth = scrWidth
    NewDesc.lHeight = scrHeight
    NewDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
    NewDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
    NewDesc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    Set DDS7 = DD.CreateSurface(NewDesc)
    UpdateFrame
    'If Not Clipped Then Trans = mTrans
    Trans = mTrans
End Sub
Private Sub UpdateFrame()
    Dim ret As Long
    Dim r1 As RECT
    Dim r2 As RECT
    Dim R As Long
    Dim B As Long
    If Not mFramed Then Exit Sub
    'If scr.Top = 48 Then Stop
    
    r2 = scr
    R = scrWidth 'r2.Right
    B = scrHeight 'r2.Bottom
    Do While R > 0
        Do While B > 0
            If R > RealWidth Then r2.Right = RealWidth Else r2.Right = r2.Left + R
            If B > RealHeight Then r2.Bottom = RealHeight Else r2.Bottom = r2.Top + B
            r1.Bottom = r1.Top + r2.Bottom - r2.Top
            r1.Right = r1.Left + r2.Right - r2.Left
            ret = DDS7.Blt(r1, mOriginal, r2, DDBLT_WAIT)
            'Debug.Print Round(Timer);
            If ret <> 0 And InVBMode Then Debug.Print "atari" & vbNewLine: Stop
            r1.Top = r1.Bottom
            B = B - (r2.Bottom - r2.Top)
            r2.Top = 0
        Loop
        r1.Top = 0
        B = scrHeight
        R = R - (r2.Right - r2.Left)
        r1.Left = r1.Right
        r2.Top = scr.Top
        r2.Left = 0
    Loop
    r1.Top = 0
    r1.Left = 0
    r1.Bottom = Height
    r1.Right = Width
    If Clipped Then ret = DDS7.Blt(r1, mClipSurf, r1, DDBLT_WAIT Or DDBLT_KEYSRC)
End Sub

Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Get Trans() As Long
    Trans = mTrans
End Property
' -2 = Upperleft Pixel is Trans, -1 = NoTrans, >=0 = Specific Trans color
Public Property Let Trans(T As Long)
    Dim Key As DDCOLORKEY
    Dim R As RECT
    Dim C(3) As Byte
    If T = -2 Then
        DDS7.Lock R, Desc, DDLOCK_WAIT, 0
        Key.High = DDS7.GetLockedPixel(0, 0)
        Key.Low = Key.High
        DDS7.Unlock R
        mTrans = Key.High
        DDS7.SetColorKey DDCKEY_SRCBLT, Key
    ElseIf T = -1 Then
        mTrans = -1
    ElseIf T >= 0 Then
    
        CopyMemory ByVal VarPtr(C(0)), ByVal VarPtr(T), 4
        Key.High = RGB(C(2), C(1), C(0))
        Key.Low = Key.High
        DDS7.SetColorKey DDCKEY_SRCBLT, Key
        mTrans = T
    End If
    SetWhitespace
End Property

Public Sub Init(NewIndex As Long, DDRef As DirectDraw7)
    mIndex = NewIndex
    DDS7.GetSurfaceDesc Desc
    Height = Desc.lHeight
    Width = Desc.lWidth
    RealHeight = Height
    RealWidth = Width
    scr.Top = 0
    scr.Left = 0
    scr.Bottom = Height
    scr.Right = Width
    TintRed = 1
    TintBlue = 1
    TintGreen = 1
    AlphaBlend = 1
    mTrans = -1
    Visible = True
    MinFilter = D3DTFN_POINT
    MagFilter = D3DTFG_POINT
    SetWhitespace
    Set DD = DDRef
End Sub
Private Sub Class_Terminate()
    Set DD = Nothing
    Set DDS7 = Nothing
    Set mClipSurf = Nothing
    Set mOriginal = Nothing
End Sub
Public Sub Invert()
    Dim Pict() As Byte
    Dim R As RECT
    Dim X As Long
    Dim Y As Long
    Dim T(0 To 3) As Byte
    If mTrans <> -1 Then CopyMemory ByVal VarPtr(T(0)), ByVal VarPtr(mTrans), 4
    DDS7.Lock R, Desc, DDLOCK_WAIT, 0
    DDS7.GetLockedArray Pict
    For Y = 0 To Desc.lHeight - 1
        For X = 0 To Desc.lWidth * 4 - 1 Step 4
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                Pict(X, Y) = 255 - Pict(X, Y)
                Pict(X + 1, Y) = 255 - Pict(X + 1, Y)
                Pict(X + 2, Y) = 255 - Pict(X + 2, Y)
            End If
        Next X
    Next Y
    DDS7.Unlock R
End Sub


Public Sub UpdatePicture(NewPicture As PictureBox)
    Dim DC As Long
    If Framed Then
        DC = mOriginal.GetDC
        BitBlt DC, 0, 0, RealWidth, RealHeight, NewPicture.hdc, 0, 0, vbSrcCopy
        mOriginal.ReleaseDC DC
    Else
        DC = DDS7.GetDC
        BitBlt DC, 0, 0, RealWidth, RealHeight, NewPicture.hdc, 0, 0, vbSrcCopy
        DDS7.ReleaseDC DC
    End If
    SetWhitespace
End Sub
Public Sub Move(Optional ByVal nLeft As Long = -1, Optional ByVal nTop As Long = -1, Optional ByVal nWidth As Long = -1, Optional ByVal nHeight As Long = -1)
    If nTop <> -1 Then Top = nTop
    If nLeft <> -1 Then Left = nLeft
    If nHeight > -1 Then Height = nHeight
    If nWidth > -1 Then Width = nWidth
End Sub
Private Function SetWhitespace() As Long
    Dim Pict() As Byte
    Dim R As RECT
    Dim X As Long
    Dim Y As Long
    Dim T(0 To 3) As Byte
    Dim W As Long
    Dim H As Long
    If mTrans <> -1 Then CopyMemory ByVal VarPtr(T(0)), ByVal VarPtr(mTrans), 4
    DDS7.Lock R, Desc, DDLOCK_WAIT, 0
    DDS7.GetLockedArray Pict
    
    WhitespaceTop = -1
    WhitespaceLeft = -1
    WhitespaceBottom = -1
    WhitespaceRight = -1
    
    W = (Desc.lWidth - 1) * 4
    H = Desc.lHeight - 1
    
    For Y = 0 To H
        For X = 0 To W Step 4
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                WhitespaceTop = Y
                Exit For
            End If
        Next X
        If WhitespaceTop > -1 Then Exit For
    Next Y
    If Y = Desc.lHeight Then
        WhitespaceTop = Desc.lHeight
        WhitespaceLeft = Desc.lWidth
        WhitespaceBottom = Desc.lHeight
        WhitespaceRight = Desc.lWidth
        DDS7.Unlock R
        Exit Function
    End If
    
    For X = 0 To W Step 4
        For Y = WhitespaceTop To H
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                WhitespaceLeft = X / 4
                Exit For
            End If
        Next Y
        If WhitespaceLeft > -1 Then Exit For
    Next X
    For X = W To 0 Step -4
        For Y = WhitespaceTop To H
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                WhitespaceRight = Desc.lWidth - X / 4 - 1
                Exit For
            End If
        Next Y
        If WhitespaceRight > -1 Then Exit For
    Next X
    For Y = H To 0 Step -1
        For X = 0 To W Step 4
            If Pict(X, Y) <> T(0) Or Pict(X + 1, Y) <> T(1) Or Pict(X + 2, Y) <> T(2) Then
                WhitespaceBottom = H - Y
                Exit For
            End If
        Next X
        If WhitespaceBottom > -1 Then Exit For
    Next Y
    DDS7.Unlock R
End Function
Public Sub HFlip()
    Dim Pict() As Byte
    Dim lngPict() As Long
    Dim R As RECT
    Dim X As Long
    Dim Y As Long
    Dim W As Long
    Dim X1 As Long
    Dim X2 As Long
    Dim Z As Long
    DDS7.Lock R, Desc, DDLOCK_WAIT, 0
    DDS7.GetLockedArray Pict
    W = Desc.lWidth * 4 - 4
    For X = 0 To Desc.lWidth * 2 - 1 Step 4
        For Y = 0 To Desc.lHeight - 1
            X1 = X
            X2 = W - X
            For Z = 1 To 4
                Pict(X1, Y) = Pict(X1, Y) Xor Pict(X2, Y)
                Pict(X2, Y) = Pict(X1, Y) Xor Pict(X2, Y)
                Pict(X1, Y) = Pict(X1, Y) Xor Pict(X2, Y)
                X1 = X1 + 1: X2 = X2 + 1
            Next Z
        Next Y
    Next X
    DDS7.Unlock R
End Sub
Public Sub VFlip()

End Sub
'Private Sub Dump(s As DirectDrawSurface7)
'    Dim DC As Long
'    On Error GoTo ETrap
'    DC = s.GetDC
'    BitBlt Battle.picTest.hdc, 0, 0, RealWidth, RealHeight, DC, 0, 0, vbSrcCopy
'    s.ReleaseDC DC
'    Battle.picTest.Refresh
'ETrap:
'End Sub
