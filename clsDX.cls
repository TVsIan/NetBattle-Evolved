VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'DirectX PictureBox Controller
'Written mostly by MasamuneXGP, some things pulled from various tutorials

Public LastFPS As Long
Private FPS As Long
Private LastTime As Double

Const PI As Double = 3.14159265358979
Const ImagePath As String = "C:\NetBattle Source Code\NB 0.9.3 Working (2)\Graphics\images\"
Const OffscreenPixels As Long = 256
'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)

Private WithEvents BG As PictureBox
Attribute BG.VB_VarHelpID = -1
Dim DX As DirectX7
Dim DD As DirectDraw7
Dim D3D As Direct3D7
Dim DEV As Direct3DDevice7

'The first buffer holds a bitmap. The second "Primary" surface
'represents what appears on screen.
Dim mSurface() As clsDXSurf
Dim ZOrder() As Long
Public BackBuffer As DirectDrawSurface7
Public Primary As DirectDrawSurface7
Public NumSurfaces As Long

'HiRes timer
Public WithEvents cTmr As cHiResTimer
Attribute cTmr.VB_VarHelpID = -1
Private mTFR As Long
Public TimerKey As String


Dim PrimaryDesc As DDSURFACEDESC2
Dim BBDesc As DDSURFACEDESC2

'The clipper handles obscured surfaces and stops our application
'from drawing over the top of other windows.
Dim ddClipper As DirectDrawClipper

'A simple initialization flag.
Dim bInit As Boolean

Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As Any) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Any) As Long
Private FrameTime As Currency
Private FrameLimit As Currency

'External Use
Private Anims() As AnimType
Public CurrentFrame As Long
Public Param1 As Long
Public Param2 As Long
Public Event AnimationFinished()
Public Event BltError()

Public Sub DeleteSurface(Index As Long)
    Dim X As Long
    Dim Y As Long
    Set mSurface(Index) = Nothing
    For X = Index To NumSurfaces - 1
        Set mSurface(X) = mSurface(X + 1)
    Next X
    Set mSurface(NumSurfaces) = Nothing
    For X = 0 To NumSurfaces
        If ZOrder(X) = Index Then Exit For
    Next X
    Y = (NumSurfaces - X) * 4
    If Y > 0 Then CopyMemory ByVal VarPtr(ZOrder(X)), ByVal VarPtr(ZOrder(X + 1)), Y
    NumSurfaces = NumSurfaces - 1
    ReDim Preserve mSurface(NumSurfaces)
    ReDim Preserve ZOrder(NumSurfaces)
    For X = 0 To NumSurfaces
        If ZOrder(X) > Index Then
            ZOrder(X) = ZOrder(X) - 1
        End If
    Next X
End Sub

Public Function InitDX(Background As PictureBox) As Boolean
    Dim Caps As DDSCAPS2
    Dim Guid As String
    Dim X As Long
    Dim DEnum As Direct3DEnumDevices
    Dim Freq As Currency
    On Error GoTo ETrap
    Set DX = New DirectX7
    Set DD = DX.DirectDrawCreate("")
    Set BG = Background
    
    BG.ScaleMode = vbPixels
    'Set to Windowed mode
    Call DD.SetCooperativeLevel(BG.Parent.hWnd, DDSCL_NORMAL)
    PrimaryDesc.lFlags = DDSD_CAPS
    PrimaryDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set Primary = DD.CreateSurface(PrimaryDesc)
    
    BBDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    BBDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    BBDesc.lWidth = BG.ScaleWidth + OffscreenPixels * 2
    BBDesc.lHeight = BG.ScaleHeight + OffscreenPixels * 2
    Set BackBuffer = DD.CreateSurface(BBDesc)


    Set D3D = DD.GetDirect3D
    Set DEnum = D3D.GetDevicesEnum
    For X = 1 To DEnum.GetCount
        If DEnum.GetGuid(X) = DeviceGUID Then Exit For
    Next X
    If X = DEnum.GetCount + 1 Then X = X - 1

    Set DEV = D3D.CreateDevice(DEnum.GetGuid(X), BackBuffer)
    'Set DEV = D3D.CreateDevice("IID_IDirect3DRGBDevice", BackBuffer)
    
    'Creates the clipper, and attaches it to the picturebox and
    'the primary surface. This is all that has to be done - the clipper
    'itself handles everything else.
    
    Set ddClipper = DD.CreateClipper(0)
    ddClipper.SetHWnd BG.hWnd
    Primary.SetClipper ddClipper
    
    'Yes it has been initialized and is ready to blit
    bInit = True
    
    CurrentFrame = 0
    ReDim Anims(0)
        
    'Ok now were ready to blit this thing, call the blt procedure
    'One huge advantage of Windowed mode is that you don't
    'have to have a loop, you just call "blt" when you need the
    'picture to be updated.
    Blt
    
    
    If UseHiResTimer Then
        X = QueryPerformanceFrequency(Freq)
        If X = 0 Then
            MsgBox "This system does not support High Performance Timers.  You can still see the animations, but they may appear too slow, or too fast.", vbExclamation, "No HiPerform Timer"
            UseHiResTimer = False
        Else
            FrameLimit = Freq * 0.016666666666 '<--- 1/60 of a second
            'Debug.Print Freq
            
            Randomize Timer
            X = Rnd * 65535
            TimerKey = "DX" & FixedHex(X, 4)
            X = Rnd * 65535
            TimerKey = TimerKey & FixedHex(X, 4)
            Set cTmr = New cHiResTimer
            cTmr.Add TimerKey, RefreshRate, True
            mTFR = 60
        End If
    End If
    
    InitDX = True
    Exit Function
ETrap:
    InitDX = False
End Function

Public Sub Blt()
    'Some local variables
    Dim TempVerts() As D3DTLVERTEX
    Dim ddrval As Long
    Dim BGRect As RECT
    Dim r1 As RECT 'The screen size
    Dim r2 As RECT 'The bitmap size
    Dim Z As Long
    Dim a As Long
    Dim Offset As Long
    Dim T As Currency
    Dim Over As Currency
    
    On Error GoTo ETrap:
    'Debug.Print Rnd

    'Has it been initialized? If not let's get out of this procedure
    If bInit = False Then Exit Sub
    FPS = FPS + 1

    'Gets the bounding rect for the entire window handle, stores in r1

    'Using Blt instead of Bltfast is essential - do not try and use bltfast.
    'The advantage of using blt is that it resizes the picture to be the same as
    'the picture box, this means that we can resize the window and the code
    'will adapt to fit the new size - even though it will look really ugly
    'when stretched.
    
    
    
    If UBound(Anims) > 0 Then
        If UseHiResTimer Then
            QueryPerformanceCounter T
            Over = T - FrameTime
        End If
        If Over >= FrameLimit Or FrameTime = 0 Or Not UseHiResTimer Then
            'FPS = FPS + 1
            If FrameTime = 0 Then
                FrameTime = T
            Else
                FrameTime = T - (Over - FrameLimit)
            End If
            
            CurrentFrame = CurrentFrame + 1
            a = 0
            For Z = 1 To UBound(Anims)
                If Anims(Z).ID <> 0 Then Call DoAnimKernel(Me, Anims(Z)): a = 1
            Next Z
            If a = 0 Or CurrentFrame > 1200 Then
                ReDim Anims(0)
                CurrentFrame = 0
                FrameTime = 0
            End If
        End If
    End If

    With r1
        .Bottom = BBDesc.lHeight
        .Left = 0
        .Right = BBDesc.lWidth
        .Top = 0
    End With
    
    'Fill the backbuffer with white
    BackBuffer.BltColorFill r1, vbWhite
    
    Call DX.GetWindowRect(BG.hWnd, BGRect)
    Offset = OffscreenPixels
    If BG.BorderStyle = 1 Then Offset = Offset + 1
    
    DEV.BeginScene
    
    For Z = 0 To NumSurfaces
        With mSurface(ZOrder(Z))
            If .Visible Then
                r2.Top = 0
                r2.Left = 0
                r2.Bottom = .Height
                r2.Right = .Width
                r1.Top = Offset + .Top
                r1.Left = Offset + .Left
                r1.Bottom = r1.Top + .Height
                r1.Right = r1.Left + .Width
            
                'Set up the TempVerts(3) vertices
                SetUpGeom TempVerts, ZOrder(Z), r2, r1, .TintRed, .TintGreen, .TintBlue, .AlphaBlend, .AngleOfRotation
                
                'Enable alpha-blending
                DEV.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True
                'Enable color-keying (ColorKey is drawn transparent)
                DEV.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
                DEV.SetRenderState D3DRENDERSTATE_COLORKEYBLENDENABLE, True
                DEV.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
                DEV.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
                DEV.SetRenderState D3DRENDERSTATE_TEXTUREFACTOR, DX.CreateColorRGBA(1, 1, 1, .AlphaBlend)
                DEV.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR
                
                DEV.SetTextureStageState 0, D3DTSS_MAGFILTER, .MagFilter ' D3DTFG_GAUSSIANCUBIC
                DEV.SetTextureStageState 0, D3DTSS_MINFILTER, .MinFilter   ' D3DTFN_LINEAR 'D3DTFN_POINT '
                
                'Set the texture on the D3D device
                DEV.SetTexture 0, .DDS7
                DEV.SetTextureStageState 0, D3DTSS_MIPFILTER, 3
                'Draw the triangles that make up our square texture
                DEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, TempVerts(0), 4, D3DDP_DEFAULT
                'Turn off alphablending after we're done
                DEV.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
            End If
        End With
    Next Z
    DEV.EndScene
    
    If LastTime < 2 Or Timer - LastTime >= 0.5 Then
        LastTime = Timer
        LastFPS = FPS * 2
        FPS = 0
    End If
    BackBuffer.DrawText BBDesc.lWidth - OffscreenPixels - 62, BBDesc.lHeight - OffscreenPixels - 15, "FPS: " & CStr(LastFPS), False
    
    r1 = BGRect
    r2.Left = 0
    r2.Top = 0
    r2.Bottom = BBDesc.lHeight
    r2.Right = BBDesc.lWidth
    r1.Top = r1.Top - OffscreenPixels
    r1.Left = r1.Left - OffscreenPixels
    r1.Bottom = r1.Top + r2.Bottom
    r1.Right = r1.Left + r2.Right
    ddrval = Primary.Blt(r1, BackBuffer, r2, DDBLT_WAIT)
    
    Exit Sub
ETrap:
    'MsgBox "There was an error whilst redrawing the screen.", vbCritical, "error"
    If InVBMode Then
        Stop
        Resume
    End If
    ReDim Anims(0)
    RaiseEvent BltError
    bInit = False
End Sub

Public Sub AddAnim(ByVal ID As Long, ByVal StartFrame As Long, Optional ByVal Param1 As Long = 0, Optional ByVal Param2 As Long = 0)
    ReDim Preserve Anims(UBound(Anims) + 1)
    With Anims(UBound(Anims))
        .Index = UBound(Anims)
        .ID = ID
        .StartFrame = StartFrame
        .P1 = Param1
        .P2 = Param2
    End With
End Sub
Public Sub AnimFinished(Index As Long)
    Dim X As Long
    Anims(Index).ID = 0
    For X = 1 To UBound(Anims)
        If Anims(X).ID <> 0 Then Exit For
    Next X
    If X = UBound(Anims) + 1 Then
        CurrentFrame = 0
        RaiseEvent AnimationFinished
    End If
End Sub

Private Sub BG_Paint()
    'DD.RestoreAllSurfaces
    'Blt
End Sub
Private Sub SetUpGeom(Verts() As D3DTLVERTEX, SurfIndex As Long, src As RECT, Dest As RECT, R As Single, G As Single, b As Single, a As Single, ByVal Angle As Single)
    'This sub sets up the vertices for a sprite, taking into account
    'width, height, vertex color, and rotation angle
    'NOTE: R, G, and B dictate the color that the sprite will be -
    '1, 1, 1 is normal, lower values will colorize the vertices
    
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    
    Dim SurfW As Single
    Dim SurfH As Single
    Dim XCenter As Single
    Dim YCenter As Single
    Dim Radius As Single
    Dim XCor As Single
    Dim YCor As Single
    Dim RelX As Single
    Dim RelY As Single
    Dim CurA As Single
    Dim T As Single
    Dim Color As Long
    Dim Test As Boolean
    
    Test = False
    
    ReDim Verts(3)
    
    
    'Width of the surface
    SurfW = mSurface(SurfIndex).Width
    'Height of the surface
    SurfH = mSurface(SurfIndex).Height
    'Center coordinates on screen of the sprite
    XCenter = mSurface(SurfIndex).OffsetX + Dest.Left + (Dest.Right - Dest.Left - 1) / 2
    YCenter = mSurface(SurfIndex).OffsetY + Dest.Top + (Dest.Bottom - Dest.Top - 1) / 2
    Color = DX.CreateColorRGBA(R, G, b, a)
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        XCor = Dest.Left
        YCor = Dest.Bottom
    Else
        RelX = Dest.Left - XCenter
        RelY = YCenter - Dest.Bottom
        T = Sqr(RelX * RelX + RelY * RelY)
        CurA = Atn(RelY / RelX) + Angle
        XCor = XCenter + Cos(CurA) * T * Sgn(RelX)
        YCor = YCenter - Sin(CurA) * T * Sgn(RelX)
    End If
    
    '0 - Bottom left vertex
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, Color, 0, src.Left / SurfW, (src.Bottom + 1) / SurfH, Verts(0)
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        XCor = Dest.Left
        YCor = Dest.Top
    Else
        RelX = Dest.Left - XCenter
        RelY = YCenter - Dest.Top
        T = Sqr(RelX * RelX + RelY * RelY)
        CurA = Atn(RelY / RelX) + Angle
        XCor = XCenter + Cos(CurA) * T * Sgn(RelX)
        YCor = YCenter - Sin(CurA) * T * Sgn(RelX)
    End If
    
    '1 - Top left vertex
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, Color, 0, src.Left / SurfW, src.Top / SurfH, Verts(1)
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        XCor = Dest.Right
        YCor = Dest.Bottom
    Else
        RelX = Dest.Right - XCenter
        RelY = YCenter - Dest.Bottom
        T = Sqr(RelX * RelX + RelY * RelY)
        CurA = Atn(RelY / RelX) + Angle
        XCor = XCenter + Cos(CurA) * T * Sgn(RelX)
        YCor = YCenter - Sin(CurA) * T * Sgn(RelX)
    End If
    
    '2 - Bottom right vertex
    T = src.Right                 'This is necessary.
    If SurfW < 128 Then T = T + 1 'Don't ask me why, I have no idea...
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, Color, 0, T / SurfW, (src.Bottom + 1) / SurfH, Verts(2)
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        XCor = Dest.Right
        YCor = Dest.Top
    Else
        RelX = Dest.Right - XCenter
        RelY = YCenter - Dest.Top
        T = Sqr(RelX * RelX + RelY * RelY)
        CurA = Atn(RelY / RelX) + Angle
        XCor = XCenter + Cos(CurA) * T * Sgn(RelX)
        YCor = YCenter - Sin(CurA) * T * Sgn(RelX)
    End If
    
    '3 - Top right vertex
    DX.CreateD3DTLVertex XCor, YCor, 0, 1, Color, 0, (src.Right + 1) / SurfW, src.Top / SurfH, Verts(3)
End Sub

Public Function CreateSurfaceFromPBox(Picture As PictureBox, Optional Trans As Long = -1) As Long
    Dim DC As Long
    'Dim R As RECT
    Dim Desc As DDSURFACEDESC2
    NumSurfaces = NumSurfaces + 1
    ReDim Preserve mSurface(NumSurfaces)
    ReDim Preserve ZOrder(NumSurfaces)
    ZOrder(NumSurfaces) = NumSurfaces
    Set mSurface(NumSurfaces) = New clsDXSurf
    With mSurface(NumSurfaces)
        Desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        Desc.ddsCaps.lCaps = DDSCAPS_TEXTURE
        Desc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        'Desc.ddckCKSrcBlt.High = RGB(255, 255, 255)
        'Desc.ddckCKSrcBlt.Low = RGB(255, 255, 255)
        Picture.ScaleMode = vbPixels
        Desc.lWidth = Picture.ScaleWidth
        Desc.lHeight = Picture.ScaleHeight
        Set .DDS7 = DD.CreateSurface(Desc)
        DC = .DDS7.GetDC
        BitBlt DC, 0, 0, Desc.lWidth, Desc.lHeight, Picture.hdc, 0, 0, vbSrcCopy
        .DDS7.ReleaseDC DC
        .Init NumSurfaces, DD
        .Trans = Trans
    End With
    CreateSurfaceFromPBox = NumSurfaces
End Function
Public Function CreateSolidColorSurface(Color As Long, Height As Long, Width As Long, Optional Trans As Long = -1)
    Dim DC As Long
    Dim r1 As RECT
    Dim Desc As DDSURFACEDESC2
    Dim C(3) As Byte
    NumSurfaces = NumSurfaces + 1
    ReDim Preserve mSurface(NumSurfaces)
    ReDim Preserve ZOrder(NumSurfaces)
    ZOrder(NumSurfaces) = NumSurfaces
    Set mSurface(NumSurfaces) = New clsDXSurf
    With mSurface(NumSurfaces)
        Desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        Desc.ddsCaps.lCaps = DDSCAPS_TEXTURE
        Desc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        'Desc.ddckCKSrcBlt.High = RGB(255, 255, 255)
        'Desc.ddckCKSrcBlt.Low = RGB(255, 255, 255)
        Desc.lWidth = Width
        Desc.lHeight = Height
        Set .DDS7 = DD.CreateSurface(Desc)
        r1.Bottom = Height
        r1.Right = Width
        CopyMemory ByVal VarPtr(C(0)), ByVal VarPtr(Color), 4
        .DDS7.BltColorFill r1, RGB(C(2), C(1), C(0))
        .Init NumSurfaces, DD
        .Trans = Trans
    End With
    CreateSolidColorSurface = NumSurfaces
End Function
Public Function DuplicateSurface(SurfaceID As Long) As Long
    Dim i As Long
    Dim DC1 As Long
    Dim DC2 As Long
    Dim R As RECT
    Dim Desc As DDSURFACEDESC2
    NumSurfaces = NumSurfaces + 1
    ReDim Preserve mSurface(NumSurfaces)
    ReDim Preserve ZOrder(NumSurfaces)
    ZOrder(NumSurfaces) = NumSurfaces
    Set mSurface(NumSurfaces) = New clsDXSurf
    With mSurface(NumSurfaces)
        mSurface(SurfaceID).DDS7.GetSurfaceDesc Desc
        Desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_CKSRCBLT
        Desc.ddsCaps.lCaps = DDSCAPS_TEXTURE
        Desc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
        Set .DDS7 = DD.CreateSurface(Desc)
        
        DC1 = .DDS7.GetDC
        DC2 = mSurface(SurfaceID).DDS7.GetDC
        BitBlt DC1, 0, 0, Desc.lWidth, Desc.lHeight, DC2, 0, 0, vbSrcCopy
        .DDS7.ReleaseDC DC1
        mSurface(SurfaceID).DDS7.ReleaseDC DC2
        
        .Init NumSurfaces, DD
        .Trans = mSurface(SurfaceID).Trans
    End With
    DuplicateSurface = NumSurfaces
End Function

Public Sub SetZOrder(SurfaceID As Long, Z As Long)
    Dim L As Long
    Dim X As Long
    Dim Y As Long
    Dim H As Long
    Y = UBound(ZOrder)
    For X = 0 To Y
        If ZOrder(X) = SurfaceID Then Exit For
    Next X
    
    If X = Y + 1 Then Exit Sub
    If X < Y Then
        L = (Y - X) * 4
        CopyMemory ByVal VarPtr(ZOrder(X)), ByVal VarPtr(ZOrder(X + 1)), L
    End If
    Z = Z - 1
    If Z = -1 Or Z >= Y Then
        ZOrder(Y) = SurfaceID
    Else
        L = (Y - Z) * 4
        CopyMemory ByVal VarPtr(ZOrder(Z + 1)), ByVal VarPtr(ZOrder(Z)), L
        ZOrder(Z) = SurfaceID
    End If

        
End Sub

Private Sub Class_Initialize()
    NumSurfaces = -1
End Sub

Private Sub Class_Terminate()
    Dim X As Long
    On Error Resume Next
    For X = 0 To NumSurfaces
        Set mSurface(X) = Nothing
    Next X
    cTmr.Remove TimerKey
    Set cTmr = Nothing
End Sub
Public Property Get Surface(ByVal Index As Long) As clsDXSurf
    Set Surface = mSurface(Index)
End Property

Public Function Animating() As Boolean
    Animating = (UBound(Anims) > 0)
End Function

Public Property Get TargetFrameRate() As Long
    TargetFrameRate = mTFR
End Property
Public Property Let TargetFrameRate(F As Long)
    mTFR = F
    cTmr.Interval(TimerKey) = Round(1000 / F)
End Property
Private Sub cTmr_Timer(ByVal sKey As String)
    Blt
End Sub
