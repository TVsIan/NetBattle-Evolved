Attribute VB_Name = "modPixFlip"
Option Explicit

' Palette Stuff
Private Declare Function GetNearestPaletteIndex Lib "GDI32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function GetPaletteEntries Lib "GDI32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function ResizePalette Lib "GDI32" (ByVal hPalette As Long, ByVal nNumEntries As Long) As Long
Private Declare Function SetPaletteEntries Lib "GDI32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function GetObject Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnResume As Long, ByRef ppstr As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32.dll" (ByVal lpStream As IUnknown, ByVal lSize As Long, ByVal fRunMode As Long, ByRef riid As Guid, ByRef lplpObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef pclsid As Guid) As Long
Private Const SIPICTURE As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private Const MAX_PALETTE_SIZE = 256
Private Const PC_NOCOLLAPSE = &H4    ' Do not match color existing entries.
Private Const NUMRESERVED = 106  ' Number of reserved entries in system palette.
Private Const SIZEPALETTE = 104  ' Size of system palette.
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type RGBTriplet 'BitmapArray Info
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type
Private Enum bmphErrors
    bmphListError = vbObjectError + 1001
    bmphPaletteError
End Enum
Private Type BITMAP ' Bitmap Stuff
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type



' Match control's palette to system palette.
Private Sub MatchColorPalette(ByVal Pic As PictureBox)
Dim log_hpal As Long
Dim sys_pal(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim orig_pal(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim i As Integer
Dim sys_pal_size As Long
Dim num_static_colors As Long
Dim static_color_1 As Long
Dim static_color_2 As Long

    ' Make sure pic has foreground palette.
    Pic.ZOrder
    RealizePalette Pic.hdc
    DoEvents

    ' Get system palette size and # static colors.
    sys_pal_size = GetDeviceCaps(Pic.hdc, SIZEPALETTE)
    num_static_colors = GetDeviceCaps(Pic.hdc, NUMRESERVED)
    static_color_1 = num_static_colors \ 2 - 1
    static_color_2 = sys_pal_size - num_static_colors \ 2

    ' Get system palette entries.
    GetSystemPaletteEntries Pic.hdc, 0, _
        sys_pal_size, sys_pal(0)

    ' Make logical palette as big as possible.
    log_hpal = Pic.Picture.hpal
    If ResizePalette(log_hpal, sys_pal_size) = 0 Then
        Err.Raise bmphListError, _
            "MatchColorPalette", _
            "Error matching bitmap palette"
    End If

    ' Blank non-static colors.
    For i = 0 To static_color_1
        orig_pal(i) = sys_pal(i)
    Next i
    For i = static_color_1 + 1 To static_color_2 - 1
        With orig_pal(i)
            .peRed = 0 'Set to black: RGB 0,0,0
            .peGreen = 0
            .peBlue = 0
            .peFlags = PC_NOCOLLAPSE
        End With
    Next i
    For i = static_color_2 To 255
        orig_pal(i) = sys_pal(i)
    Next i
    SetPaletteEntries log_hpal, 0, sys_pal_size, orig_pal(0)

    ' Insert non-static colors.
    For i = static_color_1 + 1 To static_color_2 - 1
        orig_pal(i) = sys_pal(i)
        orig_pal(i).peFlags = PC_NOCOLLAPSE
    Next i
    SetPaletteEntries log_hpal, static_color_1 + 1, static_color_2 - static_color_1 - 1, orig_pal(static_color_1 + 1)

    ' Realize new palette.
    RealizePalette Pic.hdc
End Sub

' Load PicBox bits into 2D array of RGB values. Set bits/pixel to number of bits/pixel.
Public Sub GetBitmapPixels(ByVal Pic As PictureBox, ByRef Pixels() As RGBTriplet, ByRef bits_per_pixel As Integer)

Dim hbm As Long
Dim bm As BITMAP
Dim L As Single
Dim T As Single
Dim old_color As Long
Dim bytes() As Byte
Dim num_pal_entries As Long
Dim pal_entries(0 To MAX_PALETTE_SIZE - 1) As PALETTEENTRY
Dim pal_index As Integer
Dim wid As Integer
Dim hgt As Integer
Dim X As Integer
Dim Y As Integer
Dim two_bytes As Long

    Pic.BackColor = &HC00000
    ' Get the bitmap information.
    hbm = Pic.Image
    GetObject hbm, Len(bm), bm
    bits_per_pixel = bm.bmBitsPixel

    ' If bits_per_pixel is 16, check for 15 or 16 bits per pixel.
    If bits_per_pixel = 16 Then
        ' Make the upper left pixel white.
        L = Pic.ScaleLeft
        T = Pic.ScaleTop
        old_color = Pic.Point(L, T)
        Pic.PSet (L, T), vbWhite

        ' Check Color
        ReDim bytes(0 To 0, 0 To 0)
        GetBitmapBits hbm, 2, bytes(0, 0)
        If (bytes(0, 0) And &H80) = 0 Then ' It's really a 15-bit image.
            bits_per_pixel = 15
        End If

        ' Restore pixel's color.
        Pic.PSet (L, T), old_color
    End If

    If (bits_per_pixel = 8) Or _
       (bits_per_pixel = 15) Or _
       (bits_per_pixel = 16) Or _
       (bits_per_pixel = 24) Or _
       (bits_per_pixel = 32) _
    Then
        ' Get the bits.
        ReDim bytes(0 To bm.bmWidthBytes - 1, 0 To bm.bmHeight - 1)
        GetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, bytes(0, 0)
    Else
        ' Oops! Not Readable...
        Err.Raise bmphListError, _
            "GetBitmapPixels", _
            "Invalid number of bits per pixel: " _
            & Format$(bits_per_pixel)
    End If
    
    ' Create pix array.
    wid = bm.bmWidth
    hgt = bm.bmHeight
    ReDim Pixels(0 To wid - 1, 0 To hgt - 1)
    Select Case bits_per_pixel
        Case 8
            ' Match pic's palette to system palette.
            MatchColorPalette Pic

            ' Get image's palette
            num_pal_entries = GetPaletteEntries( _
                Pic.Picture.hpal, 0, _
                MAX_PALETTE_SIZE, pal_entries(0))

            ' Get RGB color components.
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        pal_index = bytes(X, Y)
                        .rgbRed = pal_entries(pal_index).peRed
                        .rgbGreen = pal_entries(pal_index).peGreen
                        .rgbBlue = pal_entries(pal_index).peBlue
                    End With
                Next X
            Next Y

        Case 15
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        ' Get pixel's combined 2 bytes
                        two_bytes = bytes(X * 2, Y) + bytes(X * 2 + 1, Y) * 256&

                        ' Separate pixel's components
                        .rgbBlue = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbGreen = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbRed = two_bytes
                    End With
                Next X
            Next Y

        Case 16
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        ' Get pixel's combined 2 bytes
                        two_bytes = bytes(X * 2, Y) + bytes(X * 2 + 1, Y) * 256&

                        ' Separate pixel's components
                        .rgbBlue = two_bytes Mod 32
                        two_bytes = two_bytes \ 32
                        .rgbGreen = two_bytes Mod 64
                        two_bytes = two_bytes \ 64
                        .rgbRed = two_bytes
                    End With
                Next X
            Next Y

        Case 24
            ' Move pix array to bytes array via CopyMemory.
            For Y = 0 To hgt - 1
                CopyMemory Pixels(0, Y), bytes(0, Y), wid * 3
            Next Y

        Case 32
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        .rgbBlue = bytes(X * 4, Y)
                        .rgbGreen = bytes(X * 4 + 1, Y)
                        .rgbRed = bytes(X * 4 + 2, Y)
                    End With
                Next X
            Next Y

    End Select
End Sub
' Set PicBox bits using a 0-based 2D array of RGBTriplets
Public Sub SetBitmapPixels(ByVal Pic As PictureBox, ByVal bits_per_pixel As Integer, Pixels() As RGBTriplet)
Dim wid_bytes As Long
Dim wid As Integer
Dim hgt As Integer
Dim X As Integer
Dim Y As Integer
Dim bytes() As Byte
Dim hpal As Long
Dim two_bytes As Long

    ' Establish image size
    wid = UBound(Pixels, 1) + 1
    hgt = UBound(Pixels, 2) + 1

    ' Establish bytes per row needed
    Select Case bits_per_pixel
        Case 8
            wid_bytes = wid
        Case 15, 16
            wid_bytes = wid * 2
        Case 24
            wid_bytes = wid * 3
        Case 32
            wid_bytes = wid * 4
        Case Else
            ' Oops! Some weird bit/pixel...
            Err.Raise bmphListError, _
                "GetBitmapPixels", _
                "Invalid number of bits per pixel: " _
                & Format$(bits_per_pixel)
    End Select

    ' Make sure it's even.
    If wid_bytes Mod 2 = 1 Then wid_bytes = wid_bytes + 1

    ' Create bitmap bytes array.
    ReDim bytes(0 To wid_bytes - 1, 0 To hgt - 1)

    ' Set bitmap byte values.
    Select Case bits_per_pixel
        Case 8
            ' Use the nearest palette entries.
            hpal = Pic.Picture.hpal
            ' Get the RGB color components.
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        bytes(X, Y) = (&HFF And _
                            GetNearestPaletteIndex(hpal, _
                                RGB(.rgbRed, .rgbGreen, .rgbBlue) _
                            + &H2000000))
                    End With
                Next X
            Next Y

        Case 15
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        ' Validate in bounds.
                        If .rgbRed > &H1F Then .rgbRed = &H1F
                        If .rgbGreen > &H1F Then .rgbGreen = &H1F
                        If .rgbBlue > &H1F Then .rgbBlue = &H1F

                        ' Combine 2 byte values
                        two_bytes = .rgbBlue + 32 * (.rgbGreen + CLng(.rgbRed) * 32)

                        ' Set byte values.
                        bytes(X * 2, Y) = (two_bytes Mod 256) And &HFF
                        bytes(X * 2 + 1, Y) = (two_bytes \ 256) And &HFF
                    End With
                Next X
            Next Y
        Case 16
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        ' Validate in bounds.
                        If .rgbRed > &H1F Then .rgbRed = &H1F
                        If .rgbGreen > &H3F Then .rgbGreen = &H3F
                        If .rgbBlue > &H1F Then .rgbBlue = &H1F

                        ' Combine 2 byte values
                        two_bytes = .rgbBlue + 32 * (.rgbGreen + CLng(.rgbRed) * 64)

                        ' Set byte values.
                        bytes(X * 2, Y) = (two_bytes Mod 256) And &HFF
                        bytes(X * 2 + 1, Y) = (two_bytes \ 256) And &HFF

                    End With
                Next X
            Next Y
        Case 24
            ' Move pix array to bytes array via CopyMemory.
            For Y = 0 To hgt - 1
                CopyMemory bytes(0, Y), Pixels(0, Y), wid * 3
            Next Y
        Case 32
            For Y = 0 To hgt - 1
                For X = 0 To wid - 1
                    With Pixels(X, Y)
                        bytes(X * 4, Y) = .rgbBlue
                        bytes(X * 4 + 1, Y) = .rgbGreen
                        bytes(X * 4 + 2, Y) = .rgbRed
                    End With
                Next X
            Next Y
    End Select

    ' Set picture's bitmap bits.
    SetBitmapBits Pic.Image, wid_bytes * hgt, _
        bytes(0, 0)
    Pic.Refresh
End Sub
Public Sub CreateMask(InputBox As PictureBox, OutputBox As PictureBox)
    Dim Pixels() As RGBTriplet
    Dim mask_pixels() As RGBTriplet
    Dim bits_per_pixel As Integer
    Dim transparent_r As Byte
    Dim transparent_g As Byte
    Dim transparent_b As Byte
    Dim X As Integer
    Dim Y As Integer
    InputBox.ScaleMode = vbPixels
    OutputBox.ScaleMode = vbPixels
    OutputBox.Height = InputBox.Height
    OutputBox.Width = InputBox.Width
    
    ' Get pixels
    GetBitmapPixels InputBox, Pixels, bits_per_pixel

    ' Check upper left pixel's color - convert all pixels this color to white
    ' and all other values to black.
    With Pixels(0, 0)
        transparent_r = .rgbRed
        transparent_g = .rgbGreen
        transparent_b = .rgbBlue
    End With

    ' Allocate the mask pixels.
    ReDim mask_pixels( _
        LBound(Pixels, 1) To UBound(Pixels, 1), _
        LBound(Pixels, 2) To UBound(Pixels, 2))

    ' Set pixel color values.
    For Y = 0 To InputBox.ScaleHeight - 1
        For X = 0 To InputBox.ScaleWidth - 1
            With Pixels(X, Y)
                If (.rgbRed = transparent_r) And _
                   (.rgbGreen = transparent_g) And _
                   (.rgbBlue = transparent_b) _
                Then
                    ' Set pixels to white.
                    .rgbRed = 255
                    .rgbGreen = 255
                    .rgbBlue = 255
                    ' Make mask pixel white also...
                    mask_pixels(X, Y) = Pixels(X, Y)
                Else
                    ' Set pixels to black.
                    mask_pixels(X, Y).rgbRed = 0
                    mask_pixels(X, Y).rgbGreen = 0
                    mask_pixels(X, Y).rgbBlue = 0
                End If
            End With
        Next X
    Next Y

    ' Set Foreground's pixels.
    SetBitmapPixels InputBox, bits_per_pixel, Pixels
    InputBox.Picture = InputBox.Image

    ' Set Mask's pixels.
    SetBitmapPixels OutputBox, bits_per_pixel, mask_pixels
    OutputBox.Picture = OutputBox.Image
End Sub
Public Function GetYOffset(Pic As PictureBox)
    Dim Pixels() As RGBTriplet
    Dim mask_pixels() As RGBTriplet
    Dim bits_per_pixel As Integer
    Dim transparent_r As Byte
    Dim transparent_g As Byte
    Dim transparent_b As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Pic.ScaleMode = vbPixels
    GetBitmapPixels Pic, Pixels, bits_per_pixel
    With Pixels(0, 0)
        transparent_r = .rgbRed
        transparent_g = .rgbGreen
        transparent_b = .rgbBlue
    End With
    
    For Y = 0 To Pic.ScaleHeight - 1
        For X = 0 To Pic.ScaleWidth - 1
            With Pixels(X, Y)
                If Not ((.rgbRed = transparent_r) And _
                   (.rgbGreen = transparent_g) And _
                   (.rgbBlue = transparent_b)) _
                Then
                    Exit For
                End If
            End With
        Next X
        If X <> Pic.ScaleWidth Then Exit For
    Next Y
    Z = Y
    
    For Y = Pic.ScaleHeight - 1 To 0 Step -1
        For X = 0 To Pic.ScaleWidth - 1
            With Pixels(X, Y)
                If Not ((.rgbRed = transparent_r) And _
                   (.rgbGreen = transparent_g) And _
                   (.rgbBlue = transparent_b)) _
                Then
                    Exit For
                End If
            End With
        Next X
        If X <> Pic.ScaleWidth Then Exit For
        Z = Z - 1
    Next Y
    
    GetYOffset = Z
End Function
Public Sub SetTransPixels(InputBox As PictureBox, Color As Long)
    Dim Pixels() As RGBTriplet
    Dim mask_pixels() As RGBTriplet
    Dim bits_per_pixel As Integer
    Dim transparent_r As Byte
    Dim transparent_g As Byte
    Dim transparent_b As Byte
    Dim set_r As Byte
    Dim set_g As Byte
    Dim set_b As Byte
    Dim X As Integer
    Dim Y As Integer
    InputBox.ScaleMode = vbPixels
    GetBitmapPixels InputBox, Pixels, bits_per_pixel
    With Pixels(0, 0)
        transparent_r = .rgbRed
        transparent_g = .rgbGreen
        transparent_b = .rgbBlue
    End With
    set_r = Color Mod 256
    set_g = Color \ 256 Mod 256
    set_b = Color \ 65536
    
    For Y = 0 To InputBox.ScaleHeight - 1
        For X = 0 To InputBox.ScaleWidth - 1
            With Pixels(X, Y)
                If (.rgbRed = transparent_r) And _
                   (.rgbGreen = transparent_g) And _
                   (.rgbBlue = transparent_b) _
                Then
                    .rgbRed = set_r
                    .rgbGreen = set_g
                    .rgbBlue = set_b
                End If
            End With
        Next X
    Next Y
    SetBitmapPixels InputBox, bits_per_pixel, Pixels
    InputBox.Picture = InputBox.Image
End Sub

Public Sub PaintPictureTrans(Dest As PictureBox, Source As PictureBox, Mask As PictureBox, X As Single, Y As Single)
    Dest.PaintPicture Mask.Picture, X, Y, , , , , , , vbMergePaint
    Dest.PaintPicture Source.Picture, X, Y, , , , , , , vbSrcAnd
End Sub
Public Sub BitBltTrans(Dest As PictureBox, Source As PictureBox, Mask As PictureBox, X As Single, Y As Single)
    BitBlt Dest.hdc, X, Y, Mask.Width, Mask.Height, Mask.hdc, 0, 0, vbMergePaint
    BitBlt Dest.hdc, X, Y, Source.Width, Source.Height, Source.hdc, 0, 0, vbSrcAnd
End Sub
Public Function PictureFromArray(ByRef B() As Byte) As IPictureDisp
  Dim istrm As IUnknown
  Dim tGuid As Guid
  On Error GoTo ETrap
  If Not CreateStreamOnHGlobal(B(LBound(B)), False, istrm) Then
    CLSIDFromString StrPtr(SIPICTURE), tGuid
    OleLoadPicture istrm, UBound(B) - LBound(B) + 1, False, tGuid, PictureFromArray
  End If
  Set istrm = Nothing
  Exit Function
ETrap:
  'Debug.Print "Could not convert to IPicture!"
End Function
