Attribute VB_Name = "XORMod"
Option Explicit
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)
Private Const MaxLen = 252

Public Function FormatPacket(ByVal Data As String, ByVal Encrypt As Boolean) As String
    Dim X As Integer
    If Encrypt Then Data = XOREncrypt(Data)
    Data = Left$(Data, 256)
    Data = Chr$(Len(Data) - 1) & Data
    FormatPacket = Data
End Function
Public Function GetPacket(ByRef Socket As Winsock, ByVal BytesTotal As Long, ByRef PacketBuffer() As String) As Boolean
    Dim Temp As String
    Dim Packet() As String
    Dim X As Long
    Dim Y As Long
    Dim R As Long
    On Error GoTo Failed
    ReDim Packet(0)
    Y = 0
    Socket.GetData Temp, , BytesTotal
    Temp = Socket.Tag & Temp
    Socket.Tag = ""
    ''>>> Call WriteDebugLog("Packet Rcv: " & Str2Hex(Temp))
    R = Len(Temp)
    Do Until R = 0
        R = R - 1
        X = Asc(Left$(Temp, 1)) + 1
        If R < X Then
            '>>> Call WriteDebugLog("Missing Data, Saving")
            Socket.Tag = Temp
            Exit Do
        End If
        Temp = Right$(Temp, R)
        R = R - X
        Y = Y + 1
        ReDim Preserve Packet(Y)
        Packet(Y) = Left$(Temp, X)
        Temp = Right$(Temp, R)
    Loop
    PacketBuffer = Packet
    GetPacket = True
    Exit Function
Failed:
    '>>> Call WriteDebugLog("Error " & Err.Number & " - " & Err.Description)
    GetPacket = False
End Function
Public Sub RawInterpret(Raw As String)
    Dim Temp As String
    Dim Packet() As String
    Dim X As Integer
    Dim Y As Integer
    Dim R As Integer
    On Error GoTo Failed
    ReDim Packet(0)
    Y = -1
    Temp = Raw
    '>>> Call WriteDebugLog("Packet Rcv: " & Str2Hex(Temp))
    R = Len(Temp)
    While R <> 0
        R = R - 1
        X = Asc(Left$(Temp, 1)) + 1
        Temp = Right$(Temp, R)
        If R < X Then
            '>>> Call WriteDebugLog("Invalid Len Marker")
            GoTo Failed
        End If
        R = R - X
        Y = Y + 1
        ReDim Preserve Packet(Y)
        Packet(Y) = Left$(Temp, X)
        Temp = Right$(Temp, R)
    Wend
    Exit Sub
Failed:
    '>>> Call WriteDebugLog("Error " & Err.Number & " - " & Err.Description)
End Sub
Public Function XOREncrypt(ByVal Orig As String) As String
    Dim MainKey As Byte
    Dim iLen As Integer
    Dim pLen As Integer
    Dim iByte() As Byte
    Dim oByte() As Byte
    Dim tByte() As Byte
    Dim Checksum1 As Long
    Dim Checksum2 As Long
    Dim Checksum3 As Long
    Dim Working As Long
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Build As String
    On Error GoTo ETCTrap
    
    Orig = Left(Orig, MaxLen)
    iLen = Len(Orig)
    If iLen Mod 2 = 1 Then
        Orig = vbNullChar & Orig
        iLen = iLen + 1
    End If
    pLen = iLen + 4
    ReDim iByte(1 To pLen)
    ReDim oByte(1 To pLen)
    ReDim tByte(1 To pLen)
    'Alright.  The first thing we do is convert the string
    'to a byte array for speedy manipulation.  The first Chr
    'is a length marker, and the rest of the array is filled
    'with random characters.
    CopyMemory iByte(1), ByVal Orig, Len(Orig)
'    For X = 1 To iLen
'        iByte(X) = Asc(Mid$(Orig, X, 1))
'    Next X
    oByte = iByte
    
    'This generates the main key.
    
    '--------------------
    MainKey = CByte(Int(Rnd() * 256))
    '--------------------
    
    'This is where the encryption takes place
    oByte(pLen) = MainKey Xor 12
    For X = 1 To iLen
        oByte(X) = oByte(X) Xor oByte(iLen - X + 1) Xor MainKey
    Next X
    
    'And this part generates a few checksums
    Checksum1 = 0
    Checksum2 = 0
    Checksum3 = 0
    
    'The first is simply an addition of the original characters XORed
    'against the character in front of it.
    For X = 1 To iLen - 1
        Checksum1 = Checksum1 + (iByte(X) Xor iByte(X + 1))
    Next X
    Checksum1 = Checksum1 Mod 256
    oByte(pLen - 3) = CByte(Checksum1)
    
    'The second is an addition of the each character in the encrypted
    'string XORed against a its position in the string.
    For X = 1 To pLen - 3
        Checksum2 = Checksum2 + (oByte(X) Xor (256 - X))
    Next X
    Checksum2 = (Checksum2 + MainKey) Mod 256
    oByte(pLen - 2) = CByte(Checksum2)
    
    'And the third is too complicated to explain =p
    V = pLen \ 4
    For X = 1 To V
        Working = X
        For Y = 0 To 3
            Z = Y * V + X
            If Z <> pLen - 1 Then
                Working = Working + oByte(Z) * ((Y Mod 2 = 1) * 2 + 1) + 7
            End If
        Next Y
        Checksum3 = Checksum3 + Working
    Next X
    Checksum3 = Abs(Int(Checksum3 + (oByte((MainKey Mod iLen) + 1)))) Mod 256
    oByte(pLen - 1) = CByte(Checksum3)

    'And the finishing touch, the string is scrambled
    For Z = 1 To 2
        tByte = oByte
        Y = 0
        For X = 1 To pLen - 1 Step 2
            Y = Y + 1
            oByte(X) = tByte(Y)
        Next X
        Y = pLen + 1
        For X = 2 To pLen Step 2
            Y = Y - 1
            oByte(X) = tByte(Y)
        Next X
    Next Z
    
    'Finally, the string is reformed from the array, with one final ajustment
    Build = String$(pLen, vbNullChar)
'    For X = 1 To pLen
'        Mid$(Build, X, 1) = Chr$(oByte(X))
'    Next X
    CopyMemory ByVal Build, oByte(1), ByVal pLen
    Build = StrReverse(Left(Build, pLen \ 2)) & StrReverse(Right(Build, pLen \ 2))
    Select Case Left(Build, 5)
    Case "REQN:", "RPWD:", "NAME:", "SPWD:", "BANU:", "NOIP:" 'Yes, yes, it's a billion to one chance I know, but just in case...
        XOREncrypt = XOREncrypt(Orig)
    Case Else
        XOREncrypt = Build
    End Select
    Exit Function
ETCTrap:
    If Err.Number = 16 Then Resume Else Err.Raise Err.Number
End Function

Public Function XORDecrypt(ByVal EncryptedString As String, Optional SkipChecksums As Boolean = False) As String
    Dim MainKey As Byte
    Dim iByte() As Byte
    Dim oByte() As Byte
    Dim tByte() As Byte
    Dim Build As String
    Dim iLen As Integer
    Dim pLen As Integer
    Dim Checksum1 As Long
    Dim Checksum2 As Long
    Dim Checksum3 As Long
    Dim Working As Long
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Build = EncryptedString
    pLen = Len(Build)
    If pLen Mod 2 = 1 Then
        '>>> Call WriteDebugLog("DECRYPT ERROR: Odd length packet rcvd")
        '>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(Build))
        XORDecrypt = "XXXXX"
        Exit Function
    End If
    If Len(Build) < 5 Then
        '>>> Call WriteDebugLog("DECRYPT ERROR: Len(Build) < 5")
        '>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(Build))
        XORDecrypt = "XXXXX"
        Exit Function
    End If
    iLen = pLen - 4
    ReDim iByte(1 To pLen)
    ReDim oByte(1 To pLen)
    Build = StrReverse(Left(Build, pLen \ 2)) & StrReverse(Right(Build, pLen \ 2))
    CopyMemory iByte(1), ByVal Build, Len(Build)
    
    For Z = 1 To 2
        Y = 0
        For X = 1 To pLen - 1 Step 2
            Y = Y + 1
            oByte(Y) = iByte(X)
        Next X
        Y = pLen + 1
        For X = 2 To pLen Step 2
            Y = Y - 1
            oByte(Y) = iByte(X)
        Next X
        iByte = oByte
    Next Z
    
    MainKey = oByte(pLen) Xor 12
    For X = iLen To 1 Step -1
        oByte(X) = oByte(X) Xor oByte(iLen - X + 1) Xor MainKey
    Next X
    
    Checksum1 = 0
    Checksum2 = 0
    Checksum3 = 0
    
    For X = 1 To iLen - 1
        Checksum1 = Checksum1 + (oByte(X) Xor oByte(X + 1))
    Next X
    Checksum1 = Checksum1 Mod 256
    
    For X = 1 To pLen - 3
        Checksum2 = Checksum2 + (iByte(X) Xor (256 - X))
    Next X
    Checksum2 = (Checksum2 + MainKey) Mod 256
    
    V = pLen \ 4
    For X = 1 To V
        Working = X
        For Y = 0 To 3
            Z = Y * V + X
            If Z <> pLen - 1 Then
                Working = Working + iByte(Z) * ((Y Mod 2 = 1) * 2 + 1) + 7
            End If
        Next Y
        Checksum3 = Checksum3 + Working
    Next X
    Checksum3 = Abs(Int(Checksum3 + (iByte((MainKey Mod iLen) + 1)))) Mod 256
    
    X = 0
    If iByte(pLen - 3) <> Checksum1 Then X = 1
    If iByte(pLen - 2) <> Checksum2 Then X = 1
    If iByte(pLen - 1) <> Checksum3 Then X = 1
    If X = 0 Or SkipChecksums Then
        Build = String$(iLen, vbNullChar)
        For X = 1 To iLen
            Mid$(Build, X, 1) = Chr$(oByte(X))
        Next X
        If Mid(Build, 1, 1) = vbNullChar Then Build = Right$(Build, Len(Build) - 1)
    Else
        '>>> Call WriteDebugLog("DECRYPT ERROR: Checksum mismatch.")
        '>>> Call WriteDebugLog("HEXED ESTRING: " & Str2Hex(EncryptedString))
'        If iByte(pLen - 3) <> Checksum1 Then '>>> Call WriteDebugLog("CHECK1 FAIL")
'        If iByte(pLen - 2) <> Checksum2 Then '>>> Call WriteDebugLog("CHECK2 FAIL")
'        If iByte(pLen - 1) <> Checksum3 Then '>>> Call WriteDebugLog("CHECK3 FAIL")
        Build = "XXXXX"
    End If
    XORDecrypt = Build
End Function
Public Sub Benchmark(Seconds As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Long
    Dim T As Single
    Dim E As String
    T = Timer + Seconds
    While Timer < T
        Y = Int(Rnd * 250) + 2
        E = RndStr(Y)
        If E <> XORDecrypt(XOREncrypt(E)) Then Stop
        Z = Z + Y
    Wend
    MsgBox CStr(Round(Z / 1024 / Seconds, 2)) & " kilobytes per second.", vbInformation, "Result"
End Sub
Public Function RndStr(StrLen As Integer) As String
    Dim X As Integer
    Dim B As Byte
    Dim Build As String
    Build = String$(StrLen, vbNullChar)
    For X = 1 To StrLen
        Do
            B = Int(Rnd * 256)
        Loop Until X <> 1 Or B <> 0
        Mid(Build, X) = Chr$(B)
    Next X
    RndStr = Build
End Function
'Public Function Str2Hex(Orig As String) As String
'    Dim X As Integer
'    Dim Temp As String
'    Temp = String$(Len(Orig) * 3, vbNullChar)
'    For X = 1 To Len(Orig)
'        Mid(Temp, X * 3 - 2) = FixedHex(Asc(Mid(Orig, X, 1)), 2)
'    Next X
'    Str2Hex = Temp
'End Function
'Public Function Hex2Str(ByVal Hexed As String) As String
'    Dim H() As String
'    Dim X As Integer
'    Dim Temp As String
'    If Len(Hexed) > 2 And Mid(Hexed, 3, 1) <> " " Then
'        For X = 1 To Len(Hexed) Step 2
'            Temp = Temp & Mid(Hexed, X, 2) & " "
'        Next X
'        Hexed = Temp
'    End If
'    H = Split(Trim(Hexed), " ")
'    For X = 0 To UBound(H)
'        H(X) = Chr$(Dec(H(X)))
'    Next X
'    Hex2Str = Join(H, vbNullString)
'End Function
Public Function Rebuild(ByteArray() As Byte) As String
    Dim X As Integer
    Dim Temp As String
    For X = LBound(ByteArray) To UBound(ByteArray)
        Temp = Temp & Chr$(ByteArray(X))
    Next X
    Rebuild = Temp
    
End Function

Public Function Deconstruct(TheString As String) As Byte()
    Dim X As Long
    Dim Temp As String
    Dim B() As Byte
    ReDim B(0 To Len(TheString) - 1)
    For X = 0 To Len(TheString) - 1
        B(X) = Asc(Mid(TheString, X + 1, 1))
    Next X
    Deconstruct = B
End Function
