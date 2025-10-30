Attribute VB_Name = "XORMod"
Const NCS = 256 'For easy typing =p
Const MAXLEN = NCS - 5
Const NCS64 = NCS / 64 - 1

Public Function XOREncrypt(ByVal Orig As String) As String
    Dim Checksum1 As Long
    Dim Checksum2 As Long
    Dim Checksum3 As Long
    Dim Working As Long
    Dim M1 As Integer
    Dim M2 As Integer
    Dim M3 As Integer
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim MainKey As Byte
    Dim iLen As Integer
    Dim iByte() As Byte
    Dim oByte() As Byte
    Dim tByte() As Byte
    Dim Build As String
    Dim Sect(3) As Byte
    Dim t As Double
        
    Orig = Left(Orig, MAXLEN)
    iLen = Len(Orig)
BeginEncrypt:
    ReDim iByte(1 To NCS)
    ReDim oByte(1 To NCS)
    ReDim tByte(1 To NCS)
    'Alright.  The first thing we do is convert the string
    'to a byte array for speedy manipulation.  The first Chr
    'is a length marker, and the rest of the array is filled
    'with random characters.
    iByte(1) = CByte(iLen)
    For X = 1 To iLen
        iByte(X + 1) = Asc(Mid$(Orig, X, 1))
    Next X
    For X = X + 1 To NCS
        iByte(X) = CByte(Int(Rnd * 256))
    Next X
    oByte = iByte
    
    'This generates the main key.
    
    '--------------------
    MainKey = CByte(Int(Rnd() * 256))
    '--------------------
    
    'This is where the encryption takes place
    oByte(NCS) = MainKey Xor 12
    For X = 1 To MAXLEN + 1
        oByte(X) = oByte(X) Xor oByte(MAXLEN + 2 - X) Xor MainKey
    Next X
    
    'And this part generates a few checksums
    Checksum1 = 0
    Checksum2 = 0
    Checksum3 = 0
    
    'The first is simply an addition of the original characters XORed
    'against the character in front of it.
    For X = 1 To iLen
        Checksum1 = Checksum1 + (iByte(X) Xor iByte(X + 1))
    Next X
    Checksum1 = Checksum1 Mod 256
    oByte(NCS - 3) = CByte(Checksum1)
    
    'The second is an addition of the each character in the encrypted
    'string XORed against a its position in the string.
    For X = 1 To NCS - 3
        Checksum2 = Checksum2 + (oByte(X) Xor (NCS - X))
    Next X
    Checksum2 = (Checksum2 + oByte(NCS)) Mod 256
    oByte(NCS - 2) = CByte(Checksum2)
    
    'And the third is too complicated to explain =p
    For X = 1 To 64
        Working = X
        For Y = 0 To NCS64
            If Y = NCS64 Then
                If X <> 63 Then
                    Working = Working + oByte(Y * 64 + X) * ((Y Mod 2 = 1) * 2 + 1)
                End If
            Else
                Working = Working + oByte(Y * 64 + X) * ((Y Mod 2 = 1) * 2 + 1)
            End If
        Next Y
        Checksum3 = Checksum3 + Working
    Next X
    Checksum3 = Abs(Int(Checksum3 + (oByte((MainKey Mod 128) + 1)))) Mod 256
    oByte(NCS - 1) = CByte(Checksum3)

    'And the finishing touch, the string is scrambled
    tByte = oByte
    Y = 0
    For X = 1 To NCS - 1 Step 2
        Y = Y + 1
        oByte(X) = tByte(Y)
    Next X
    Y = NCS + 1
    For X = 2 To NCS Step 2
        Y = Y - 1
        oByte(X) = tByte(Y)
    Next X
    
    'Finally, the string is reformed from the array, with one final ajustment
    Build = String$(NCS, " ")
    For X = 1 To NCS
        Mid$(Build, X, 1) = Chr$(oByte(X))
    Next X
    Build = StrReverse(Left(Build, NCS \ 2)) & StrReverse(Right(Build, NCS \ 2))
    Select Case Left(Build, 5)
    Case "REQN:", "RPWD:", "NAME:", "SPWD:", "BANU:", "NOIP:" 'Yes, yes, it's a billion to one chance I know, but just in case...
        GoTo BeginEncrypt
    End Select
    XOREncrypt = Build
End Function

Public Function XORDecrypt(ByVal EncryptedString As String) As String
    Dim Checksum1 As Long
    Dim Checksum2 As Long
    Dim Checksum3 As Long
    Dim Working As Long
    Dim M1 As Integer
    Dim M2 As Integer
    Dim M3 As Integer
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim A As Integer
    Dim MainKey As Byte
    Dim iByte() As Byte
    Dim oByte() As Byte
    Dim tByte() As Byte
    Dim Build As String
    Dim Sect(3) As Byte
    Dim t As Double
    
    ReDim iByte(NCS)
    ReDim oByte(NCS)
    If Len(EncryptedString) <> NCS Then
        XORDecrypt = "XXXXX"
        Exit Function
    End If
    Build = EncryptedString
    Build = StrReverse(Left(Build, NCS \ 2)) & StrReverse(Right(Build, NCS \ 2))
    For X = 1 To NCS
        iByte(X) = CByte(Asc(Mid$(Build, X, 1)))
    Next X
    
    Y = 0
    For X = 1 To NCS - 1 Step 2
        Y = Y + 1
        oByte(Y) = iByte(X)
    Next X
    Y = NCS + 1
    For X = 2 To NCS Step 2
        Y = Y - 1
        oByte(Y) = iByte(X)
    Next X
    
    iByte = oByte
    MainKey = oByte(NCS) Xor 12
    
    For X = MAXLEN + 1 To 1 Step -1
        oByte(X) = oByte(MAXLEN + 2 - X) Xor oByte(X) Xor MainKey
    Next X
    Z = oByte(1)
        
    For X = 1 To Z
        Checksum1 = Checksum1 + (oByte(X) Xor oByte(X + 1))
    Next X
    Checksum1 = Checksum1 Mod 256

    For X = 1 To NCS - 3
        Checksum2 = Checksum2 + (iByte(X) Xor (NCS - X))
    Next X
    Checksum2 = (Checksum2 + iByte(NCS)) Mod 256
    
    For X = 1 To 64
        Working = X
        For Y = 0 To NCS64
            If Y = NCS64 Then
                If X <> 63 Then
                    Working = Working + iByte(Y * 64 + X) * ((Y Mod 2 = 1) * 2 + 1)
                End If
            Else
                Working = Working + iByte(Y * 64 + X) * ((Y Mod 2 = 1) * 2 + 1)
            End If
        Next Y
        Checksum3 = Checksum3 + Working
    Next X
    Checksum3 = Abs(Int(Checksum3 + (iByte((MainKey Mod 128) + 1)))) Mod 256
    
    X = 0
    If iByte(NCS - 3) <> Checksum1 Then X = 1
    If iByte(NCS - 2) <> Checksum2 Then X = 1
    If iByte(NCS - 1) <> Checksum3 Then X = 1
    If Z < 5 Or Z > MAXLEN + 1 Then X = 1
    If X = 0 Then
        Build = String$(Z, " ")
        For X = 1 To Z
            Mid$(Build, X, 1) = Chr$(oByte(X + 1))
        Next X
    Else
        Build = "XXXXX"
    End If
    XORDecrypt = Build
End Function
