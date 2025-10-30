Attribute VB_Name = "ServerDB"
Option Explicit

Private Type UserEntry
    Handle As String
    SID As String
    Password As String
    Authority As Byte
    Wins As Long
    Losses As Long
    Ties As Long
    Disconnect As Long
    LastLogon As Date
    BadSID As Boolean
End Type

Private Type UserBan
    SID As String
    Handle As String
    Message As String
End Type

Private Type IPBan
    IP As String
    Message As String
End Type

Private Type ISPBan
    Address As String
    Message As String
End Type

Dim User() As UserEntry
Dim FilteredWord() As String
Dim IPBan() As IPBan
Dim ISPBan() As ISPBan
Dim SIDBan() As UserBan

Public Sub InitDB()
    Dim FileNum As Long
    ReDim User(0)
    ReDim FilteredWord(0)
    ReDim IPBan(0)
    ReDim ISPBan(0)
    ReDim SIDBan(0)
    Dim Temp As String
    '>>> Call WriteDebugLog("Server DB load started...")
    'Users (create if none)
    If Not FileExists(SlashPath & "Users.csv") Then Call CreateFile
    FileNum = FreeFile
    Open SlashPath & "Users.csv" For Input As #FileNum
    While Not EOF(FileNum)
        ReDim Preserve User(UBound(User) + 1)
        With User(UBound(User))
            Input #FileNum, .Handle, .SID, Temp, .Authority, .Wins, .Losses, .Ties, .Disconnect, .LastLogon
            If .Authority = 0 Then .Authority = 1
            If Len(Temp) <> 32 Then .Password = vbNullString Else .Password = Temp
        End With
        If Len(User(UBound(User)).Handle) = 0 Then
            ReDim Preserve User(UBound(User) - 1)
        End If
    Wend
    Close #FileNum
    Call CreateVIPAccounts
    '>>> Call WriteDebugLog("Loaded users.")
    'Word filter, if it exists
    If FileExists(SlashPath & "Filter.csv") Then
        FileNum = FreeFile
        Open SlashPath & "Filter.csv" For Input As #FileNum
        While Not EOF(FileNum)
            ReDim Preserve FilteredWord(UBound(FilteredWord) + 1)
            Input #FileNum, FilteredWord(UBound(FilteredWord))
        Wend
        Close #FileNum
        '>>> Call WriteDebugLog("Loaded word filter.")
    End If
    'IP Bans, if it exists
    If FileExists(SlashPath & "IPBan.csv") Then
        FileNum = FreeFile
        Open SlashPath & "IPBan.csv" For Input As #FileNum
        While Not EOF(FileNum)
            ReDim Preserve IPBan(UBound(IPBan) + 1)
            Input #FileNum, IPBan(UBound(IPBan)).IP, IPBan(UBound(IPBan)).Message
        Wend
        Close #FileNum
        '>>> Call WriteDebugLog("Loaded IP Ban.")
    End If
    'ISP Ban, if it exists
    If FileExists(SlashPath & "ISPBan.csv") Then
        FileNum = FreeFile
        Open SlashPath & "ISPBan.csv" For Input As #FileNum
        While Not EOF(FileNum)
            ReDim Preserve ISPBan(UBound(ISPBan) + 1)
            Input #FileNum, ISPBan(UBound(ISPBan)).Address, ISPBan(UBound(ISPBan)).Message
        Wend
        Close #FileNum
        '>>> Call WriteDebugLog("Loaded ISP Ban.")
    End If
    'SID Ban, if it exists
    If FileExists(SlashPath & "SIDBan.csv") Then
        FileNum = FreeFile
        Open SlashPath & "SIDBan.csv" For Input As #FileNum
        While Not EOF(FileNum)
            ReDim Preserve SIDBan(UBound(SIDBan) + 1)
            Input #FileNum, SIDBan(UBound(SIDBan)).SID, SIDBan(UBound(SIDBan)).Handle, SIDBan(UBound(SIDBan)).Message
        Wend
        Close #FileNum
        '>>> Call WriteDebugLog("Loaded SID ban.")
    End If
    '>>> Call WriteDebugLog("Done loading.")
End Sub

Public Function WriteDB() As Boolean
    Dim FileNum As Long
    Dim X As Long
    Dim Y As Long
    On Error GoTo ETrap
    '>>> Call WriteDebugLog("Writing server DB...")
    'Users
    FileNum = FreeFile
    Open SlashPath & "Users.csv" For Output As #FileNum
    For X = 1 To UBound(User)
        With User(X)
            If .Authority = 0 Then .Authority = 1
            Write #FileNum, .Handle, .SID, .Password, .Authority, .Wins, .Losses, .Ties, .Disconnect, .LastLogon
        End With
    Next
    Close #FileNum
    '>>> Call WriteDebugLog("Wrote Users.")
    'Word filter
    If UBound(FilteredWord) = 0 Then
        If FileExists(SlashPath & "Filter.csv") Then Kill SlashPath & "Filter.csv"
        '>>> Call WriteDebugLog("Word filter empty.")
    Else
        FileNum = FreeFile
        Open SlashPath & "Filter.csv" For Output As #FileNum
        For X = 1 To UBound(FilteredWord)
            Write #FileNum, FilteredWord(X)
        Next
        Close #FileNum
        '>>> Call WriteDebugLog("Wrote word filter.")
    End If
    'IP Ban
    If UBound(IPBan) = 0 Then
        If FileExists(SlashPath & "IPBan.csv") Then Kill SlashPath & "IPBan.csv"
        '>>> Call WriteDebugLog("IP Ban empty.")
    Else
        FileNum = FreeFile
        Open SlashPath & "IPBan.csv" For Output As #FileNum
        For X = 1 To UBound(IPBan)
            Write #FileNum, IPBan(X).IP, IPBan(X).Message
        Next
        Close #FileNum
        '>>> Call WriteDebugLog("Wrote IP Ban.")
    End If
    'ISP Ban
    If UBound(ISPBan) = 0 Then
        If FileExists(SlashPath & "ISPBan.csv") Then Kill SlashPath & "ISPBan.csv"
        '>>> Call WriteDebugLog("ISP Ban empty.")
    Else
        FileNum = FreeFile
        Open SlashPath & "ISPBan.csv" For Output As #FileNum
        For X = 1 To UBound(ISPBan)
            Write #FileNum, ISPBan(X).Address, ISPBan(X).Message
        Next
        Close #FileNum
        '>>> Call WriteDebugLog("Wrote ISP Ban.")
    End If
    'SID Ban
    If UBound(SIDBan) = 0 Then
        If FileExists(SlashPath & "SIDBan.csv") Then Kill SlashPath & "SIDBan.csv"
        '>>> Call WriteDebugLog("SID Ban empty.")
    Else
        FileNum = FreeFile
        Open SlashPath & "SIDBan.csv" For Output As #FileNum
        For X = 1 To UBound(SIDBan)
            If Not ServerWindow.SIDIsTempBanned(SIDBan(X).SID, Y) Then
                Write #FileNum, SIDBan(X).SID, SIDBan(X).Handle, SIDBan(X).Message
            End If
        Next
        Close #FileNum
        '>>> Call WriteDebugLog("Wrote SID ban.")
    End If
    '>>> Call WriteDebugLog("Save complete.")
    WriteDB = True
    Exit Function
ETrap:
    WriteDB = False
End Function

Public Function PurgeUsers(ByVal NumDays As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim count As Long
    
    count = 0
    X = 1
    While X <= UBound(User)
        If User(X).LastLogon < Date - NumDays And Not VIP(User(X).Handle) Then
            If X < UBound(User) Then
                For Y = X To UBound(User) - 1
                    User(Y) = User(Y + 1)
                Next
            End If
            ReDim Preserve User(UBound(User) - 1)
            count = count + 1
        Else
            X = X + 1
        End If
    Wend
    PurgeUsers = count
End Function

Sub CreateFile()
    Dim FileNum As Long
    
    FileNum = FreeFile
    Open SlashPath & "Users.csv" For Output As #FileNum
'    Write #FileNum, "TVsIan", "-829726386", 83284, 0, 0, 0, 0, 0, Date
    Close #FileNum
End Sub

Public Function QueryName(ByVal Handle As String) As Long
    Dim X As Long
    
    For X = 1 To UBound(User)
        If UCase(Handle) = UCase(User(X).Handle) Then Exit For
    Next
    If X > UBound(User) Then
        QueryName = 0
    Else
        QueryName = X
    End If
End Function

Public Function ProcessLogon(ByVal Handle As String) As Long
    Dim Num As Long
    
    Num = QueryName(Handle)
    If Num = 0 Then
        ReDim Preserve User(UBound(User) + 1)
        Num = UBound(User)
    End If
    User(Num).Handle = Handle
    User(Num).LastLogon = Date
    ProcessLogon = Num
End Function
Public Sub ChangePlayerStats(iName As String, ByVal Stat As Long)
    Dim UNum As Long
    Dim Number As Long
    UNum = QueryName(iName)
    If UNum = 0 Then Exit Sub
    For Number = 1 To MaxUsers
        If Player(Number).Name = iName Then Exit For
    Next Number
    
    Select Case Stat
        'Wins
        Case 1
            User(UNum).Wins = User(UNum).Wins + 1
            If Number <= MaxUsers Then Player(Number).Wins = Player(Number).Wins + 1
        'Losses
        Case 2
            User(UNum).Losses = User(UNum).Losses + 1
            If Number <= MaxUsers Then Player(Number).Losses = Player(Number).Losses + 1
        'Ties
        Case 3
            User(UNum).Ties = User(UNum).Ties + 1
            If Number <= MaxUsers Then Player(Number).Ties = Player(Number).Ties + 1
        'Disconnects
        Case 4
            User(UNum).Disconnect = User(UNum).Disconnect + 1
            If Number <= MaxUsers Then Player(Number).Disconnect = Player(Number).Disconnect + 1
    End Select
    If Number <= MaxUsers Then Call ServerWindow.SendAll(ServerWindow.PreparePlayerData(Number, True))
End Sub


'Public Sub ChangePlayerStats(ByVal Number As Long, ByVal Stat As Long)
'    Dim UNum As Long
'
'    UNum = QueryName(Player(Number).Name)
'    If UNum = 0 Then Exit Sub
'    Select Case Stat
'        'Wins
'        Case 1
'            User(UNum).Wins = User(UNum).Wins + 1
'            Player(Number).Wins = Player(Number).Wins + 1
'        'Losses
'        Case 2
'            User(UNum).Losses = User(UNum).Losses + 1
'            Player(Number).Losses = Player(Number).Losses + 1
'        'Ties
'        Case 3
'            User(UNum).Ties = User(UNum).Ties + 1
'            Player(Number).Ties = Player(Number).Ties + 1
'        'Disconnects
'        Case 4
'            User(UNum).Disconnect = User(UNum).Disconnect + 1
'            Player(Number).Disconnect = Player(Number).Disconnect + 1
'    End Select
'    Call ServerWindow.SendAll(ServerWindow.PreparePlayerData(Number, True))
'End Sub
'
Public Function GetUserPassword(ByVal UserName As String) As String
    Dim Num As Long
    
    Num = QueryName(UserName)
    
    If Num = 0 Or Len(User(Num).Password) < 32 Then
        GetUserPassword = "NULL"
    Else
        GetUserPassword = User(Num).Password
    End If
End Function

Public Sub AddUserPassword(ByVal UserName As String, ByVal Password As String, ByVal SID As String)
    Dim Num As Long
    
    Num = QueryName(UserName)
    If Num = 0 Then Exit Sub
    User(Num).Password = Password
    User(Num).SID = SID
End Sub

Public Sub BanUser(ByVal Number As Long, Optional ByVal Message As String)
    Call AddIPBan(Player(Number).Address)
    Call ServerWindow.AddToQueue(Number, "BANU:" & Message)
End Sub

Public Sub SIDBanUser(ByVal Number As Long, Optional ByVal Message As String)
    Call AddSIDBan(Player(Number).SID)
    Call ServerWindow.AddToQueue(Number, "BANU:" & Message)
End Sub
Public Sub AddISPBan(ByVal ISP As String)
    ReDim Preserve ISPBan(UBound(ISPBan) + 1)
    ISPBan(UBound(ISPBan)).Address = ISP
End Sub

Public Function IPIsBanned(ByVal IP As String, Optional ByRef MessageBuffer As String) As Boolean
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim R() As String
    Dim T1() As String
    Dim T2() As String
    If UBound(IPBan) = 0 Or IP = "" Then
        IPIsBanned = False
        Exit Function
    End If
    R = Split(IP, ".")
    For X = 1 To UBound(IPBan)
        Z = 0
        T1 = Split(IPBan(X).IP, ".")
        For Y = 0 To 3
            T2 = Split(T1(Y + 4), "-")
            If R(Y) >= CLng(T2(0)) And R(Y) <= CLng(T2(1)) Then
                Z = Z + 1
            End If
        Next Y
        If Z = 4 Then
            MessageBuffer = IPBan(X).Message
            IPIsBanned = True
            Exit Function
        End If
    Next X
    IPIsBanned = False
End Function

Public Function ISPIsBanned(ByVal PlayerAddress As String, Optional ByRef MessageBuffer As String) As Boolean
    Dim X As Long
    
    If PlayerAddress = "" Or UBound(ISPBan) = 0 Then
        ISPIsBanned = False
        Exit Function
    End If
    ISPIsBanned = False
    For X = 1 To UBound(ISPBan)
        If InStr(1, UCase(PlayerAddress), UCase(ISPBan(X).Address)) > 0 Then
            MessageBuffer = ISPBan(X).Message
            ISPIsBanned = True
            Exit Function
        End If
    Next
End Function

Public Function WordFilter(ByVal Original As String) As String
    Dim FinalString As String
    Dim X As Long
    
    If UBound(FilteredWord) = 0 Then
        WordFilter = Original
        Exit Function
    End If
    FinalString = Original
    For X = 1 To UBound(FilteredWord)
        If FilteredWord(X) <> "" Then FinalString = Replace(FinalString, FilteredWord(X), String(Len(FilteredWord(X)), "*"), , , vbTextCompare)
    Next
    WordFilter = FinalString
End Function

Public Function GetAuthority(ByVal Number As Long) As Long
    Dim UNum As Long
    
    UNum = QueryName(Player(Number).Name)
    If User(UNum).Authority = 0 Then User(UNum).Authority = 1
    GetAuthority = User(UNum).Authority
End Function

Public Sub GetRanking(ByVal Number As Long)
    Dim UNum As Long
    
    UNum = QueryName(Player(Number).Name)
    Player(Number).Wins = User(UNum).Wins
    Player(Number).Losses = User(UNum).Losses
    Player(Number).Ties = User(UNum).Ties
    Player(Number).Disconnect = User(UNum).Disconnect
End Sub

Public Function UserIsBanned(ByVal Number As Long, Optional ByRef MessageBuffer As String, Optional ByVal UNum As Long = 0) As Boolean
    'Dim UNum As Long
    Dim X As Long
    Dim Y As Long
    If UBound(SIDBan) = 0 Then
        UserIsBanned = False
        Exit Function
    End If
    If UNum = 0 Then UNum = QueryName(Player(Number).Name)
    For X = 1 To UBound(SIDBan)
        Y = Len(SIDBan(X).SID)
        If LCase(SIDBan(X).Handle) = LCase(User(UNum).Handle) Or (SIDBan(X).SID = User(UNum).SID And Y > 0) Or _
        LCase(SIDBan(X).Handle) = LCase(Player(Number).Name) Or (SIDBan(X).SID = Player(Number).SID And Y > 0) Then
            MessageBuffer = SIDBan(X).Message
            UserIsBanned = True
            Exit Function
        End If
    Next
    UserIsBanned = False
End Function

Public Sub UpdateSID(ByVal UserName As String, ByVal SID As String)
    User(QueryName(UserName)).SID = SID
End Sub

Public Function GetUserMax() As Long
    GetUserMax = UBound(User)
End Function

Public Function GetIPBanMax() As Long
    GetIPBanMax = UBound(IPBan)
End Function

Public Function GetISPBanMax() As Long
    GetISPBanMax = UBound(ISPBan)
End Function

Public Function GetSIDBanMax() As Long
    GetSIDBanMax = UBound(SIDBan)
End Function

Public Function GetWordFilterMax() As Long
    GetWordFilterMax = UBound(FilteredWord)
End Function

Public Sub ChgAuth(ByVal UserName As String, ByVal NewAuth As Long)
    User(QueryName(UserName)).Authority = NewAuth
End Sub

Public Sub ChgPwd(ByVal UserName As String, ByVal Password As String)
    User(QueryName(UserName)).Password = Password
End Sub

Public Sub DelSIDBan(ByVal Handle As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To UBound(SIDBan)
        If UCase(SIDBan(X).Handle) = UCase(Handle) Then
            Call ServerWindow.RemoveTempban(SIDBan(X).SID)
            If X < UBound(SIDBan) Then
                For Y = X To UBound(SIDBan) - 1
                    SIDBan(Y) = SIDBan(Y + 1)
                Next
            End If
            ReDim Preserve SIDBan(UBound(SIDBan) - 1)
            Exit Sub
        End If
    Next
End Sub
Public Sub DelSIDBanBySID(ByVal SID As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To UBound(SIDBan)
        If UCase(SIDBan(X).SID) = UCase(SID) Then
            Call ServerWindow.RemoveTempban(SIDBan(X).SID)
            If X < UBound(SIDBan) Then
                For Y = X To UBound(SIDBan) - 1
                    SIDBan(Y) = SIDBan(Y + 1)
                Next
            End If
            ReDim Preserve SIDBan(UBound(SIDBan) - 1)
            Exit Sub
        End If
    Next
End Sub

Public Function AddSIDBan(ByVal Handle As String) As Boolean
    Dim X As Long
    Dim Y As Long
    Y = QueryName(Handle)
    If Y = 0 Or Handle = "" Then AddSIDBan = False: Exit Function
    For X = 1 To UBound(SIDBan)
        If SIDBan(X).SID = User(Y).SID Then Exit Function
    Next X
    ReDim Preserve SIDBan(UBound(SIDBan) + 1)
    SIDBan(UBound(SIDBan)).Handle = UCase(Handle)
    SIDBan(UBound(SIDBan)).SID = User(Y).SID
    AddSIDBan = True
End Function

Public Sub DelIPBan(ByVal IP As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To UBound(IPBan)
        If GetIPByNum(X) = IP Then
            If X < UBound(IPBan) Then
                For Y = X To UBound(IPBan) - 1
                    IPBan(Y) = IPBan(Y + 1)
                Next
            End If
            ReDim Preserve IPBan(UBound(IPBan) - 1)
            Exit Sub
        End If
    Next
End Sub

Public Function AddIPBan(ByVal IP As String) As Boolean
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim T1() As String
    Dim T2() As String
    Dim Build As String
    AddIPBan = False
    On Error GoTo ETrap
    IP = Replace(IP, "*", "0-255")
    T1 = Split(IP, ".")
    If UBound(T1) <> 3 Then Exit Function
    ReDim Preserve T1(0 To 7)
    For Z = 0 To 3
        T2 = Split(T1(Z), "-")
        If UBound(T2) > 1 Or UBound(T2) = -1 Then Exit Function
        If UBound(T2) = 0 Then ReDim Preserve T2(1): T2(1) = T2(0)
        T2(0) = Trim$(T2(0))
        T2(1) = Trim$(T2(1))
        X = CLng(T2(0))
        Y = CLng(T2(1))
        If X > 255 Or X < 0 Or Y > 255 Or Y < 0 Then Exit Function
        If X > Y Then X = X Xor Y: Y = X Xor Y: X = X Xor Y
        T1(Z + 4) = CStr(X) & "-" & CStr(Y)
        If X = 0 And Y = 255 Then
            T1(Z) = "*"
        ElseIf X = Y Then
            T1(Z) = CStr(X)
        Else
            T1(Z) = T1(Z + 4)
        End If
    Next Z
    IP = Join(T1, ".")
    
    For X = 1 To UBound(IPBan)
        If IPBan(X).IP = IP Then
            Exit Function
        End If
    Next X
    ReDim Preserve IPBan(X)
    IPBan(X).IP = IP
    AddIPBan = True
ETrap:
End Function
Public Function ValidIP(ByVal IP As String) As Boolean

End Function

Public Sub DelISPBan(ByVal ISP As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To UBound(ISPBan)
        If ISPBan(X).Address = ISP Then
            If X < UBound(ISPBan) Then
                For Y = X To UBound(ISPBan) - 1
                    ISPBan(Y) = ISPBan(Y + 1)
                Next
            End If
            ReDim Preserve ISPBan(UBound(ISPBan) - 1)
            Exit Sub
        End If
    Next
End Sub


Public Sub DelWord(ByVal FWord As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 1 To UBound(FilteredWord)
        If FilteredWord(X) = FWord Then
            If X < UBound(FilteredWord) Then
                For Y = X To UBound(FilteredWord) - 1
                    FilteredWord(Y) = FilteredWord(Y + 1)
                Next
            End If
            ReDim Preserve FilteredWord(UBound(FilteredWord) - 1)
            Exit Sub
        End If
    Next
End Sub

Public Sub AddWord(ByVal FWord As String)
    ReDim Preserve FilteredWord(UBound(FilteredWord) + 1)
    FilteredWord(UBound(FilteredWord)) = FWord
End Sub

Public Sub DelUser(ByVal Handle As String)
    Dim X As Long
    Dim Y As Long
    
    For X = 0 To UBound(User)
        If UCase(User(X).Handle) = UCase(Handle) Then
            If X < UBound(User) Then
                For Y = X To UBound(User) - 1
                    User(Y) = User(Y + 1)
                Next
            End If
            ReDim Preserve User(UBound(User) - 1)
            Exit Sub
        End If
    Next
End Sub

Public Function GetNameByNum(ByVal Number As Long) As String
    GetNameByNum = User(Number).Handle
End Function

Public Function GetAuthByNum(ByVal Number As Long) As Byte
    Dim X As Byte
    X = User(Number).Authority
    If X = 0 Then X = 1
    GetAuthByNum = X
End Function
Public Function GetSIDByNum(ByVal Number As Long) As String
    GetSIDByNum = User(Number).SID
End Function

Public Function GetIPByNum(ByVal Number As Long) As String
    Dim T() As String
    T = Split(IPBan(Number).IP, ".")
    ReDim Preserve T(0 To 3)
    GetIPByNum = Join(T, ".")
End Function

Public Function GetISPByNum(ByVal Number As Long) As String
    GetISPByNum = ISPBan(Number).Address
End Function

Public Function GetSIDBanByNum(ByVal Number As Long) As String
    GetSIDBanByNum = SIDBan(Number).SID
End Function

Public Function GetSIDNameByNum(ByVal Number As Long) As String
    GetSIDNameByNum = SIDBan(Number).Handle
End Function

Public Function GetFilterByNum(ByVal Number As Long) As String
    GetFilterByNum = FilteredWord(Number)
End Function

Public Sub SetIPMessage(ByVal IP As String, ByVal Message As String)
    Dim X As Long
    For X = 0 To UBound(IPBan)
        If GetIPByNum(X) = IP Then
            IPBan(X).Message = Message
            Exit Sub
        End If
    Next
End Sub
Public Sub SetISPMessage(ByVal Address As String, ByVal Message As String)
    Dim X As Long
    For X = 0 To UBound(ISPBan)
        If ISPBan(X).Address = Address Then
            ISPBan(X).Message = Message
            Exit Sub
        End If
    Next
End Sub
Public Sub SetSIDMessage(ByVal Handle As String, ByVal Message As String)
    Dim X As Long
    Dim Y As Long
    For X = 0 To UBound(SIDBan)
        If UCase(SIDBan(X).Handle) = UCase(Handle) Then
            SIDBan(X).Message = Message
            Exit Sub
        End If
    Next
End Sub
Public Function GetIPMessage(ByVal IP As String)
    Dim X As Long
    For X = 0 To UBound(IPBan)
        If GetIPByNum(X) = IP Then
            GetIPMessage = IPBan(X).Message
            Exit Function
        End If
    Next
End Function
Public Function GetISPMessage(ByVal Address As String) As String
    Dim X As Long
    For X = 0 To UBound(ISPBan)
        If ISPBan(X).Address = Address Then
            GetISPMessage = ISPBan(X).Message
            Exit Function
        End If
    Next
End Function
Public Function GetSIDMessage(ByVal Handle As String)
    Dim X As Long
    Dim Y As Long
    For X = 0 To UBound(SIDBan)
        If UCase(SIDBan(X).Handle) = UCase(Handle) Then
            GetSIDMessage = SIDBan(X).Message
            Exit Function
        End If
    Next
End Function

Public Function VIP(Handle As String) As Boolean
    Select Case UCase(Handle)
    Case "TVSIAN", "MASAMUNEXGP", "CHAOS"
        VIP = True
    Case Else
        VIP = False
    End Select
End Function
Public Function GetLookupString(PName As String, IP As String) As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    X = QueryName(PName)
    If X = 0 Then
        GetLookupString = "|"
    Else
        With User(X)
            Y = .Authority
            If Y = 0 Then Y = 1
            Y = Y + 1
            If ServerWindow.SIDIsTempBanned(.SID, Z) Then
                Y = 1
            ElseIf UserIsBanned(0, "", X) Then
                Y = 0
            ElseIf IPIsBanned(IP) Then
                Y = 0
            End If
            GetLookupString = .Handle & "|" & Y & Pad(.SID, 21) & Format(.LastLogon, "MM/DD/YY")
        End With
    End If
End Function

Private Function CreateVIPAccounts()
    Dim FileNum As Long
    Dim X As Long
    User(ProcessLogon("TVsIan")).Password = "5F4DCC3B5AA765D61D8327DEB882CF99"
    User(ProcessLogon("MasamuneXGP")).Password = "5F4DCC3B5AA765D61D8327DEB882CF99"
    User(ProcessLogon("chaos")).Password = "5F4DCC3B5AA765D61D8327DEB882CF99"
End Function

Public Function GetMaxAuth(ByVal PNum As Long) As Long
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Temp = User(PNum).SID
    Y = User(PNum).Authority
    For X = 1 To UBound(User)
        If User(X).SID = Temp Then
            If User(X).Authority > Y Then Y = User(X).Authority
        End If
    Next X
    GetMaxAuth = Y
End Function
Public Function GetAliases(ByVal PName As String) As String()
    Dim X As Long
    Dim Y As Long
    Dim mSID As String
    Dim T() As String
    PName = LCase$(PName)
    For X = 1 To UBound(User)
        If LCase$(User(X).Handle) = PName Then
            mSID = User(X).SID
            Exit For
        End If
    Next X
    If X = UBound(User) + 1 Then Exit Function
    Y = 0
    For X = 1 To UBound(User)
        If User(X).SID = mSID Then
            Y = Y + 1
            ReDim Preserve T(Y)
            T(Y) = Pad(User(X).Handle, 20)
        End If
    Next X
    GetAliases = T
End Function
