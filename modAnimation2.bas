Attribute VB_Name = "modAnimation"
Option Explicit
'This module will handle animations
Const iPath = "C:\Pokemon\Animation Notes\Named Pics\"
Public Type AnimType
    Index As Long
    ID As Long
    StartFrame As Long
    FirstSurf As Long
    P1 As Long
    P2 As Long
    T() As Single
End Type
Public Enum AnimIndex
    nbaMove = 0
    nbaSwitch = 400
    nbaWeather = 500
    nbaStatus = 600
End Enum
    
Public Sub DoAnim(Window As Battle, ByVal ID As AnimIndex, Optional ByVal P1 As Long, Optional ByVal P2 As Long, Optional ByVal P3 As Long, Optional ByVal P4 As Long)
    Dim F As Long
    Dim X As Long
    If Window Is Nothing Then Exit Sub
    If Window.DX Is Nothing Then Exit Sub
    If Not UseDX Then Exit Sub
    With Window.DX
        F = .CurrentFrame + 1
        Select Case ID
        Case nbaMove + GetMoveNum("Ice Beam")
            Call .AddAnim(18, F, P1, P2)  'Moving Particles
            Call .AddAnim(19, F + 99, P1, P2) 'Flickering Ice
        Case nbaMove + GetMoveNum("Thunderbolt")
            Call .AddAnim(20, F, P1, P2)  'Falling bolts
            Call .AddAnim(21, F + 79, P1, P2) 'Big spark
        Case nbaMove + GetMoveNum("Flamethrower")
            Call .AddAnim(22, F, P1, P2) 'Flames
        Case nbaMove + GetMoveNum("Surf")
            If TeamNum(P1) = Window.PNum Then X = 1 Else X = 2
            Call .AddAnim(23, F, X, 0)
        Case nbaMove + GetMoveNum("Earthquake")
            Call .AddAnim(24, F, Window.ActNum, 0)
        
        Case nbaMove + GetMoveNum("Metronome")
            Call .AddAnim(5025, F, P1, CLng(TeamNum(P1) = Window.PNum))
        Case nbaMove + GetMoveNum("Sunny Day"), nbaWeather + nbSunny
            Call .AddAnim(5026, F)
            
            
            
        Case nbaSwitch 'Switch Retract
        'P1 = AtkPos, P2 = PokeBall Type
            'P2 = 7
            Call .AddAnim(5, F, P1, P2)
            Call .AddAnim(P2 + 6, F, P1, 0)
            While .CurrentFrame < 35 And .Animating
                Sleep 1
                DoEvents
            Wend
            Exit Sub
        Case nbaSwitch + 1 'Switch Send Out
        'P1 = AtkPos, P2 = PokeBall Type
            'P2 = 7
            Call Window.UpdateImages
            Call Window.UpdateStats
            .Surface(P1).Visible = False
            Call .AddAnim(1, F, P1, (TeamNum(P1) = Window.PNum))
            If TeamNum(P1) = Window.PNum Then
                F = F + 45
            Else
                F = F + 15
            End If
            Call .AddAnim(4, F, P1, P2)
            Call .AddAnim(P2 + 6, F, P1, 0)
        Case nbaSwitch + 2 'Doubles start
        'P1 = Team, P2 = PokeBall Type #1, P3 = PokeBall Type #2
            If P1 = Window.PNum Then
                Call .AddAnim(1, F, P1, True)
                Call .AddAnim(1, F, P1 + 2, True)
                Call .AddAnim(4, F + 45, P1, P2)
                Call .AddAnim(P2 + 6, F + 45, P1, 0)
                Call .AddAnim(4, F + 73, P1 + 2, P3)
                Call .AddAnim(P3 + 6, F + 73, P1 + 2, 0)
            Else
                Call .AddAnim(1, F, P1, 0)
                Call .AddAnim(1, F, P1 + 2, 2)
                Call .AddAnim(4, F + 15, P1, P2)
                Call .AddAnim(P2 + 6, F + 15, P1, 0)
                Call .AddAnim(4, F + 43, P1 + 2, P3)
                Call .AddAnim(P3 + 6, F + 43, P1 + 2, 0)
            End If
        End Select
        
        While .Animating
            Sleep 1
            DoEvents
        Wend
    End With
    
    
End Sub

Public Sub DoAnimKernel(DX As clsDX, A As AnimType)
    Const PI As Double = 3.14159265358979
    Dim S As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim sngT As Single
    Dim CenterX As Long
    Dim CenterY As Long
    Dim Temp As String
    With DX
        F = .CurrentFrame - A.StartFrame
        S = A.FirstSurf
        If F < 0 Then Exit Sub
        Select Case A.ID
        Case 1 'Pokeball throw
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "PokeBall0.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                .Surface(S).Framed = True
            End If
            With .Surface(S)
                If A.P2 = 0 Or A.P2 = 2 Then
                    Select Case F
                    Case 0
                        .Left = DX.Surface(A.P1).sngX + (DX.Surface(A.P1).Width - .Width) \ 2
                        .Top = DX.Surface(A.P1).sngY + DX.Surface(A.P1).Height - (.Height \ 2)
                    Case 16 + A.P2 * 14
                        .scrTop = 16
                    Case 20 + A.P2 * 14
                        .scrTop = 32
                    Case 24 + A.P2 * 14
                        DX.AnimFinished A.Index
                    End Select
                    If A.P1 < 3 Then
                        DX.Surface(5).Visible = False
                    Else
                        DX.Surface(6).Visible = False
                    End If
                Else
                    If F > 8 Then .AngleOfRotation = .AngleOfRotation + 0.62
                    Select Case DX.Surface(A.P1).sngX
                    Case Is < 40 'Doubles/Left
                        Select Case F
                        Case 0: .Move 16, 60
                        Case 1: .Move 16, 58
                        Case 2: .Move 16, 56
                        Case 3: .Move 16, 55
                        Case 4: .Move 17, 53
                        Case 5: .Move 17, 52
                        Case 6: .Move 17, 51
                        Case 7: .Move 18, 51
                        Case 8: .Move 19, 51
                        Case 15: .Move 20, 52
                        Case 19: .Move 20, 53
                        Case 21: .Move 20, 54
                        Case 22: .Move 20, 55
                        Case 23: .Move 20, 56
                        Case 25: .Move 21, 57
                        Case 26: .Move 21, 58
                        Case 28: .Move 21, 59
                        Case 29: .Move 21, 60
                        Case 31: .Move 21, 62
                        Case 32: .Move 21, 63
                        Case 33: .Move 21, 65
                        Case 34: .Move 22, 68
                        Case 35: .Move 22, 72
                        Case 36: .Move 22, 77
                        Case 37: .Move 23, 82
                        Case 38: .Move 23, 87
                        Case 39: .Move 23, 92
                        Case 40: .Move 23, 97
                        Case 41: .Move 24, 102
                        Case 42: .Move 24, 108
                        Case 56
                            .Visible = False
                            DX.AnimFinished A.Index
                        End Select
                    Case Is < 50 'Singles
                        Select Case F
                        Case 0: .Move 16, 60
                        Case 1: .Move 17, 57
                        Case 2: .Move 19, 55
                        Case 3: .Move 21, 54
                        Case 4: .Move 23, 52
                        Case 5: .Move 25, 50
                        Case 6: .Move 27, 49
                        Case 7: .Move 29, 49
                        Case 8: .Move 31, 49
                        Case 10: .Move 32, 49
                        Case 11: .Move 32, 49
                        Case 12: .Move 33, 49
                        Case 13: .Move 34, 49
                        Case 14: .Move 34, 49
                        Case 15: .Move 35, 50
                        Case 16: .Move 36, 50
                        Case 17: .Move 36, 51
                        Case 18: .Move 37, 51
                        Case 20: .Move 38, 52
                        Case 21: .Move 39, 52
                        Case 22: .Move 39, 53
                        Case 23: .Move 40, 53
                        Case 24: .Move 41, 54
                        Case 26: .Move 42, 55
                        Case 27: .Move 43, 55
                        Case 28: .Move 43, 56
                        Case 29: .Move 44, 56
                        Case 30: .Move 44, 57
                        Case 31: .Move 46, 58
                        Case 32: .Move 47, 59
                        Case 33: .Move 47, 61
                        Case 34: .Move 49, 63
                        Case 35: .Move 51, 67
                        Case 36: .Move 53, 72
                        Case 37: .Move 55, 76
                        Case 38: .Move 57, 80
                        Case 39: .Move 59, 86
                        Case 40: .Move 61, 91
                        Case 41: .Move 63, 97
                        Case 42: .Move 65, 104
                        Case 43: .Move 66, 108
                        Case 56
                            .Visible = False
                            DX.AnimFinished A.Index
                        End Select
                    Case Else 'Doubles/Right
                        Select Case F
                        Case 0: .Move 16, 60
                        Case 1: .Move 18, 59
                        Case 2: .Move 21, 57
                        Case 3: .Move 23, 57
                        Case 4: .Move 26, 55
                        Case 5: .Move 29, 54
                        Case 6: .Move 31, 54
                        Case 7: .Move 34, 54
                        Case 8: .Move 36, 54
                        Case 9: .Move 37, 54
                        Case 10: .Move 38, 55
                        Case 11: .Move 39, 55
                        Case 12: .Move 39, 55
                        Case 13: .Move 40, 55
                        Case 14: .Move 41, 56
                        Case 15: .Move 42, 55
                        Case 16: .Move 43, 55
                        Case 17: .Move 44, 56
                        Case 18: .Move 45, 56
                        Case 19: .Move 45, 57
                        Case 20: .Move 46, 58
                        Case 21: .Move 47, 58
                        Case 22: .Move 48, 59
                        Case 23: .Move 49, 60
                        Case 24: .Move 50, 61
                        Case 25: .Move 51, 62
                        Case 26: .Move 52, 62
                        Case 27: .Move 52, 63
                        Case 28: .Move 53, 64
                        Case 29: .Move 54, 65
                        Case 30: .Move 55, 66
                        Case 31: .Move 56, 67
                        Case 32: .Move 57, 68
                        Case 33: .Move 58, 70
                        Case 34: .Move 60, 73
                        Case 35: .Move 63, 78
                        Case 36: .Move 67, 83
                        Case 37: .Move 68, 88
                        Case 38: .Move 71, 93
                        Case 39: .Move 73, 98
                        Case 40: .Move 76, 103
                        Case 41: .Move 79, 110
                        Case 56
                            .Visible = False
                            DX.AnimFinished A.Index
                        End Select
                    End Select
                End If
            End With
        Case 2 'Options test
            If F = 0 Then
                MainContainer.DoPicture "006fl.gif"
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                With .Surface(S)
                    .sngY = 0.005
                    .Left = 10
                    .Top = 10
                End With
            End If
            With .Surface(S)
                .sngX = .sngX + .sngY
                If .sngX > 1 Or .sngX < 0 Then .sngY = .sngY * -1
                .sngX = .sngX + .sngY
                .AlphaBlend = .sngX
                .AngleOfRotation = (F Mod 100) * PI / 50
            End With
        Case 4 'Pokemon Emerage
            If F = 0 Then
                'Background overlay surface
                A.FirstSurf = .CreateSolidColorSurface(vbWhite, .Surface(0).Height, .Surface(0).Width)
                S = A.FirstSurf
                .SetZOrder S, 2
                .Surface(S).AlphaBlend = 0
                
                'Pokmeon overlay surface
                Call .CreateSolidColorSurface(GetPokeballColor(A.P2), .Surface(0).Height, .Surface(0).Width)
                
                With .Surface(S + 1)
                    .ClipSurface = DX.Surface(A.P1)
                    .Height = 14
                    .Width = 14
                    .Top = DX.Surface(A.P1).sngY + DX.Surface(A.P1).Height - 14
                    .Left = DX.Surface(A.P1).sngX + (DX.Surface(A.P1).Width - 14) \ 2
                    .sngX = .Left
                    
                    Temp = DX.Surface(A.P1).Tag
                    .Tag = "0"
                    If InStr(1, Temp, "b") = 0 And (InStr(1, Temp, "rs") > 0 Or InStr(1, Temp, "fl") > 0 Or Left(Temp, 5) = "unown") Then
                        .Tag = CStr(BasePKMN(CLng(Left$(Temp, 3))).Offset)
                    End If
                    
                End With
                .SetZOrder S + 1, S
            End If
            With .Surface(S + 1)
                Select Case F
                Case 1 To 10
                    .sngX = .sngX - 2.5
                    .Move Round(.sngX), .Top - 5, .Width + 5, .Height + 5
                    .sngY = .Top
                    If F > 1 Then DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend + 0.07692
                Case 11
                    DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend + 0.07692
                    X = CLng(.Tag)
                    If X > 0 Then
                        If A.P1 < 3 Then
                            DX.Surface(5).Visible = True
                        Else
                            DX.Surface(6).Visible = True
                        End If
                    End If
                    X = .Top - X
                    .Tag = CStr(X)
                    .sngY = .sngY - 1.333
                    .Top = Round(.sngY)
                    If .Top < X Then .Top = X
                Case 12 To 22
                    X = CLng(.Tag)
                    If F < 14 Then DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend + 0.07692
                    .sngY = .sngY - 1.333
                    .Top = Round(.sngY)
                    If .Top < X Then .Top = X
                Case 23 To 35
                    DX.Surface(A.P1).Visible = True
                    DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend - 0.07692
                    .AlphaBlend = .AlphaBlend - 0.06666
                Case 36 To 38
                    .AlphaBlend = .AlphaBlend - 0.06666
                Case 40
                    DX.AnimFinished A.Index
                End Select
            End With
        Case 5 'Pokemon Retreat
            If F = 0 Then
                'Background overlay surface
                A.FirstSurf = .CreateSolidColorSurface(vbWhite, .Surface(0).Height, .Surface(0).Width)
                S = A.FirstSurf
                .SetZOrder S, 2
                .Surface(S).AlphaBlend = 0
                'Pokmeon overlay surface
                
                Call .CreateSolidColorSurface(GetPokeballColor(A.P2), .Surface(0).Height, .Surface(0).Width)
                Call .DuplicateSurface(A.P1)
                Call .SetZOrder(S + 1, 0)
                With .Surface(S + 1)
                    .ClipSurface = DX.Surface(A.P1)
                    With DX.Surface(A.P1)
                        DX.Surface(S + 1).Move .Left, .Top, .Width, .Height
                        DX.Surface(S + 2).Move .Left, .Top, .Width, .Height
                    End With
                    .sngX = .Left
                    .AlphaBlend = 0
                End With
                .Surface(A.P1).Visible = False
            End If
            
            With .Surface(S + 1)
                Select Case F
                Case 1 To 13
                    .AlphaBlend = .AlphaBlend + 0.0625
                    DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend + 0.0625
                Case 14 To 23
                    If F < 17 Then
                        .AlphaBlend = .AlphaBlend + 0.0625
                        DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend + 0.0625
                    End If
                    .sngX = .sngX + 2.5
                    .Move Round(.sngX), .Top + 5, .Width - 5, .Height - 5
                    DX.Surface(S + 2).Move .Left, .Top, .Width, .Height
                Case 24
                    .Visible = False
                    DX.Surface(S + 2).Visible = False
                    DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend - 0.0625
                Case 25 To 40
                    DX.Surface(S).AlphaBlend = DX.Surface(S).AlphaBlend - 0.0625
                    .AlphaBlend = .AlphaBlend - 0.06666
                Case 41
                    DX.AnimFinished A.Index
                End Select
            End With
        Case 6 'Pokeball Effect (two groups of eight lines flying outward)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "linething.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 15
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 15
                    With .Surface(S + X)
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 5 Then
                If F > 13 Then X = 15 Else X = 7
                For X = 0 To X
                    With .Surface(S + X)
                        .sngY = .sngY + 2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                        Y = .scrTop \ 8
                        Do
                            Z = Int(Rnd * 4)
                        Loop Until Z <> Y
                        .scrTop = Z * 8
                    End With
                Next X
            End If
            Select Case F
            Case 13: .Surface(S + 3).Visible = True: .Surface(S + 4).Visible = True: .Surface(S + 5).Visible = True
            Case 31:
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
            Case 39:
                For X = 7 To 15
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished A.Index
            End Select
        Case 7: 'Great Ball Effect (Two spinning circles of lines)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "linething.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 15
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 15
                    With .Surface(S + X)
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                If F > 14 Then Y = 15 Else Y = 7
                For X = 0 To Y
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 16
                        .sngY = .sngY + 2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                        Y = .scrTop \ 8
                        Do
                            Z = Int(Rnd * 4)
                        Loop Until Z <> Y
                        .scrTop = Z * 8
                    End With
                Next X
            End If
            Select Case F
            Case 50
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
            Case 56
                For X = 7 To 15
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End Select
        Case 8 'Ultra Ball Effect (One spinning circle of green dots)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "greenspark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .TintBlue = 0.5
                        .TintRed = 0.5
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                        .scrTop = 48
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 9: 'Master Ball Effect (two spinning ovals of stars)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "star.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 15
                    .DuplicateSurface S
                Next X
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 15
                    With .Surface(S + X)
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 2 Then
                For X = 0 To 15
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        If X < 8 Then
                            .Top = CenterY - Sin(.sngX) * .sngY - Sin(.sngX) * 25
                            .Left = CenterX + Cos(.sngX) * .sngY
                        Else
                            .Top = CenterY - Sin(.sngX) * .sngY
                            .Left = CenterX + Cos(.sngX) * .sngY + Cos(.sngX) * 25
                        End If
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 15
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 10: 'Dive Ball Effect (Spinning Oval of bubbles)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "bubble.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        .Top = CenterY - Sin(.sngX) * .sngY - Sin(.sngX) * 25
                        .Left = CenterX + Cos(.sngX) * .sngY
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 11 'Nest Ball (Spinning circle of hearts)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "heart.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 12 'Timer Ball Effect (Spinning oval of green dots)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "greenspark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .TintBlue = 0.5
                        .TintRed = 0.5
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                        .scrTop = 48
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY + Cos(.sngX) * 35
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished A.Index
            End If
        Case 13 'Net Ball (Spinning circle of bubbles)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "bubble.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 18
                        .sngY = .sngY + 1.2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                    End With
                Next X
            End If
            If F = 40 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 14 'Safari Ball Effect (spinning circle of lines)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "linething.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                Y = .Surface(S).scrTop \ 8
                Do
                    Z = Int(Rnd * 4)
                Loop Until Z <> Y
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 16
                        .sngY = .sngY + 1
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                        .scrTop = Z * 8
                    End With
                Next X
            End If
            If F = 50 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 15: 'Luxury Ball Effect (Two spinning circles of morphing green sparkles)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "greenspark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 15
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 15
                    With .Surface(S + X)
                        .TintBlue = 0.5
                        .TintRed = 0.5
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                        .scrTop = 48
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                If F > 14 Then Y = 15 Else Y = 7
                For X = 0 To Y
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 16
                        .sngY = .sngY + 2
                        .Top = CenterY - Sin(.sngX) * .sngY
                        .Left = CenterX + Cos(.sngX) * .sngY
                        Select Case F Mod 8
                        Case 0: If X < 8 Then .scrTop = 48
                        Case 1: If X >= 8 Then .scrTop = 48
                        Case 4: If X < 8 Then .scrTop = 32
                        Case 5: If X >= 8 Then .scrTop = 32
                        End Select
                    End With
                Next X
            End If
            Select Case F
            Case 50
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
            Case 56
                For X = 7 To 15
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End Select
        Case 16 'Premier Ball Effect (Spiral of morphing green sparkles)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "greenspark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .TintBlue = 0.5
                        .TintRed = 0.5
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * PI / 4
                        .scrTop = 48
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + PI / 20
                        .sngY = .sngY + 0.8
                        If Sgn(Sin(.sngX)) = 1 And Sgn(Cos(.sngX)) = -1 Then
                            .Top = CenterY - Sin(.sngX) * .sngY * Cos(.sngX) * -1
                            .Left = CenterX + Cos(.sngX) * .sngY * Cos(.sngX) * -1
                        Else
                            .Top = CenterY - Sin(.sngX) * .sngY
                            .Left = CenterX + Cos(.sngX) * .sngY
                        End If
                        Select Case F Mod 8
                        Case 0: If X < 8 Then .scrTop = 48
                        Case 4: If X < 8 Then .scrTop = 32
                        End Select
                    End With
                Next X
            End If
            If F = 50 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
        Case 17 'Repeat Ball Effect (Infinity-shape of green dots)
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "greenspark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                .Surface(S).Framed = True
                CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
                CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
                For X = 0 To 7
                    With .Surface(S + X)
                        .TintBlue = 0.5
                        .TintRed = 0.5
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngY = 0
                        .sngX = X * 0.25
                        .scrTop = 48
                    End With
                Next X
            End If
            CenterX = .Surface(A.P1).sngX + (.Surface(A.P1).Width - .Surface(S).Width) \ 2
            CenterY = .Surface(A.P1).sngY + .Surface(A.P1).Height - (.Surface(S).Height \ 2)
            If F > 4 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        .sngX = .sngX + 0.0625
                        Do While .sngX > 1
                            .sngX = .sngX - 1
                        Loop
                        .sngY = (1 - (.sngX * 1) ^ 4) * (F - 4)
                        .Top = CenterY - Sin(.sngX) * .sngY
                        If X > 4 Then .Left = CenterX + Cos(.sngX) * .sngY
                        If X < 5 Then .Left = CenterX - Cos(.sngX) * .sngY
                    End With
                Next X
            End If
            If F = 50 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished (A.Index)
            End If
            
        '******* ICE BEAM *******
        Case 18 'Ice Particles
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "IceParticleSmall.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 20
                    .DuplicateSurface S
                Next X
                
                .CreateSolidColorSurface RGB(157, 205, 215), 64, 64
                .Surface(S + 21).ClipSurface = .Surface(A.P2)
                .Surface(S + 21).AlphaBlend = 0
                
                .Surface(S).Framed = True
                .Surface(S).scrHeight = 16
                With .Surface(A.P1)
                    CenterX = .sngX - DX.Surface(S).Width / 2
                    CenterY = .Top + .WhitespaceTop + (.Height - .WhitespaceTop - DX.Surface(S).Height) / 2
                End With
                With .Surface(A.P2)
                    X = .sngX + (.Width - DX.Surface(S).Width) \ 2
                    Y = .Top + .WhitespaceTop + (.Height - .WhitespaceTop) / 2 - DX.Surface(S).Height / 2
                End With
                If X > CenterX Then
                    .Surface(S + 21).sngX = .Surface(A.P1).Width - .Surface(A.P1).WhitespaceRight
                    CenterX = CenterX + .Surface(S + 21).sngX
                Else
                    .Surface(S + 21).sngX = .Surface(A.P1).WhitespaceLeft
                    CenterX = CenterX + .Surface(S + 21).sngX
                End If
                ReDim A.T(0)
                A.T(0) = Sqr((CenterY - Y) ^ 2 + (X - CenterX) ^ 2)
                .Surface(S).sngX = Atn((CenterY - Y) / (X - CenterX))
                If CenterX > X Then .Surface(S).sngX = .Surface(S).sngX + PI
                
                For X = 0 To 20
                    With .Surface(S + X)
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngX = DX.Surface(S).sngX
                        .sngY = 0
                        If X < 4 Then
                            .scrHeight = 16
                        Else
                            .scrTop = 16
                        End If
                        .AlphaBlend = 0.75
                        .Visible = False
                    End With
                Next X
            End If
            With .Surface(A.P1)
                CenterX = .sngX + DX.Surface(S + 21).sngX - DX.Surface(S).Width / 2
                CenterY = .Top + .WhitespaceTop + (.Height - .WhitespaceTop - DX.Surface(S).Height) / 2
            End With
            
            
            If F < 16 Then 'Dim the background
                If F < 10 Then
                    With .Surface(0)
                        .TintBlue = .TintBlue - 0.04
                        .TintGreen = .TintGreen - 0.04
                        .TintRed = .TintRed - 0.04
                    End With
                End If
            Else
                'The next for blocks control moving the ice particles
                If F Mod 3 = 1 And F < 56 Then
                    For X = 0 To 5
                        If .Surface(S + X).sngY = 0 Then
                            With .Surface(S + X)
                                .Visible = True
                                .sngY = -7.9
                                .AngleOfRotation = 0
                            End With
                            Exit For
                        End If
                    Next X
                End If
                If F Mod 3 = 0 And F < 51 Then
                    For X = 5 To 12
                        If .Surface(S + X).sngY = 0 Then
                            With .Surface(S + X)
                                .Visible = True
                                .sngY = -5.6
                            End With
                            With .Surface(S + X + 8)
                                .Visible = True
                                .sngY = 1
                            End With
                            Exit For
                        End If
                    Next X
                End If
                For X = 0 To 4
                    With .Surface(S + X)
                        If .Visible Then
                            '.AngleOfRotation = .AngleOfRotation + PI / 20
                            .sngY = .sngY + 8
                        End If
                        If .sngY > 100 Then .sngY = 100
                        If .sngY = 100 And F Mod 3 = 1 Then
                            .sngY = 0
                            .Visible = False
                        Else
                            sngT = A.T(0) * .sngY / 100
                            .Top = CenterY - Sin(.sngX) * sngT
                            .Left = CenterX + Cos(.sngX) * sngT
                        End If
                    End With
                Next X
                For X = 5 To 12
                    With .Surface(S + X)
                        If .Visible Then
                            .sngY = .sngY + 5.7
                            If .sngY > 100 Then .sngY = 100
                            If .sngY = 100 And F Mod 3 = 0 Then
                                .sngY = 0
                                .Visible = False
                                DX.Surface(S + X + 8).Visible = False
                            Else
                                sngT = A.T(0) * .sngY / 100
                                .Top = CenterY - Sin(.sngX) * sngT - 5
                                .Left = CenterX + Cos(.sngX) * sngT
                                DX.Surface(S + X + 8).Move .Left, .Top + 20
                            End If
                        End If
                    End With
                Next X
                
                'Shake and tint the enemy poke
                If F > 28 And F < 76 Then
                    With .Surface(S + 21)
                        If .AlphaBlend < 0.5 Then .AlphaBlend = .AlphaBlend + 0.02
                    End With
                    With .Surface(A.P2)
                        If .Left = .sngX Then
                            .Left = .Left - 2
                        ElseIf .Left > .sngX Then
                            .Left = .Left - 4
                        Else
                            .Left = .Left + 4
                        End If
                    End With
                    .Surface(S + 21).Left = .Surface(A.P2).Left
                ElseIf F = 76 Then
                    With .Surface(A.P2)
                        .Left = .sngX
                    End With
                    .Surface(S + 21).Left = .Surface(A.P2).Left
                ElseIf F > 143 And F < 174 Then
                    'Un-tint the poke
                    With .Surface(S + 21)
                        If .AlphaBlend > 0 Then .AlphaBlend = .AlphaBlend - 0.02
                    End With
                ElseIf F > 175 And F < 187 Then
                    'Un-dim the background
                    With .Surface(0)
                        .TintBlue = .TintBlue + 0.04
                        .TintGreen = .TintGreen + 0.04
                        .TintRed = .TintRed + 0.04
                    End With
                ElseIf F = 188 Then
                    DX.AnimFinished A.Index
                End If
            End If
        Case 19 'Flickering Ice Particles
            
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "IceParticleSmall.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 6
                    .DuplicateSurface S
                Next X
                                
                .Surface(S).Framed = True
                .Surface(S).scrHeight = 16
                With .Surface(A.P2)
                    CenterX = .Left + (.Width - DX.Surface(S).Width) / 2
                    CenterY = .Top + .WhitespaceTop + (.Height - .WhitespaceTop) / 2 '- DX.Surface(S).Height) / 2
                End With
                
                For X = 0 To 6
                    With .Surface(S + X)
                        .Framed = True
                        Select Case X
                        Case 0: .Move CenterX - 10, CenterY - 14
                        Case 1: .Move CenterX + 10, CenterY + 20
                        Case 2: .Move CenterX - 5, CenterY + 6
                        Case 3: .Move CenterX + 17, CenterY - 12
                        Case 4: .Move CenterX - 10, CenterY - 14
                        Case 5: .Move CenterX + 0, CenterY + 0
                        Case 6: .Move CenterX + 20, CenterY - 2
                        End Select
                        .sngY = .Top
                        .sngX = .Left
                        If X = 0 Or X = 2 Or X = 6 Then
                            .scrHeight = 16
                        Else
                            .scrTop = 16
                        End If
                        .AlphaBlend = 0.75
                        .Visible = False
                    End With
                Next X
            End If
            
            For X = 0 To 6
                Y = F - X * 6
                With .Surface(S + X)
                    If Y < 0 Then
                        .Visible = False
                    ElseIf Y = 0 Then
                        .Width = .Width - 2
                        .Left = .Left + 1
                        If .scrHeight = 16 Then
                            .Height = .Height - 6
                            .Top = .Top + 3
                        End If
                        .Visible = True
                    ElseIf Y < 4 Then
                        If .scrHeight = 16 Then
                            .Height = .Height + 2
                            .Top = .Top - 1
                        End If
                    ElseIf Y = 9 Then
                        .Move .sngX, .sngY, .scrWidth, .scrHeight
                    ElseIf Y > 17 And Y < 37 Then
                        .Visible = (Y Mod 2 = 1)
                    End If
                End With
            Next X
            If F = 72 Then .AnimFinished A.Index
            
        '******* THUNDERBOLT *******
        Case 20 'Falling Thunderbolts
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "BoltSmall.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                .DuplicateSurface S
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "BoltLarge.gif")
                .CreateSurfaceFromPBox MainContainer.SwapSpace, -2
                
                X = .Surface(A.P2).Left + .Surface(A.P2).Width / 2
                .Surface(S).Left = X - 20
                .Surface(S + 1).Left = X + 20
                .Surface(S + 2).Left = X - 2
                Y = .Surface(A.P2).sngY + .Surface(A.P2).Height - .Surface(S).Height
                .Surface(S).Top = Y + 4
                .Surface(S + 1).Top = Y + 4
                .Surface(S + 2).Top = Y
                For X = S To S + 2
                    With .Surface(X)
                        .sngX = .Left
                        .sngY = .Top
                        .Framed = True
                        .scrHeight = 8
                        .Visible = False
                    End With
                Next X
            End If
            
            If F < 20 Then 'Dim the background
                If F < 6 Then
                    With .Surface(0)
                        .TintBlue = .TintBlue - 0.06666
                        .TintGreen = .TintGreen - 0.06666
                        .TintRed = .TintRed - 0.06666
                    End With
                End If
            Else
                Select Case F
                Case Is > 37: X = 2
                Case Is > 29: X = 1
                Case Is > 19: X = 0
                End Select
                For X = S To S + X
                    With .Surface(X)
                        .Visible = True
                        If .scrHeight < 96 And .scrTop = 0 Then
                            .scrHeight = .scrHeight + 8
                        Else
                            If .scrHeight - 8 = 0 Then
                                .Visible = False
                            Else
                                .scrTop = .scrTop + 8
                                .scrHeight = .scrHeight - 8
                                .Top = .Top + 8
                            End If
                        End If
                    End With
                Next X
            End If
            
            If F > 46 And F < 60 Then
                With .Surface(A.P2)
                    .TintBlue = .TintBlue - 0.0714286
                    .TintGreen = .TintGreen - 0.0714286
                    .TintRed = .TintRed - 0.0714286
                End With
            ElseIf F > 59 And F < 75 Then
                With .Surface(A.P2)
                    .TintBlue = .TintBlue + 0.0714286
                    .TintGreen = .TintGreen + 0.0714286
                    .TintRed = .TintRed + 0.0714286
                End With
            End If
             
            If F > 142 Then
                With .Surface(0)
                    .TintBlue = .TintBlue + 0.06666
                    .TintGreen = .TintGreen + 0.06666
                    .TintRed = .TintRed + 0.06666
                End With
            End If
            If F = 148 Then
                With .Surface(0)
                    .TintBlue = 1: .TintGreen = 1: .TintRed = 1
                End With
                With .Surface(A.P2)
                    .TintBlue = 1: .TintGreen = 1: .TintRed = 1
                End With
                .AnimFinished A.Index
            End If
        Case 21 'Big spark thingy
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "DotSpark.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "Spark.gif")
                Call .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                For X = 1 To 6
                    .DuplicateSurface S + 1
                Next X
                For X = S To S + 7
                    .Surface(X).Framed = True
                    .Surface(X).Visible = False
                Next X
                With .Surface(A.P2)
                    DX.Surface(S).Left = .Left + (.Width - DX.Surface(S).Width) / 2
                    DX.Surface(S).Top = .Top + .WhitespaceTop + (.Height - .WhitespaceTop - DX.Surface(S).Height) / 2
                End With
            End If
            
            With .Surface(S)
                Select Case F
                Case 1, 2, 3, 18, 19, 40, 41
                    .Visible = True
                    .scrTop = 0
                Case 8, 9, 11, 12, 23, 24, 27, 42, 43
                    .Visible = True
                    .scrTop = 64
                Case 16, 17, 32 To 35
                    .Visible = True
                    .scrTop = 128
                Case Else
                    .Visible = False
                End Select
            End With
            
            With .Surface(A.P2)
                CenterX = .Left + (.Width - DX.Surface(S + 1).Width) / 2
                CenterY = .Top + .WhitespaceTop + (.Height - .WhitespaceTop - DX.Surface(S + 1).Height) / 2
            End With
            For X = 1 To Int(Rnd * 6) + 1
                With .Surface(S + X)
                    .Visible = True
                    .Left = CenterX + Int(Rnd * 60) - 30
                    .Top = CenterY + Int(Rnd * 60) - 30
                    .scrTop = .Width * Int(Rnd * 2)
                    .AngleOfRotation = Int(Rnd * 2 * PI)
                End With
            Next X
            For X = X To 7
                .Surface(S + X).Visible = False
            Next X
            
            If F = 3 Or F = 19 Then
                With .Surface(0)
                    .TintBlue = 1
                    .TintGreen = 1
                    .TintRed = 1
                End With
            ElseIf F = 11 Or F = 27 Then
                With .Surface(0)
                    .TintBlue = 0.6
                    .TintGreen = 0.6
                    .TintRed = 0.6
                End With
            End If
            
            If F = 45 Then
                .AnimFinished A.Index
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
            End If
            
        '******* FLAMETHROWER *******
        Case 22 'Flamethrower
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "Flame.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 7
                    .DuplicateSurface S
                Next X
                
                .Surface(S).Framed = True
                With .Surface(A.P1)
                    CenterX = .sngX - DX.Surface(S).Width / 2
                    CenterY = .Top + .WhitespaceTop + (.Height - .WhitespaceTop - DX.Surface(S).Height) / 2
                End With
                With .Surface(A.P2)
                    X = .sngX + (.Width - DX.Surface(S).Width) \ 2
                    Y = .Top + .WhitespaceTop + (.Height - .WhitespaceTop) / 2 - DX.Surface(S).Height / 2
                End With
                
                ReDim A.T(2)
                If X > CenterX Then
                    A.T(1) = .Surface(A.P1).Width - .Surface(A.P1).WhitespaceRight
                    CenterX = CenterX + A.T(1)
                Else
                    A.T(1) = .Surface(A.P1).WhitespaceLeft
                    CenterX = CenterX + A.T(1)
                End If
                
                
                A.T(0) = Sqr((CenterY - Y) ^ 2 + (X - CenterX) ^ 2)
                .Surface(S).sngX = Atn((CenterY - Y) / (X - CenterX))
                If CenterX > X Then .Surface(S).sngX = .Surface(S).sngX + PI
                
                For X = 0 To 7
                    With .Surface(S + X)
                        .Framed = True
                        .Left = CenterX
                        .Top = CenterY
                        .sngX = DX.Surface(S).sngX
                        .sngY = 0
                        .Visible = False
                        .TintBlue = 0.95
                        .TintRed = 0.95
                        .TintGreen = 0.95
                    End With
                Next X
            End If
            With .Surface(A.P1)
                CenterX = .sngX + A.T(1) + 5 - DX.Surface(S).Width / 2
                CenterY = .Top + .WhitespaceTop + 5 + (.Height - .WhitespaceTop - DX.Surface(S).Height) / 2
            End With
            
            
            'Create new flame
            If F Mod 4 = 0 And F > 9 And F < 95 Then
                For X = 0 To 7
                    If .Surface(S + X).sngY = 0 Then
                        With .Surface(S + X)
                            .Visible = True
                            .sngY = -3.9
                            .scrTop = .scrWidth
                        End With
                        Exit For
                    End If
                Next X
            End If
            
            'Move the flames
            A.T(1) = A.T(1) + 0.05
            For X = 0 To 7
                With .Surface(S + X)
                    If .Visible Then .sngY = .sngY + 4
                    If .sngY > 100 Then
                        .sngY = 0
                        .Visible = False
                    Else
                        sngT = A.T(0) * .sngY / 100
                        
                        .Top = CenterY - Sin(.sngX) * sngT - Sin(A.T(1) + .sngY * PI / 100) * 12 'Adding a Sin gives it a nice wavy effect
                        .Left = CenterX + Cos(.sngX) * sngT
                    End If
                End With
            Next X
            
            'Advance the flame frames
            If F Mod 2 = 1 Then
                For X = 0 To 7
                    With .Surface(S + X)
                        Y = .scrTop
                        Y = Y + 32
                        If Y = 128 Then Y = 32
                        .scrTop = Y
                    End With
                Next X
            End If
            
            'Shake the enemy
            If F > 34 And F < 116 And F Mod 2 = 0 Then
                With .Surface(A.P2)
                    If .Left = .sngX Then
                        .Left = .Left - 2
                    ElseIf .Left > .sngX Then
                        .Left = .Left - 4
                    Else
                        .Left = .Left + 4
                    End If
                End With
            End If
            If F = 116 Then .Surface(A.P2).Left = .Surface(A.P2).sngX
            
            'Bounce the attacker
            If F > 0 And F < 91 Then
                If F Mod 4 = 1 Then
                    .Surface(A.P1).Top = .Surface(A.P1).Top + 1
                ElseIf F Mod 4 = 3 Then
                    .Surface(A.P1).Top = .Surface(A.P1).Top - 1
                End If
            End If
            
            If F = 122 Then
                For X = 0 To 7
                    .Surface(S + X).Visible = False
                Next X
                DX.AnimFinished A.Index
            End If
            
        '******* SURF ******
        Case 23 'Surf
            If F = 0 Then
                X = RGB(96, 88, 248)
                S = .CreateSolidColorSurface(X, 140, 256)
                
                A.FirstSurf = S
                If A.P1 = 1 Then 'Player to Opponent
                    Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "SurfAtk.gif")
                    .CreateSurfaceFromPBox MainContainer.SwapSpace, RGB(104, 144, 136)
                    Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "WaterAtk.gif")
                Else
                    Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "SurfDef.gif")
                    .CreateSurfaceFromPBox MainContainer.SwapSpace, RGB(104, 144, 136)
                    Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "WaterDef.gif")
                End If
                
                .CreateSurfaceFromPBox MainContainer.SwapSpace, RGB(104, 144, 136)
                .DuplicateSurface S + 2
                
                If A.P1 = 1 Then
                    .Surface(S).Move 0, 116
                    .Surface(S + 1).Move 0, 52
                    .Surface(S + 2).Move -256, 100
                    .Surface(S + 3).Move 128, 100
                Else
                    .Surface(S).Move 0, -204
                    .Surface(S + 1).Move 224, -64
                    .Surface(S + 2).Move 352, -64
                    .Surface(S + 3).Move -32, -64
                End If
                
                For X = 0 To 3
                    With .Surface(S + X)
                        .AlphaBlend = 0
                    End With
                Next X
            End If
            
            
            For X = S To S + 3
                With .Surface(X)
                    If A.P1 = 1 Then
                        If X <> S Then .Left = .Left + 2
                        If .Left > 250 And X > S + 3 Then .Left = .Left - 352
                        .Top = .Top - 1
                    Else
                        If X <> S Then .Left = .Left - 2
                        If .Left < -40 And X > S + 3 Then .Left = .Left + 352
                        .Top = .Top + 1
                    End If
                    If F < 28 Then
                        .AlphaBlend = .AlphaBlend + 0.03
                    ElseIf F > 106 Then
                        .AlphaBlend = .AlphaBlend - 0.03
                    End If
                End With
            Next X
            
            
            If F = 132 Then
                For X = S To S + 3
                    .Surface(S).Visible = False
                Next X
                DX.AnimFinished A.Index
            End If
        '******* EARTHQUAKE *******
        Case 24 'Shake the ground
            If F = 0 Then
                S = .DuplicateSurface(0)
                A.FirstSurf = S
                With .Surface(S)
                    .Framed = True
                    .scrHeight = DX.Surface(0).Height
                    .scrWidth = 5
                    .Width = DX.Surface(0).Width
                End With
                    
                .CreateSolidColorSurface RGB(255, 255, 255), .Surface(0).Height, .Surface(0).Width
                .CreateSolidColorSurface RGB(0, 0, 0), .Surface(0).Height, .Surface(0).Width
                .SetZOrder S, 2
                .SetZOrder S + 1, 3
                .SetZOrder S + 2, 4
                .Surface(S + 1).Visible = False
                .Surface(S + 2).Visible = False
            End If
            If F < 120 Then
                X = 15
            Else
                X = Round(15 * (200 - F) / 80)
            End If
            
            If F Mod 2 = 0 Then
                With .Surface(0)
                    .Left = 0
                    If F Mod 4 = 0 Then
                        .Left = .Left + X
                    Else
                        .Left = .Left - X
                    End If
                    If X = 0 Then
                        DX.Surface(S).Visible = False
                    Else
                        DX.Surface(S).Left = .Left - .Width * Sgn(.Left)
                    End If
                End With
            Else
                For Y = 1 To A.P1
                    With .Surface(Y)
                        .Left = .sngX
                        X = X \ 2
                        If F Mod 4 = 1 Then
                            .Left = .Left + X
                        Else
                            .Left = .Left - X
                        End If
                    End With
                Next Y
            End If
            
            Select Case F
            Case 13, 31
                .Surface(S + 1).Visible = True
            Case 16, 34
                .Surface(S + 1).Visible = False
                .Surface(S + 2).Visible = True
            Case 21, 39
                .Surface(S + 2).Visible = False
            Case 200
                For X = 1 To A.P1
                    With .Surface(X)
                        .Left = .sngX
                    End With
                Next X
                .Surface(0).Left = 0
                .Surface(S).Visible = False
                .AnimFinished A.Index
            End Select
        
        
        '******* METRONOME *******
        Case 5025:
            If F = 0 Then
                MainContainer.SwapSpace.Picture = LoadPicture(iPath & "ThoughtBubble.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                With .Surface(S)
                    If A.P2 = -1 Then
                        .Left = DX.Surface(A.P1).Left + 48
                        .HFlip
                    Else
                        .Left = DX.Surface(A.P1).Left - 12
                    End If
                    .Framed = True
                    .Top = DX.Surface(A.P1).Top
                    .Visible = True
                End With
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "Metronome.gif")
                Call .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                .Surface(S + 1).Visible = False
            End If
            With .Surface(S)
                Select Case F
                    Case 2, 113:
                        .scrTop = .scrWidth * 1
                    Case 4, 111:
                        .scrTop = .scrWidth * 2
                    Case 6:
                        .scrTop = .scrWidth * 3
                    Case 115:
                        .scrTop = .scrWidth * 0
                    End Select
            End With
            With .Surface(S + 1)
                Select Case F
                    Case 9 To 16:
                        If F = 9 Then .Visible = True
                        .Height = .scrHeight / (17 - F)
                        .Width = .scrWidth / (17 - F)
                        .Left = DX.Surface(S).Left + (DX.Surface(S).Width) / 2 - .Width / 2
                        .Top = DX.Surface(S).Top + (DX.Surface(S).Height) / 2 - .Height / 2
                    Case 37 To 47, 59 To 69, 81 To 91:
                        .AngleOfRotation = .AngleOfRotation + 0.1
                    Case 48 To 58, 70 To 80, 92 To 102:
                        .AngleOfRotation = .AngleOfRotation - 0.1
                    Case 101 To 107:
                        .Height = .scrHeight / (F - 100)
                        .Width = .scrWidth / (F - 100)
                        .Left = DX.Surface(S).Left + (DX.Surface(S).Width) / 2 - .Width / 2
                        .Top = DX.Surface(S).Top + (DX.Surface(S).Height) / 2 - .Height / 2
                    Case 108
                        .Visible = False
                End Select
            End With
            If F = 118 Then .AnimFinished A.Index
             
        '******* SUNNY DAY *******
        Case 5026:
            If F = 0 Then
                Set MainContainer.SwapSpace.Picture = LoadPicture(iPath & "SunnyDay.gif")
                A.FirstSurf = .CreateSurfaceFromPBox(MainContainer.SwapSpace, -2)
                S = A.FirstSurf
                For X = 1 To 3
                    .DuplicateSurface S
                    .Surface(S + X).Trans = -2
                Next X
                Call .CreateSolidColorSurface(vbWhite, .Surface(0).Height, .Surface(0).Width)
                For X = S To S + 3
                    With .Surface(X)
                        .Height = .scrHeight / 2
                        .Width = .scrWidth / 2
                        .AngleOfRotation = .AngleOfRotation - PI / 4
                        .Left = -19 * X
                        .Top = -11 * X
                        .sngX = .Left
                        .sngY = .Top
                        .sngZ = 0
                    End With
                Next X
                .Surface(S + 4).AlphaBlend = 0
            End If
            With .Surface(S + 4)
                Select Case F
                    Case 3, 6, 9, 12
                        .AlphaBlend = .AlphaBlend + 0.05
                    Case 102, 105, 108, 111
                        .AlphaBlend = .AlphaBlend - 0.05
                End Select
            End With
            If .Surface(S + 3).Left < .Surface(0).Width Then
                    For X = 0 To 3
                        With .Surface(S + X)
                            .sngX = .sngX + 2
                            .sngY = .sngY + 1
                            .sngZ = .sngZ + 0.005
                            .Left = Round(.sngX)
                            .Top = Round(.sngY)
                            .AngleOfRotation = .AngleOfRotation - PI / 16
                            If .Left >= 60 And .Top >= 60 Then
                                If .Height < .scrHeight * 0.75 Then .Height = .Height + .sngZ
                                If .Width < .scrWidth * 0.75 Then .Width = .Width + .sngZ
                            End If
                        End With
                    Next X
            Else: .AnimFinished A.Index
            End If
        '******* GENERAL *******
        End Select
    End With
End Sub

Private Function GetPokeballColor(Ball As Long) As Long
    Select Case Ball
    Case 0: GetPokeballColor = RGB(248, 176, 240)  'Poke
    Case 1: GetPokeballColor = RGB(128, 184, 240)  'Great
    Case 2: GetPokeballColor = RGB(248, 248, 120)  'Ultra
    Case 3: GetPokeballColor = RGB(184, 160, 224)  'Master
    Case 4: GetPokeballColor = RGB(80, 200, 240)   'Dive
    Case 5: GetPokeballColor = RGB(240, 216, 80)   'Nest
    Case 6: GetPokeballColor = RGB(232, 240, 240)  'Timer
    Case 7: GetPokeballColor = RGB(168, 248, 200)  'Net
    Case 8: GetPokeballColor = RGB(184, 240, 160)  'Safari
    Case 9: GetPokeballColor = RGB(248, 136, 80)   'Luxury
    Case 10: GetPokeballColor = RGB(248, 72, 80)   'Premier
    Case 11: GetPokeballColor = RGB(248, 192, 128) 'Repeat
    End Select
End Function



