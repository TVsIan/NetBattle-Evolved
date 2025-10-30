Attribute VB_Name = "CopyMod"
Option Explicit

Public Function LegalMove(Pokemon As Pokemon, Optional IgnoreDVs As Boolean = False) As String
    Dim X As Integer
    Dim Y As Integer
    Dim BreedingMoves As Integer
    Dim GSBreedingMoves As Integer
    Dim RBYMoves As Integer
    Dim RBYConflict As Integer
    Dim GSCConflict As Integer
    Dim InvalidMove As Boolean
    Dim SpecialMoves As Integer
    Dim GSSpecialMoves As Integer
    Dim SurfingPika As Integer
    Dim OddEggPoke As Boolean
    Dim Temp As String
    Dim B As Boolean
    Dim Mv(4) As Integer
    Dim AllMoves() As Boolean
    
    'First, catalog all the moves into an array.
    ReDim AllMoves(UBound(Moves), 1 To 13)
    With Pokemon
        Select Case .GameVersion
            Case nbTrueRBY
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbTrueGSC
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbRBYTrade, nbGSCTrade
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbRBYTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbGSCSpecial)
            Case nbTrueRuSa
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvSpecial)
            Case nbFullAdvance
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvLevel)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTM)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvEgg)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvSpecial)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvTutor)
                Call ValidMoveArray(Pokemon, AllMoves(), nbAdvFL)
            Case Else
                If InVBMode Then Stop
                LegalMove = "Invalid Input"
        End Select
        
        For X = 1 To UBound(AllMoves, 2)
            AllMoves(0, X) = False
        Next X
        For X = 1 To 4
            Mv(X) = CInt(Pokemon.Move(X))
        Next X
        InvalidMove = False
        BreedingMoves = 0
        GSBreedingMoves = 0
        RBYMoves = 0
        LegalMove = ""
        
        'Invalid moves.
        For X = 1 To 4
            If Mv(X) <> 0 Then
                 InvalidMove = True
                 For Y = 1 To UBound(AllMoves, 2)
                     If AllMoves(Mv(X), Y) Then InvalidMove = False
                 Next Y
                 If InvalidMove Then
                     LegalMove = .Name & " can't learn " & Moves(Mv(X)).Name & ".  There may have been a recent change to the database."
                     Exit Function
                 End If
             End If
        Next X
        
        'Invalid breeding combinations.
        If Not BreedCheck(Pokemon.No, Mv) Then
            LegalMove = .Name & "'s moveset contains a combination of Breeding Moves that is not legally obtainable."
            Exit Function
        End If
        
        'RBY Moves <-/-> GSCBreed, GSCSpecial
        RBYMoves = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbRBYLevel, nbRBYTM) Then
                RBYMoves = Mv(X)
                Exit For
            End If
        Next X
        
        RBYConflict = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg, nbGSCSpecial) Then
                If Not Moves(Mv(X)).RBYMove Then
                    RBYConflict = Mv(X)
                    B = False
                    Exit For
                End If
            End If
        Next X
        
        If RBYConflict <> 0 And RBYMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(RBYMoves).Name & " and " & Moves(RBYConflict).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine RBY Moves and non-RBY Breeding or Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        
        'Only 1 Special Move
        SpecialMoves = 0
        Y = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbAdvSpecial, nbGSCSpecial) Then
                If Y = 0 Then
                    Y = Mv(X)
                Else
                    SpecialMoves = Mv(X)
                    Exit For
                End If
            End If
        Next X
        
        If SpecialMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(Y).Name & " and " & Moves(SpecialMoves).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine two or more Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        '5th Check: Breeding Moves <-/-> Special Moves
        BreedingMoves = 0
        
        'EXCEPTION: Gligar can have Earthquake and {Wing Attack and/or Counter}
        B = False
        If Pokemon.No = 207 Then
            For X = 1 To 4
                If Mv(X) = 55 Then B = True: Exit For
            Next X
        End If
        
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg, nbAdvEgg) Then
                If (Mv(X) <> 248 And Mv(X) <> 34) Or Not B Then
                    BreedingMoves = Mv(X)
                    Exit For
                End If
            End If
        Next X
        
        SpecialMoves = 0
        For X = 1 To 4
            If LegalMoveCheck(AllMoves, Mv(X), nbGSCSpecial, nbAdvSpecial) Then
                SpecialMoves = Mv(X)
                Exit For
            End If
        Next X
        
        If BreedingMoves <> 0 And SpecialMoves <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(BreedingMoves).Name & " and " & Moves(SpecialMoves).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine Breeding Moves with Special Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        If .GameVersion < nbTrueRuSa And .GameVersion <> nbFullAdvance Then
        
            '6th Check: GSC Breeding Moves <-/-> Special Moves or RBY Moves
            'It sounds redundant but it's not.  Just trust me. -_-
            GSBreedingMoves = 0
            For X = 1 To 4
                If LegalMoveCheck(AllMoves, Mv(X), nbGSCEgg) Then
                    If Not Moves(Mv(X)).RBYMove Then
                        GSBreedingMoves = Mv(X)
                        Exit For
                    End If
                End If
            Next X
            
            GSCConflict = 0
            For X = 1 To 4
                If LegalMoveCheck(AllMoves, Mv(X), nbRBYLevel, nbRBYTM) Then
                    GSCConflict = Mv(X)
                    Exit For
                End If
            Next X
                    
            If GSBreedingMoves <> 0 And GSCConflict <> 0 Then
                Temp = .Name & " cannot learn both " & Moves(GSBreedingMoves).Name & " and " & Moves(GSCConflict).Name & "." & vbNewLine
                Temp = Temp & "(Cannot combine non-RBY Breeding Moves with RBY Moves or Special Moves.)"
                LegalMove = Temp
                Exit Function
            End If
                    
            'Odd Eggs and Dizzy Punch
            Select Case Pokemon.No
            Case 173, 35, 36, 174, 39, 40, 236, 106, 107, 175, 176, 237, 238, 124, 240, 126
                For X = 1 To 4
                    If Mv(X) = 45 Then OddEggPoke = True
                Next
            End Select
            If OddEggPoke And BreedingMoves <> 0 Then
                Temp = .Name & " cannot learn both Dizzy Punch and " & Moves(BreedingMoves).Name & "." & vbNewLine
                Temp = Temp & "(Cannot combine Dizzy Punch and Breeding Moves on Odd Egg Pokémon.)"
                LegalMove = Temp
                Exit Function
            End If
            With Pokemon
                If Not IgnoreDVs Then
                    If OddEggPoke And Not ((.DV_Atk = 2 And .DV_Def = 10 And .DV_SAtk = 10 And .DV_Spd = 10) _
                    Or (.DV_Atk = 0 And .DV_Def = 0 And .DV_SAtk = 0 And .DV_Spd = 0)) Then
                        LegalMove = "DVs must be either 2/10/10/10 or 0/0/0/0 in order for Dizzy Punch to be on " & .Name
                        Exit Function
                    End If
                End If
            End With
        End If
        'GSC Pre-evolution moves <-/-> RBY Moves
        RBYConflict = 0
        Select Case Pokemon.No
        Case 35, 36, 39, 40
            For X = 1 To 4
                If Mv(X) = 25 Or Mv(X) = 100 Or Mv(X) = 219 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        Case 25, 26, 124
            For X = 1 To 4
                If Mv(X) = 25 Or Mv(X) = 219 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        Case 130
            For X = 1 To 4
                If Mv(X) = 68 Then
                    RBYConflict = Mv(X)
                    Exit For
                End If
            Next X
        End Select
        If RBYMoves <> 0 And RBYConflict <> 0 Then
            Temp = .Name & " cannot learn both " & Moves(RBYMoves).Name & " and " & Moves(RBYConflict).Name & "." & vbNewLine
            Temp = Temp & "(Cannot combine GSC Pre-Evolution Moves with RBY Moves.)"
            LegalMove = Temp
            Exit Function
        End If
        
        'Vaporeon
        If Pokemon.No = 134 Then
            If HasMoves(Mv(), 12, 129) Then
                Temp = .Name & " cannot learn both " & Moves(12).Name & " and " & Moves(129).Name & "." & vbNewLine
                LegalMove = Temp
                Exit Function
            End If
            If HasMoves(Mv(), 12, 24) Then
                Temp = .Name & " cannot learn both " & Moves(12).Name & " and " & Moves(24).Name & "." & vbNewLine
                LegalMove = Temp
                Exit Function
            End If
        End If
    End With
End Function

Private Function LegalMoveCheck(AllMoves() As Boolean, MoveNum As Integer, C1 As MoveTypes, Optional C2 As MoveTypes, Optional C3 As MoveTypes, Optional C4 As MoveTypes, Optional C5 As MoveTypes, Optional C6 As MoveTypes, Optional C7 As MoveTypes, Optional C8 As MoveTypes, Optional C9 As MoveTypes) As Boolean
    Dim Build As Boolean
    Dim X As Integer
    Build = False
    On Error Resume Next
    If AllMoves(MoveNum, C1) = True Then Build = True
    If AllMoves(MoveNum, C2) = True Then Build = True
    If AllMoves(MoveNum, C3) = True Then Build = True
    If AllMoves(MoveNum, C4) = True Then Build = True
    If AllMoves(MoveNum, C5) = True Then Build = True
    If AllMoves(MoveNum, C6) = True Then Build = True
    If AllMoves(MoveNum, C7) = True Then Build = True
    If AllMoves(MoveNum, C8) = True Then Build = True
    If AllMoves(MoveNum, C9) = True Then Build = True
    For X = LBound(AllMoves, 2) To UBound(AllMoves, 2)
        If X <> C1 And X <> C2 And X <> C3 And X <> C4 And X <> C5 And X <> C6 And X <> C7 And X <> C8 And X <> C9 Then
            If AllMoves(MoveNum, X) = True Then Build = False
        End If
    Next X
    LegalMoveCheck = Build
End Function
Sub ValidMoveArray(ByRef PKMN As Pokemon, ByRef FillArray() As Boolean, ByVal MoveType As MoveTypes)
    Dim X As Integer
    
    With PKMN
        Select Case MoveType
            'RBY Moves
            Case nbRBYLevel
                For X = 1 To UBound(.RBYMoves)
                    FillArray(.RBYMoves(X), MoveType) = True
                Next
            'RBY TMs
            Case nbRBYTM
                For X = 1 To UBound(.RBYTM)
                    FillArray(.RBYTM(X), MoveType) = True
                Next
            'GSC Moves
            Case nbGSCLevel
                For X = 1 To UBound(.BaseMoves)
                    FillArray(.BaseMoves(X), MoveType) = True
                Next
            'GSC TMs
            Case nbGSCTM
                For X = 1 To UBound(.MachineMoves)
                    FillArray(.MachineMoves(X), MoveType) = True
                Next
            'GSC Egg Moves
            Case nbGSCEgg
                For X = 1 To UBound(.BreedingMoves)
                    FillArray(.BreedingMoves(X), MoveType) = True
                Next
            'GSC Tutor Moves
            Case nbGSCTutor
                For X = 1 To UBound(.MoveTutor)
                    FillArray(.MoveTutor(X), MoveType) = True
                Next
            'GSC Special Moves
            'We'll use this array for True RBY and True GSC's Stadium/Crystal/Odd Egg moves - don't fill in the normal Special moves.
            Case nbGSCSpecial
                If .GameVersion <> nbTrueGSC And .GameVersion <> nbTrueRBY Then
                    For X = 1 To UBound(.SpecialMoves)
                        FillArray(.SpecialMoves(X), MoveType) = True
                    Next
                End If
            'Advance moves
            Case nbAdvLevel
                For X = 1 To UBound(.AdvMoves)
                    FillArray(.AdvMoves(X), MoveType) = True
                Next
            'Advance TMs
            Case nbAdvTM
                For X = 1 To UBound(.ADVTM)
                    FillArray(.ADVTM(X), MoveType) = True
                Next
            'Advance Egg Moves
            Case nbAdvEgg
                For X = 1 To UBound(.AdvBreeding)
                    FillArray(.AdvBreeding(X), MoveType) = True
                Next
            'Advance Tutor Moves
            Case nbAdvTutor
                For X = 1 To UBound(.AdvTutor)
                    FillArray(.AdvTutor(X), MoveType) = True
                Next
            'Advance Special Moves
            Case nbAdvSpecial
                For X = 1 To UBound(.AdvSpecial)
                    FillArray(.AdvSpecial(X), MoveType) = True
                Next
            'Fire/Leaf Moves
            Case nbAdvFL
                For X = 1 To UBound(.LFOnly)
                    FillArray(.LFOnly(X), MoveType) = True
                Next
        End Select
        Select Case .GameVersion
            Case nbRBYTrade
                'RBY Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).RBYMove Then FillArray(X, MoveType) = False
                Next
            Case nbGSCTrade
                'RBY/GSC Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).GSCMove Then FillArray(X, MoveType) = False
                Next
            Case nbTrueRuSa, nbFullAdvance, nbAdvTrades
                'Nothing special
            Case nbTrueRBY
                'RBY Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).RBYMove Then FillArray(X, MoveType) = False
                Next
                'I know this part could be optimized, but I want to keep it easy for any future changes we may need.
                If MoveType = nbGSCSpecial Then
                    Select Case .No
                        Case 54, 55
                            FillArray(6, MoveType) = True
                    End Select
                End If
            Case nbTrueGSC
                'RBY/GSC Moves Only
                For X = 1 To UBound(Moves)
                    If Not Moves(X).GSCMove Then FillArray(X, MoveType) = False
                Next
                If MoveType = nbGSCSpecial Then
                    Select Case .No
                        Case 83
                            FillArray(12, MoveType) = True
                        Case 147 To 149
                            FillArray(61, MoveType) = True
                        Case 207
                            FillArray(55, MoveType) = True
                        Case 25, 26, 35, 36, 39, 40, 106, 107, 125, 126, 135, 172 To 174, 236 To 239
                            FillArray(45, MoveType) = True
                    End Select
                End If
        End Select
    End With
End Sub
Public Function HasMoves(MoveArray() As Integer, ByVal M1 As Integer, Optional ByVal M2 As Integer = 0, Optional ByVal M3 As Integer = 0, Optional ByVal M4 As Integer = 0) As Boolean
    Dim Check() As Integer
    Dim A As Integer
    Dim X As Integer
    Dim Y As Integer
    If M4 = 0 Then
        If M3 = 0 Then
            If M2 = 0 Then
                ReDim Check(1 To 1)
            Else
                ReDim Check(1 To 2)
            End If
        Else
            ReDim Check(1 To 3)
        End If
    Else
        ReDim Check(1 To 4)
    End If
    On Error Resume Next
    Check(1) = M1
    Check(2) = M2
    Check(3) = M3
    Check(4) = M4
    A = 0
    For X = 1 To UBound(Check)
        For Y = LBound(MoveArray) To UBound(MoveArray)
            If MoveArray(Y) <> 0 And MoveArray(Y) = Check(X) Then
                A = A + 1
                Exit For
            End If
        Next Y
    Next X
    HasMoves = (A = UBound(Check))
End Function



