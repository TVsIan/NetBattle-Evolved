Attribute VB_Name = "ScriptMod2"
Option Explicit
'Scripting Module v2.0
'by MasamuneXGP
Private Type ValueType
    Real As Boolean
    Link As Long
    Op As Operators
    UseNeg As Boolean
    UseNot As Boolean
    UseAlt As Boolean
End Type
Private Type TermType
    V() As ValueType
End Type
Private Type BlockType
    FunctionNum As Long
    ArgList() As TermType
End Type
Private Type LineType
    LineNum As Long
    CommandNum As Long
    MainBlock As Long
End Type
Private Type EventType
    L() As LineType
End Type
Private Type ArgType
    Name As String
    Type As VbVarType
End Type
Private Type VarStruct
    FuncLink As Long
    Value() As Variant
    LB As Long
    UB As Long
End Type
Private Type FuncType
    Name As String
    Args() As ArgType
    Type As VbVarType
    Index As Long
    IsVar As Boolean
    Opt As Byte
End Type
Private Type CommandType
    Name As String
    Args() As ArgType
    Index As Long
End Type
Private Type LevelType
    Type As Integer
    Brace As Boolean
    StartNum As Long
    StartLine As String
End Type
Private Type ForStruct
    LNum As Long
    Var As Long
    LB As Variant
    UB As Variant
    Stp As Variant
    Jump As Long
End Type
Type LocalType
    Vars() As VarStruct
    Funcs() As FuncType
End Type
Private Enum Operators
    nbsNone
    nbsExpon
    nbsNegat
    nbsMulti
    nbsFloatDiv
    nbsMod
    nbsAdd
    nbsConcat
    nbsEqual
    nbsInequal
    nbsLessThan
    nbsGreaterThan
    nbsLessEqual
    nbsGreaterEqual
    nbsNot
    nbsAnd
    nbsOr
    nbsXor
    nbsEqv
End Enum
Private StackLevel As Long
Private Stack() As LocalType
Private VarNum As Long
Private FuncNum As Long
Private Values() As Variant
Private EventList() As EventType
Private EventName() As String
Private VarList() As VarStruct
Private FuncList() As EventType
Private FuncDefs() As FuncType
Private SubList() As EventType
Private SubDefs() As CommandType
Private BlockList() As BlockType
Const REALSUBSTART As Long = 6
Const FUNCLIMIT As Long = 35
Const SUBLIMIT As Long = 35
Const EVENTNUM As Long = 15
Public Sub InitScript()
    'ReDim Values(0)
    ReDim FuncDefs(0 To FUNCLIMIT)
    ReDim FuncList(0)
    ReDim SubDefs(0 To SUBLIMIT)
    ReDim SubList(0)
    ReDim EventList(0 To EVENTNUM * 2)
    ReDim EventName(0 To EVENTNUM)

    '==============COMMANDS==============
    '****INTERNAL NON-SCRIPT COMMANDS****
    With SubDefs(1)
        '1 Set %VarIndex, %VarSubscript, !NewValue
        '  - Sets a variable to the assigned value.  NewValue's
        '    type is determined by the Variable's type.
        ReDim .Args(1 To 3)
        .Args(1).Type = vbLong
        .Args(2).Type = vbLong
    End With
    With SubDefs(2)
        '2 GoTo %LineNum
        '  - Moves the execution point to the line specified by
        '    LineNum.
        ReDim .Args(1 To 1)
        .Args(1).Type = vbLong
    End With
    With SubDefs(3)
        '3 If @Condition, %LineNum
        '  - If Condition is False, moves the execution point to the
        '    line specified by LineNum.
        ReDim .Args(1 To 2)
        .Args(1).Type = vbBoolean
        .Args(2).Type = vbLong
    End With
    With SubDefs(4)
        '4 For %IncVarIndex, #Start, #End, #Step, %LineNum
        '  - Initiates a For-Next loop.  If it is already initiated,
        '    it increases the value of the variable of IncVarIndex by
        '    #Step.  If the value is greater than #End, it moves the
        '    execution point to the line specified by LineNum.
        ReDim .Args(1 To 5)
        .Args(1).Type = vbLong
        .Args(2).Type = vbDouble
        .Args(3).Type = vbDouble
        .Args(4).Type = vbDouble
        .Args(5).Type = vbLong
    End With
    With SubDefs(5)
        '5 Local %VarType, $VarName, %LBound, %UBound
        '  - Sets a Local Variable with type %VarType,
        '    name $VarName, and a subscript of %LBound
        '    to %UBound
        ReDim .Args(1 To 4)
        .Args(1).Type = vbLong
        .Args(2).Type = vbString
        .Args(3).Type = vbLong
        .Args(4).Type = vbLong
    End With
    '****SCRIPTABLE COMMANDS****
    With SubDefs(6)
        '6 Print
        '  - Adds the argument to the server Text Box
        .Name = "?"
        ReDim .Args(1 To 1)
        '.Args(1).Type = vbString
    End With
    
    Call SetFuncs
End Sub
Private Sub SetFuncs()
    ReDim FuncDefs(0).Args(1 To 1) 'No Function
    With FuncDefs(1)
        '$CStr(!Var)
        .Name = "CStr"
        ReDim .Args(1 To 1)
    End With
    With FuncDefs(2)
        '%CInt(!Var)
        .Name = "CInt"
        ReDim .Args(1 To 1)
    End With
    With FuncDefs(3)
        '#CFloat(!Var)
        .Name = "CFloat"
        ReDim .Args(1 To 1)
    End With
    With FuncDefs(4)
        '@CBool(!Var)
        .Name = "CBool"
        ReDim .Args(1 To 1)
    End With
    With FuncDefs(5)
        '#Round(#Num, [%Places = 0])
        .Name = "Round"
        ReDim .Args(1 To 2)
        .Args(1).Type = vbDouble
        .Args(2).Type = vbLong
        .Opt = 1
    End With
    With FuncDefs(6)
        '#SysTimer
        .Name = "SysTimer"
    End With
    With FuncDefs(7)
        '%Random(%Bound1, [%Bound2 = 0])
        .Name = "Random"
        ReDim .Args(1 To 2)
        .Args(1).Type = vbLong
        .Args(2).Type = vbLong
        .Opt = 1
    End With
    With FuncDefs(8)
        '#Rnd
        .Name = "Rnd"
    End With
    With FuncDefs(9)
        '%Asc($Chr)
        .Name = "Asc"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbString
    End With
    With FuncDefs(10)
        '$Chr(%Asc)
        .Name = "Chr"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbLong
    End With
    With FuncDefs(11)
        '%Len($Text)
        .Name = "Len"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbString
    End With
    With FuncDefs(12)
        '%Search($Text, $Find, [%Start = 1], [@CaseSense = False])
        .Name = "Search"
        ReDim .Args(1 To 4)
        .Args(1).Type = vbString
        .Args(2).Type = vbString
        .Args(3).Type = vbLong
        .Args(4).Type = vbBoolean
        .Opt = 2
    End With
    With FuncDefs(13)
        '$Replace($Text, $Find, $ReplaceWith)
        .Name = "Replace"
        ReDim .Args(1 To 3)
        .Args(1).Type = vbString
        .Args(2).Type = vbString
        .Args(3).Type = vbString
    End With
    With FuncDefs(14)
        '$Left($Text, %Number)
        .Name = "Left"
        ReDim .Args(1 To 2)
        .Args(1).Type = vbString
        .Args(2).Type = vbLong
    End With
    With FuncDefs(15)
        '$Right($Text, %Number)
        .Name = "Right"
        ReDim .Args(1 To 2)
        .Args(1).Type = vbString
        .Args(2).Type = vbLong
    End With
    With FuncDefs(16)
        '$Mid($Text, %Start, %Length)
        .Name = "Mid"
        ReDim .Args(1 To 3)
        .Args(1).Type = vbString
        .Args(2).Type = vbLong
        .Args(3).Type = vbLong
    End With
    With FuncDefs(17)
        '$LCase($Text)
        .Name = "LCase"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbString
    End With
    With FuncDefs(18)
        '$UCase($Text)
        .Name = "UCase"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbString
    End With
    With FuncDefs(19)
        '$Trim($Text)
        .Name = "Trim"
        ReDim .Args(1 To 1)
        .Args(1).Type = vbString
    End With
End Sub
Public Function CompileScript(ByVal Script As String) As String
    Dim A As Long
    Dim B As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim L As Long
    Dim T As Single
    Dim Temp As String
    Dim Piece As String
    Dim Build As String
    Dim E As String
    Dim iLine() As String
    Dim Strings() As String
    Dim Mode As Integer
    Dim TempArray() As String
    Dim TempEvent As EventType
    Dim LNum() As Long
    Dim Level() As LevelType
    Dim DeleteFlag As Boolean
    Dim EndFlag As Integer
    T = Timer
    On Error GoTo ErrorTrap
    E = ""
    ReDim BlockList(0)
    ReDim EventList(0)
    ReDim VarList(0)
    ReDim Strings(0)
    ReDim Level(0)
    ReDim Values(0)
    Erase Stack
    
    'If (Not FuncDefs) = -1 Then
    Call InitScript
    
    'Parse into lines and remove the strings
    iLine = Split(Script, vbNewLine)
    For L = 0 To UBound(iLine)
        iLine(L) = Trim$(iLine(L))
        If InStr(1, iLine(L), vbNullChar) > 0 Then GoTo Nullchar
        If InStr(1, iLine(L), Chr$(1)) > 0 Then GoTo Nullchar
        RemoveStrings iLine(L), Strings, E
        If iLine(L) = "EndIf" Then iLine(L) = "End If"
        If E <> "" Then
            ReDim LNum(L)
            LNum(L) = L + 1
            GoTo NotMyFault
        End If
    Next L
    
    'Delete comments
    Mode = 0
    For L = 0 To UBound(iLine)
        X = InStr(iLine(L), "'")
        If X > 0 Then iLine(L) = left$(iLine(L), X - 1)
        iLine(L) = Trim$(iLine(L))
    Next L

    'Apply linenumbers
    Temp = String(4, vbNullChar)
    For L = 0 To UBound(iLine)
        iLine(L) = Format(L + 1, "0000000000") & vbNullChar & iLine(L)
    Next L
    
    'Compensate for linebreak character (;)
    Script = Join(iLine, ";")
    Erase iLine
    Script = Replace(Script, "{", "{;")
    Script = Replace(Script, "}", ";};")
    Script = Replace(Script, "|", ";|;")
    Script = ";" & Script
    iLine = Split(Script, ";")
    For L = 1 To UBound(iLine)
        If Asc(Mid$(iLine(L), 11, 1)) = 0 Then
            Temp = Mid$(iLine(L), 1, 11)
        Else
            iLine(L) = Temp & iLine(L)
        End If
    Next L
    
    'Delete whitespace
    X = 0
    For L = 1 To UBound(iLine)
        Do
            If L + X > UBound(iLine) Then Exit For
            If Len(iLine(L + X)) = 11 Then X = X + 1 Else Exit Do
        Loop
        If X > 0 Then iLine(L) = iLine(L + X)
    Next L
    ReDim Preserve iLine(UBound(iLine) - X)
    
    'Read Linenums into memory
    ReDim LNum(1 To UBound(iLine))
    For L = 1 To UBound(iLine)
        LNum(L) = Val(left$(ChopString(iLine(L), 11), 10))
        iLine(L) = Trim$(iLine(L))
        Autofix iLine(L)
    Next L
    
    'Add Variables, Functions, Commands, and Events to the existing database
    For L = 1 To UBound(iLine)
        'VARIABLES
        If left$(iLine(L), 4) = "Dim " Then
            'If Left$(iLine(L), 1) = "@" Or Left$(iLine(L), 1) = "#" Or Left$(iLine(L), 1) = "$" Or Left$(iLine(L), 1) = "%" Then
            'ChopString iLine(L), 4
            Call AddVariable(Right$(iLine(L), Len(iLine(L)) - 4), E)
            If E <> "" Then GoTo NotMyFault
        'FUNCTIONS
        ElseIf left$(iLine(L), 9) = "Function " Then
            Build = iLine(L)
            Temp = ChopString(Build, 9)
            Temp = ChopString(Build, 1)
            ReDim Preserve FuncDefs(UBound(FuncDefs) + 1)
            ReDim Preserve FuncList(UBound(FuncList) + 1)
            Y = UBound(FuncDefs)
            FuncDefs(Y).Type = Symbol2Index(Temp)
            If FuncDefs(Y).Type = vbEmpty Then GoTo InvalidFuncDec
            For X = L To UBound(iLine)
                If iLine(X) = "End Function" Then Exit For
            Next X
            If X = UBound(iLine) + 1 Then GoTo MissingEndFunc
            Z = X + 1
            iLine(X) = iLine(X) & vbNullChar & UBound(FuncList)
            X = InStr(1, Build, "(")
            If X > 0 Then
                If Right(Build, 1) <> ")" Then GoTo MissingPar
                Temp = Mid$(Build, X + 1, Len(Build) - X - 1)
                Build = left(Build, X - 1)
                TempArray = Split(Temp, ",")
                ReDim FuncDefs(Y).Args(UBound(TempArray) + 1)
                For X = 0 To UBound(TempArray)
                    TempArray(X) = Trim(TempArray(X))
                    FuncDefs(Y).Args(X + 1).Type = Symbol2Index(ChopString(TempArray(X), 1))
                    If FuncDefs(Y).Args(X + 1).Type = vbEmpty Then GoTo InvalidFuncDec
                    If Not ValidVariable(TempArray(X)) Then GoTo InvalidVarName
                    FuncDefs(Y).Args(X + 1).Name = TempArray(X)
                Next X
            Else
                Erase FuncDefs(Y).Args
            End If
            If Not ValidVariable(Build) Then GoTo InvalidFuncName
            FuncDefs(Y).Name = Build
            FuncDefs(Y).Index = UBound(FuncList)
            iLine(L) = "Function " & CStr(Y)
            L = Z
        'SUBROUTINES
        ElseIf left$(iLine(L), 8) = "Command " Then
            Build = iLine(L)
            Temp = ChopString(Build, 8)
            ReDim Preserve SubDefs(UBound(SubDefs) + 1)
            ReDim Preserve SubList(UBound(SubList) + 1)
            Y = UBound(SubDefs)
            For X = L To UBound(iLine)
                If iLine(X) = "End Command" Then Exit For
            Next X
            Z = X + 1
            If X = UBound(iLine) + 1 Then GoTo MissingEndSub
            iLine(X) = iLine(X) & " " & vbNullChar & UBound(SubList)
            X = InStr(1, Build, "(")
            If X > 0 Then
                If Right(Build, 1) <> ")" Then GoTo MissingPar
                Temp = Mid$(Build, X + 1, Len(Build) - X - 1)
                Build = left(Build, X - 1)
                TempArray = Split(Temp, ",")
                ReDim SubDefs(Y).Args(1 To UBound(TempArray) + 1)
                For X = 0 To UBound(TempArray)
                    TempArray(X) = Trim(TempArray(X))
                    SubDefs(Y).Args(X + 1).Type = Symbol2Index(ChopString(TempArray(X), 1))
                    If SubDefs(Y).Args(X + 1).Type = vbEmpty Then GoTo InvalidFuncDec
                    If Not ValidVariable(TempArray(X)) Then GoTo InvalidVarName
                    SubDefs(Y).Args(X + 1).Name = TempArray(X)
                Next X
            Else
                Erase SubDefs(Y).Args
            End If
            If Not ValidVariable(Build) Then GoTo InvalidFuncName
            SubDefs(Y).Name = Build
            SubDefs(Y).Index = UBound(SubList)
            iLine(L) = "Command " & CStr(Y)
            L = Z
        'EVENT
        ElseIf left$(iLine(L), 6) = "Event " Then
            Build = iLine(L)
            Temp = ChopString(Build, 6)
            Temp = ChopString(Build, 1)
            For X = 1 To EVENTNUM
                If EventName(X) = Build Then Exit For
            Next X
            If X = EVENTNUM + 1 Then GoTo BadEvent
            X = X * 2
            If Temp = "-" Then
                X = X - 1
            ElseIf Temp <> "+" Then
                GoTo BadEvent
            End If
            For Y = L To UBound(iLine)
                If iLine(Y) = "End Event" Then Exit For
            Next Y
            If Y = UBound(iLine) + 1 Then GoTo MissingEndEvent
            iLine(Y) = iLine(Y) & "   " & vbNullChar & X
            L = Y + 1
            iLine(L) = "Event " & CStr(X)
        Else
            GoTo OutsideEvent
        End If
    Next L
    VarNum = UBound(VarList)
    FuncNum = UBound(FuncDefs)
    
    'And finally, do the big parse.
    Mode = 0
    DeleteFlag = False
    EndFlag = 0
    Erase TempEvent.L
    ReDim TempEvent.L(0)
    For L = 1 To UBound(iLine)
        If Mode = 0 Then
            Select Case ChopString(iLine(L), InStr(1, iLine(L), " "))
            Case "Dim ": Mode = 0
            Case "Event ": Mode = 4
            Case "Command ": Mode = 5
            Case "Function ": Mode = 6
            Case Else
                GoTo OutsideEvent
            End Select
            If Mode > 0 Then
                Erase TempEvent.L
                ReDim TempEvent.L(0)
                Select Case Mode
                Case 4
                Case 5
                    With SubDefs(Val(iLine(L)))
                        If (Not .Args) <> -1 Then
                            For X = 1 To UBound(.Args)
                                Call MakeLocal(.Args(X).Type, .Args(X).Name, 0, -1)
                            Next X
                        End If
                    End With
                Case 6
                    With FuncDefs(Val(iLine(L)))
                        If (Not .Args) <> -1 Then
                            For X = 1 To UBound(.Args)
                                Call MakeLocal(.Args(X).Type, .Args(X).Name, 0, -1)
                            Next X
                        End If
                        Call MakeLocal(.Type, .Name, 0, -1)
                    End With
                End Select
            End If
        ElseIf left$(iLine(L), 6) = "Local " Then
            If Mode < 4 Then GoTo InvalidLocal
            ChopString iLine(L), 6
            ReDim Preserve TempEvent.L(UBound(TempEvent.L) + 1)
            With TempEvent.L(UBound(TempEvent.L))
                .LineNum = LNum(L)
                Call AddVariable(iLine(L), E)
                If E <> "" Then GoTo NotMyFault
                .CommandNum = 5
            End With
            With VarList(UBound(VarList))
                X = UBound(Strings) + 1
                ReDim Preserve Strings(X)
                Strings(X) = FuncDefs(.FuncLink).Name
                Temp = CStr(FuncDefs(.FuncLink).Type)
                Temp = Temp & ", [" & CStr(X) & "]"
                Temp = Temp & ", " & CStr(.LB) & ", " & CStr(.UB)
                TempEvent.L(UBound(TempEvent.L)).MainBlock = ParseBlock(Temp, Strings, E, 4)
                If E <> "" Then GoTo NotMyFault
            End With
        Else
            If Mode > 3 Then Mode = Mode - 3
            ReDim Preserve TempEvent.L(UBound(TempEvent.L) + 1)
            With TempEvent.L(UBound(TempEvent.L))
                .LineNum = LNum(L)
                'First, parse every set of () in the line.
                Do
                    X = InStrRev(iLine(L), "(")
                    If X = 0 Then Exit Do
                    Y = InStr(X, iLine(L), ")")
                    If Y = 0 Then GoTo MissingPar
                    Temp = Mid(iLine(L), X + 1, Y - X - 1)
                    Y = ParseBlock(Temp, Strings, E)
                    If E <> "" Then GoTo NotMyFault
                    iLine(L) = left$(iLine(L), X - 1) & Replace(iLine(L), "(" & Temp & ")", ";" & CStr(Y) & ";", X, 1)
                Loop
                
                'Figure out the command
                X = InStr(1, iLine(L), " ")
                If X = 0 Then
                    Temp = iLine(L)
                    iLine(L) = ""
                Else
                    Temp = Trim$(ChopString(iLine(L), X))
                End If
                iLine(L) = Trim$(iLine(L))
                Y = InStr(1, iLine(L), ";")
                For X = REALSUBSTART To UBound(SubDefs)
                    If SubDefs(X).Name = Temp Then Y = -1: Exit For
                Next X
                If Y = -1 Then
                    .CommandNum = X
                    .MainBlock = ParseBlock(iLine(L), Strings, E, UBound(SubDefs(X).Args))
                    If E <> "" Then GoTo NotMyFault
                ElseIf left$(iLine(L), 1) = "=" Then 'Set
                    .CommandNum = 1
                    ChopString iLine(L), 1
                    iLine(L) = Trim$(iLine(L))
                    If Y > 0 Then
                        Build = ChopString(Temp, Y - 1)
                    Else
                        Build = Temp
                        Temp = ""
                    End If
                    X = GetVarIndex(Build)
                    If X = 0 Then GoTo UndefinedVar
                    If Temp = "" Then Temp = "-1"
                    Temp = CStr(X) & ", " & Temp & ", " & iLine(L)
                    If E <> "" Then GoTo NotMyFault
                    .MainBlock = ParseBlock(Temp, Strings, E, 3)
                    If E <> "" Then GoTo NotMyFault
                Else
                    If Temp = "End" And iLine(L) <> "If" Then
                        Select Case left$(iLine(L), 9)
                        Case "Event   " & vbNullChar, "Command " & vbNullChar, "Function" & vbNullChar
                            Temp = Temp & Trim$(ChopString(iLine(L), 8))
                        Case Else: GoTo SyntaxError
                        End Select
                    End If
                    Select Case Temp
                    Case "If", "While", "For", vbNullChar
                        X = UBound(Level) + 1
                        ReDim Preserve Level(X)
                        Level(X).StartNum = UBound(TempEvent.L)
                        Select Case Temp
                        Case "If": .CommandNum = 3: Level(X).Type = 1
                        Case "While": .CommandNum = 3: Level(X).Type = 2
                        Case "For": .CommandNum = 4: Level(X).Type = 3
                        Case vbNullChar: .CommandNum = 2: Level(X).Type = 4
                        End Select
                        If Right$(iLine(L), 1) = "{" Then
                            Level(X).Brace = True
                            iLine(L) = Trim$(left$(iLine(L), Len(iLine(L)) - 1))
                        End If
                        Level(X).StartLine = iLine(L)
                    Case "}", "End", "Next", "Wend", "Else"
                        Temp = Trim$(Temp & " " & iLine(L))
                        Select Case Temp
                        Case "}": A = 0
                        Case "End If": A = 1
                        Case "Else": A = 2
                        Case "Wend": A = 3
                        Case Else
                            If left$(Temp, 5) = "Next " Then A = 4 Else GoTo SyntaxError
                        End Select
                        Y = UBound(Level)
                        If Y = 0 Then
                            Select Case A
                            Case 0: E = "}:{"
                            Case 1: E = "End If:If"
                            Case 2: E = "Else:If"
                            Case 3: E = "Wend:While"
                            Case 4: E = "Next:For"
                            End Select
                            GoTo WithoutError
                        Else
                            If Level(Y).Brace Then
                                If Temp <> "}" Then
                                    E = "{:}"
                                    GoTo WithoutError
                                ElseIf L <> UBound(iLine) Then
                                    If Level(Y).Type = 1 And iLine(L + 1) = "|" Then iLine(L + 1) = vbNullChar
                                End If
                            Else
                                Select Case Level(Y).Type
                                Case 1
                                    If A <> 1 And A <> 2 Then
                                        E = "End If or Else"
                                    ElseIf A = 2 Then
                                         iLine(L) = vbNullChar
                                    End If
                                Case 2: If A <> 3 Then E = "Wend"
                                Case 3: If A <> 4 Then E = "Next"
                                Case 4: If A <> 1 Then E = "End If"
                                End Select
                                If E <> "" Then GoTo Expected
                            End If
                            
                            A = UBound(TempEvent.L)
                            X = Level(Y).StartNum
                            Build = Level(Y).StartLine
                            Z = L
                            L = TempEvent.L(X).LineNum
                            B = 0
                            Select Case Level(Y).Type
                            Case 1, 4 'If-Else-EndIf
                                If Z <> UBound(iLine) Then
                                    If Temp = "}" And iLine(Z + 1) = vbNullChar Then A = A + 1
                                End If
                                If Temp = "Else" Then A = A + 1
                                If Build <> "" Then Build = Build & ", "
                                Build = Build & CStr(A)
                                TempEvent.L(X).MainBlock = ParseBlock(Build, Strings, E, 2)
                                If E <> "" Then GoTo NotMyFault
                                DeleteFlag = True
                                If Temp = "Else" Then Z = Z - 1
                            Case 2 'While-Wend
                                .CommandNum = 2
                                .MainBlock = ParseBlock(CStr(Level(Y).StartNum), Strings, E)
                                Build = Build & ", " & CStr(A + 1)
                                TempEvent.L(X).MainBlock = ParseBlock(Build, Strings, E, 2)
                                If E <> "" Then GoTo NotMyFault
                            Case 3 'For-Next
                                .CommandNum = 2
                                .MainBlock = ParseBlock(CStr(Level(Y).StartNum), Strings, E)
                                ReDim TempArray(4)
                                X = InStr(1, Build, "=")
                                If X = 0 Then GoTo SyntaxError
                                Piece = Trim$(ChopString(Build, X - 1))
                                If Piece = "" Then GoTo SyntaxError
                                X = GetVarIndex(Piece)
                                If X = 0 Then GoTo UndefinedVar
                                If VarList(X).UB <> -1 Then GoTo ForNextError
                                TempArray(0) = CStr(X)
                                ChopString Build, 1
                                X = InStr(1, Build, "To")
                                If X = 0 Then GoTo SyntaxError
                                TempArray(1) = Trim$(ChopString(Build, X - 1))
                                ChopString Build, 2
                                X = InStr(1, Build, "Step")
                                If X = 0 Then
                                    TempArray(2) = Trim$(Build)
                                    TempArray(3) = "1"
                                Else
                                    TempArray(2) = Trim$(ChopString(Build, X - 1))
                                    ChopString Build, 4
                                    TempArray(3) = Trim$(Build)
                                End If
                                TempArray(4) = CStr(A + 1)
                                For X = 0 To 4
                                    If TempArray(X) = "" Then GoTo SyntaxError
                                Next X
                                Build = Join(TempArray, ", ")
                                TempEvent.L(Level(Y).StartNum).MainBlock = ParseBlock(Build, Strings, E, 5)
                                If E <> "" Then GoTo NotMyFault
                            End Select
                            ReDim Preserve Level(UBound(Level) - 1)
                            L = Z
                        End If
                    Case "EndEvent", "EndCommand", "EndFunction"
                        ChopString iLine(L), 1
                        EndFlag = Val(iLine(L))
                        DeleteFlag = True
                    Case Else
                        GoTo SyntaxError
                    End Select
                End If
            End With
            If DeleteFlag Then
                ReDim Preserve TempEvent.L(UBound(TempEvent.L) - 1)
                DeleteFlag = False
            End If
            If EndFlag > 0 Then
                If Mode > 3 Then Mode = Mode - 3
                Select Case Mode
                Case 1: EventList(EndFlag) = TempEvent
                Case 2: SubList(EndFlag) = TempEvent
                Case 3: FuncList(EndFlag) = TempEvent
                End Select
                Mode = 0
                EndFlag = 0
                ReDim Preserve VarList(VarNum) 'Destroy locals
                ReDim Preserve FuncDefs(FuncNum) 'Destroy locals
            End If
        End If
    Next L
    Erase Strings
    CompileScript = "Script Compiled in " & CStr(Timer - T) & " seconds."
    Exit Function
    
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
ErrorTrap:
    CompileScript = CErr(LNum(L)) & "RTE " & CStr(Err.Number) & " - " & Err.Description
    Exit Function
    Resume
NotMyFault:
    CompileScript = CErr(LNum(L)) & E
    Exit Function
MissingQuote:
    CompileScript = CErr(L) & "Missing quote"
    Exit Function
Nullchar:
    CompileScript = CErr(L) & "Null character detected"
    Exit Function
Expected:
    CompileScript = CErr(LNum(L)) & "Expected: " & E
    Exit Function
ForNextError:
    CompileScript = CErr(LNum(L)) & "Cannot use Arrays in For-Next loops"
    Exit Function
WithoutError:
    TempArray = Split(E, ":")
    CompileScript = CErr(LNum(L)) & TempArray(0) & " without " & TempArray(1)
    Exit Function
SyntaxError:
    CompileScript = CErr(LNum(L)) & "Syntax error"
    Exit Function
MissingPar:
    CompileScript = CErr(LNum(L)) & "Missing parenthesis"
    Exit Function
UndefinedVar:
    CompileScript = CErr(LNum(L)) & "Variable undefined"
    Exit Function
InvalidLocal:
    CompileScript = CErr(LNum(L)) & "Local Variables must come before all other statements"
    Exit Function
InvalidSubscript:
    CompileScript = CErr(LNum(L)) & "Invalid variable subscript"
    Exit Function
InvalidVarName:
    CompileScript = CErr(LNum(L)) & "Invalid variable name"
    Exit Function
InvalidSubName:
    CompileScript = CErr(LNum(L)) & "Invalid command name"
    Exit Function
InvalidFuncName:
    CompileScript = CErr(LNum(L)) & "Invalid function name"
    Exit Function
InvalidSubDec:
    CompileScript = CErr(LNum(L)) & "Invalid command declaration"
    Exit Function
InvalidFuncDec:
    CompileScript = CErr(LNum(L)) & "Invalid function declaration"
    Exit Function
MissingEndSub:
    CompileScript = CErr(LNum(L)) & "Command without End Command"
    Exit Function
MissingEndFunc:
    CompileScript = CErr(LNum(L)) & "Function without End Function"
    Exit Function
MissingEndEvent:
    CompileScript = CErr(LNum(L)) & "Event without End Event"
    Exit Function
OutsideEvent:
    CompileScript = CErr(LNum(L)) & "Statement outside Command, Function, or Event"
    Exit Function
BadEvent:
    CompileScript = CErr(LNum(L)) & "Unrecognized event"
    Exit Function
End Function
Private Function ParseBlock(ByVal BlockText As String, Strings() As String, ByRef ErrorBuffer As String, Optional ByVal MaxArgs = -1, Optional ByVal CheckNested As Boolean = False) As Long
    Dim D As Double
    Dim A As Long
    Dim B As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim R As Long
    Dim TArray() As String
    Dim DataArray() As String
    Dim Build As String
    Dim Temp As String
    If CheckNested Then
        Do
            X = InStrRev(BlockText, "(")
            If X = 0 Then Exit Do
            Y = InStr(X, BlockText, ")")
            If Y = 0 Then GoTo MissingPar
            Temp = Mid$(BlockText, X + 1, Y - X - 1)
            Y = ParseBlock(Temp, Strings, ErrorBuffer)
            If ErrorBuffer <> "" Then Exit Function
            BlockText = left$(BlockText, X - 1) & Replace(BlockText, "(" & Temp & ")", ";" & CStr(Y) & ";", X, 1)
        Loop
    End If
    
    R = UBound(BlockList) + 1
    ReDim Preserve BlockList(R)
    TArray = Split(BlockText, ",")
    X = UBound(TArray)
    If X + 1 > MaxArgs And MaxArgs > -1 Then GoTo TooManyArgs
    If X = -1 Then
        Erase BlockList(R).ArgList
    Else
        ReDim BlockList(R).ArgList(1 To X + 1)
        For X = 1 To X + 1
            Build = Trim$(TArray(X - 1))
            If Build <> "" Then
                Build = Replace(Build, "^", vbNullChar & "01")
                Build = Replace(Build, "*", vbNullChar & "03")
                Build = Replace(Build, "/", vbNullChar & "03" & Chr$(1))
                Build = Replace(Build, "\", vbNullChar & "04")
                Build = Replace(Build, " Mod ", vbNullChar & "05")
                Build = Replace(Build, "+", vbNullChar & "06")
                Build = Replace(Build, "&", vbNullChar & "07")
                Build = Replace(Build, "<>", vbNullChar & "09")
                Build = Replace(Build, "<=", vbNullChar & "12")
                Build = Replace(Build, ">=", vbNullChar & "13")
                Build = Replace(Build, "<", vbNullChar & "10")
                Build = Replace(Build, ">", vbNullChar & "11")
                Build = Replace(Build, "=", vbNullChar & "08")
                'Build = Replace(Build, " Not ", vbNullChar & "14")
                Build = Replace(Build, " And ", vbNullChar & "15")
                Build = Replace(Build, " Xor ", vbNullChar & "17")
                Build = Replace(Build, " Eqv ", vbNullChar & "18")
                Build = Replace(Build, " Or ", vbNullChar & "16")
                'The "-" character is a bit trickier, as it can be either
                'a minus sign or a negation sign.
                Y = 0
                Do
                    Y = InStr(Y + 1, Build, "-")
                    If Y = 0 Then Exit Do
                    If Y > 1 Then Temp = Right$(Trim$(left$(Build, Y - 1)), 3)
                    If left$(Temp, 1) <> vbNullChar And Y <> 1 Then
                        Build = left$(Build, Y - 1) & vbNullChar & "06" & Chr$(1) & Right$(Build, Len(Build) - Y)
                    End If
                Loop
                
                DataArray = Split(vbNullChar & "00" & Build, vbNullChar)
                'With BlockList(R).ArgList(X)
                'Can't use this With here because there's
                'no way around raising RTE10 with it.  Ugh.
                ReDim BlockList(R).ArgList(X).V(1 To UBound(DataArray))
                For Y = 1 To UBound(DataArray)
                    Build = DataArray(Y)
                    BlockList(R).ArgList(X).V(Y).Op = Val(ChopString(Build, 2))
                    Build = Trim$(Build)
                    If left$(Build, 1) = Chr$(1) Then
                        BlockList(R).ArgList(X).V(Y).UseAlt = True
                        ChopString Build, 1
                        Build = Trim$(Build)
                    End If
                    If left$(Build, 4) = "Not " Then
                        BlockList(R).ArgList(X).V(Y).UseNot = True
                        ChopString Build, 4
                        Build = Trim$(Build)
                    End If
                    If left$(Build, 1) = "-" Then
                        BlockList(R).ArgList(X).V(Y).UseNeg = True
                        ChopString Build, 1
                        Build = Trim$(Build)
                    End If
                    If InStr(1, Build, " ") > 0 Then GoTo SyntaxError
                    
                    'And now, it's time to play...
                    'GUESS!  THAT!  DATA TYPE!!
                    B = UBound(Values) + 1
                    'Could it be... a String?!
                    If left$(Build, 1) = "[" And Right$(Build, 1) = "]" Then
                        ReDim Preserve Values(B)
                        Values(B) = Strings(Val(Mid$(Build, 2, Len(Build) - 2)))
                        BlockList(R).ArgList(X).V(Y).Real = True
                        BlockList(R).ArgList(X).V(Y).Link = B
                    'Oh, I know... it's a link to another Block!
                    ElseIf left$(Build, 1) = ";" And Right$(Build, 1) = ";" Then
                        BlockList(R).ArgList(X).V(Y).Real = False
                        BlockList(R).ArgList(X).V(Y).Link = Val(Mid$(Build, 2, Len(Build) - 2))
                    'Maybe it's a Boolean!
                    ElseIf Build = "True" Or Build = "False" Then
                        ReDim Preserve Values(B)
                        Values(B) = CBool(Build)
                        BlockList(R).ArgList(X).V(Y).Real = True
                        BlockList(R).ArgList(X).V(Y).Link = B
                    'What about a Double or a Long?
                    ElseIf IsValidNumber(Build) Then
                        ReDim Preserve Values(B)
                        D = CDbl(Build)
                        If CStr(CLng(D)) = CStr(D) Then
                            Values(B) = CLng(Build)
                        Else
                            Values(B) = CDbl(Build)
                        End If
                        BlockList(R).ArgList(X).V(Y).Real = True
                        BlockList(R).ArgList(X).V(Y).Link = B
                    'Well the only other thing it could be is
                    'a Variable or a Function...
                    Else
                        BlockList(R).ArgList(X).V(Y).Real = False
                        If Right$(Build, 1) = ";" Then
                            Z = InStrRev(Build, ";", Len(Build) - 1)
                            Temp = ChopString(Build, Z - 1)
                            Z = Val(Mid$(Build, 2, Len(Build) - 2))
                        Else
                            Temp = Build
                            Build = ""
                            Z = ParseBlock("", Strings, ErrorBuffer)
                        End If
                        For A = 1 To UBound(FuncDefs)
                            If LCase(FuncDefs(A).Name) = LCase(Temp) Then Exit For
                        Next A
                        If A <= UBound(FuncDefs) Then
                            BlockList(Z).FunctionNum = A
                            BlockList(R).ArgList(X).V(Y).Link = Z
                        'Nothing?!  Grr... *steals prize money anyway*
                        Else
                            GoTo Unrecognized
                        End If
                    End If
                Next Y
                'End With
            End If
        Next X
    End If
    ParseBlock = R
Exit Function
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
TooManyArgs:
    ErrorBuffer = "Wrong number of Arguments"
    Exit Function
Unrecognized:
    ErrorBuffer = "Unrecognized data"
    Exit Function
MissingPar:
    ErrorBuffer = "Missing parenthesis"
    Exit Function
SyntaxError:
    ErrorBuffer = "Syntax error"
    Exit Function
End Function
Private Function IsValidNumber(Expression As String) As Boolean
    Dim X As Long
    For X = 1 To Len(Expression)
        Select Case Mid$(Expression, X, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "."
        Case Else
            IsValidNumber = False
            Exit Function
        End Select
    Next X
    IsValidNumber = True
End Function
Private Function GetVarIndex(VarName As String) As Long
    Dim X As Long
    For X = 1 To UBound(VarList)
        If LCase(FuncDefs(VarList(X).FuncLink).Name) = LCase(VarName) Then Exit For
    Next X
    If X = UBound(VarList) + 1 Then X = 0
    GetVarIndex = X
End Function
Private Function CErr(LineNum As Long) As String
    CErr = "Compile Error in Line #" & CStr(LineNum) & ": "
End Function
Private Function EErr(LineNum As Long) As String
    EErr = "Execution Error in Line #" & CStr(LineNum) & ": "
End Function
Private Sub RemoveStrings(ByRef Source As String, ByRef Target() As String, ErrorBuffer As String)
    Dim X As Long
    Dim Y As Long
    Do
        X = InStr(1, Source, Chr(34))
        If X = 0 Then Exit Do
        Y = InStr(X + 1, Source, Chr(34))
        If Y = 0 Then GoTo MissingQuote
        ReDim Preserve Target(UBound(Target) + 1)
        Target(UBound(Target)) = Mid$(Source, X + 1, Y - X - 1)
        Source = left$(Source, X - 1) & Chr$(0) & CStr(UBound(Target)) & Chr$(1) & Right$(Source, Len(Source) - Y)
    Loop
    If InStr(1, Source, "[") > 0 Or InStr(1, Source, "]") > 0 Then GoTo Bracket
    Source = Replace(Source, Chr$(0), "[")
    Source = Replace(Source, Chr$(1), "]")
    Exit Sub
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
MissingQuote:
    ErrorBuffer = "Missing quote"
    Exit Sub
Bracket:
    ErrorBuffer = "Bracket detected"
    Exit Sub
End Sub
Private Sub ReinsertStrings(ByRef TheLine As String, SArray() As String)
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Temp As String
    Dim TA() As String
    ReDim TA(0)
    X = InStr(1, TheLine, "[")
    If X = 0 Then Exit Sub
    Do
        Y = InStr(X + 1, TheLine, "]")
        Z = UBound(TA) + 1
        ReDim Preserve TA(Z)
        TA(Z) = ChopString(TheLine, Y)
        Y = CLng(Mid$(TA(Z), X + 1, Y - X - 1))
        TA(Z) = left$(TA(Z), X - 1) & Chr$(34) & SArray(Y) & Chr$(34)
        X = InStr(1, TheLine, "[")
    Loop Until X = 0
    Temp = ""
    For X = 1 To UBound(TA)
        Temp = Temp & TA(X)
    Next X
    TheLine = Temp & TheLine
End Sub
Public Function ScriptTest()
    Dim V() As Variant
    Dim E As String
'    ReDim V(1 To 2)
'    V(1) = CDbl(3.14)
'    V(2) = CDbl(2.32)
    Call Execute(1, 36, V, E)
    If E <> "" Then Call ServerWindow.AddMessage(E)
End Function

'Public Sub StringTest()
'    Dim T() As String
'    Dim Temp As String
'    Temp = "This is ""a test"" to ""see"" if the text repl""acement"" system works."
'    Debug.Print Temp
'    Call RemoveStrings(Temp, T)
'    Temp = Right$(Temp, 20)
'    Debug.Print Temp
'    Call ReinsertStrings(Temp, T)
'    Debug.Print Temp
'End Sub
Private Sub SetVariant(ByRef Var As Variant, ByVal DataType As VbVarType)
    Select Case DataType
    Case vbLong: Var = CLng(Var)
    Case vbDouble: Var = CDbl(Var)
    Case vbString: Var = CStr(Var)
    Case vbBoolean: Var = CBool(Var)
    End Select
End Sub
Private Function ValidVariable(ByVal Name As String) As Boolean
    Dim X As Long
    For X = 1 To Len(Name)
        Select Case Asc(Mid$(Name, X, 1))
        Case 65 To 90  'Capital Letters
        Case 97 To 122 'Lowercase Letters
        Case 48 To 57  'Numbers
        Case 95        'Underscore
        Case Else
            ValidVariable = False
            Exit Function
        End Select
    Next X
    ValidVariable = True
End Function
Private Function Symbol2Index(TypeSymbol As String) As VbVarType
    Symbol2Index = vbEmpty
    Select Case TypeSymbol
    Case "@": Symbol2Index = vbBoolean
    Case "#": Symbol2Index = vbDouble
    Case "$": Symbol2Index = vbString
    Case "%": Symbol2Index = vbLong
    Case Else: Symbol2Index = vbEmpty
    End Select
End Function
'Private Function Index2Symbol(TypeSymbol As String) As String
'    Symbol2Index = vbEmpty
'    Select Case TypeSymbol
'    Case vbBoolean: Index2Symbol = "@"
'    Case vbDouble: Index2Symbol = "#"
'    Case vbString: Index2Symbol = "$"
'    Case vbLong: Index2Symbol = "%"
'    End Select
'End Function
Private Sub AddVariable(ByVal Name As String, ErrorBuffer As String)
    Dim Temp As String
    Dim X As Long
    Dim TempArray() As String
    Temp = ChopString(Name, 1)
    ReDim Preserve FuncDefs(UBound(FuncDefs) + 1)
    ReDim Preserve VarList(UBound(VarList) + 1)
    With FuncDefs(UBound(FuncDefs))
        .Type = Symbol2Index(Temp)
        If .Type = vbEmpty Then GoTo InvalidType
        .IsVar = True
        .Index = UBound(VarList)
    End With
    VarList(UBound(VarList)).UB = -1
    X = InStr(1, Name, "(")
    If X > 0 Then
        If Right(Name, 1) <> ")" Then GoTo MissingPar
        Temp = Mid$(Name, X + 1, Len(Name) - X - 1)
        Name = left(Name, X - 1)
        With VarList(UBound(VarList))
            If Temp = "P" Then
                .UB = -2
                ReDim .Value(1 To 255)
            ElseIf Temp = "" Then
                .UB = -3
            ElseIf CStr(Abs(Int(Val(Temp)))) = Temp Then
                .UB = Val(Temp)
                .LB = 0
                ReDim .Value(0 To .UB)
            Else
                TempArray = Split(Temp, " to ")
                If UBound(TempArray) = 1 Then
                    If CStr(Abs(Int(Val(TempArray(0))))) = TempArray(0) And _
                       CStr(Abs(Int(Val(TempArray(1))))) = TempArray(1) Then
                        .LB = Val(TempArray(0))
                        .UB = Val(TempArray(1))
                        ReDim .Value(.LB To .UB)
                    Else
                        GoTo InvalidSubscript
                    End If
                Else
                    GoTo InvalidSubscript
                End If
            End If
        End With
        With FuncDefs(UBound(FuncDefs))
            ReDim .Args(1 To 1)
            .Args(1).Type = vbLong
        End With
    Else
        Erase FuncDefs(UBound(FuncDefs)).Args
        ReDim VarList(UBound(VarList)).Value(0)
    End If
    If ValidVariable(Name) Then
        FuncDefs(UBound(FuncDefs)).Name = Name
        VarList(UBound(VarList)).FuncLink = UBound(FuncDefs)
    Else
        GoTo InvalidVarName
    End If
    Exit Sub
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
InvalidType:
    ErrorBuffer = "Invalid type symbol"
    Exit Sub
InvalidSubscript:
    ErrorBuffer = "Invalid variable subscript"
    Exit Sub
InvalidVarName:
    ErrorBuffer = "Invalid variable name"
    Exit Sub
MissingPar:
    ErrorBuffer = "Missing parenthesis"
    Exit Sub
End Sub
Private Sub MakeLocal(ByVal VarType As VbVarType, ByVal Name As String, ByVal LB As Long, ByVal UB As Long)
    Dim Temp As String
    Dim X As Long
    Dim TempArray() As String
    ReDim Preserve FuncDefs(UBound(FuncDefs) + 1)
    ReDim Preserve VarList(UBound(VarList) + 1)
    With FuncDefs(UBound(FuncDefs))
        .Type = VarType
        .IsVar = True
        .Index = UBound(VarList)
        .Name = Name
    End With
    With VarList(UBound(VarList))
        .UB = UB
        .LB = LB
        .FuncLink = UBound(FuncDefs)
    End With
    If UB = -1 Or UB = -3 Then
        ReDim VarList(UBound(VarList)).Value(0)
        Erase FuncDefs(UBound(FuncDefs)).Args
    Else
        If UB = -2 Then UB = 255: LB = 1
        ReDim VarList(UBound(VarList)).Value(LB To UB)
        With FuncDefs(UBound(FuncDefs))
            ReDim .Args(1 To 1)
            .Args(1).Type = vbLong
            .IsVar = True
            .Index = UBound(VarList)
        End With
    End If
End Sub
Public Function Execute(ByVal SubType As Byte, ByVal Index As Long, ByRef ArgList() As Variant, ByRef ErrorBuffer As String) As Variant
    Dim L As Long
    Dim C As Long
    Dim E As String
    Dim R As Long
    Dim V As Variant
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim F() As ForStruct
    Dim TempEvent As EventType
    Dim Param() As ArgType
    Dim iArgs() As Variant
    
    'Finally, the execution sub!!
    'LET'S DO IT!!!
    Select Case SubType
    Case 0 'Events
        TempEvent = EventList(Index)
        'iargs=
    Case 1 'Commands
        TempEvent = SubList(SubDefs(Index).Index)
        Param = SubDefs(Index).Args
    Case 2 'Functions
        TempEvent = FuncList(FuncDefs(Index).Index)
        Param = FuncDefs(Index).Args
    End Select

    Call SaveLocals

    If (Not Param) <> -1 Then
        For X = 1 To UBound(Param)
            Call MakeLocal(Param(X).Type, Param(X).Name, 0, -1)
        Next X
        Erase Param
    End If
    If (Not ArgList) <> -1 Then
        Y = UBound(ArgList)
        For X = UBound(VarList) To 1 Step -1
            VarList(X).Value(0) = ArgList(Y)
            Y = Y - 1
            If Y = 0 Then Exit For
        Next X
        Erase ArgList
    End If
    
    'Create the Return Variable for Functions.
    If SubType = 2 Then
        Call MakeLocal(FuncDefs(Index).Type, FuncDefs(Index).Name, 0, -1)
        R = UBound(VarList)
    Else
        R = 0
    End If
    ReDim F(0)
    
    'And here she is, the main execution loop
    For L = 1 To UBound(TempEvent.L)
        C = TempEvent.L(L).CommandNum
        If C = 0 Then Stop
        'If C = 6 Then Stop
        If TempEvent.L(L).MainBlock = 0 Then Stop
        With BlockList(TempEvent.L(L).MainBlock)
            If .FunctionNum <> 0 Then Stop
            If (Not .ArgList) <> -1 Then
                Y = UBound(.ArgList)
                ReDim iArgs(1 To Y)
                If (Not SubDefs(C).Args) = -1 Then X = 0 Else X = UBound(SubDefs(C).Args)
                If Y > X Then GoTo TooManyArgs
                If C > SUBLIMIT And Y < X Then GoTo TooManyArgs
                For X = 1 To Y
                    iArgs(X) = EvalArg(.ArgList(X), E)
                    If E <> "" Then GoTo NotMyFault
                    If Not TypeCompare(SubDefs(C).Args(X).Type, iArgs(X)) Then GoTo TypeMismatch
                Next X
            ElseIf (Not SubDefs(C).Args) <> -1 Then
                GoTo TooManyArgs
            End If
        End With
        
        'This big block is for the intrinsic commands
        Select Case C
        Case 1 'Set
            With VarList(iArgs(1))
                X = .UB
                If iArgs(2) = -1 Then
                    If X <> -1 Then GoTo NoIndex
                    X = 0
                Else
                    If X = -1 Then GoTo NoArray
                End If
                If Not TypeCompare(FuncDefs(.FuncLink).Type, iArgs(3)) Then GoTo TypeMismatch
                .Value(X) = iArgs(3)
            End With
        Case 2 'GoTo
            L = iArgs(1) - 1
        Case 3 'If
            If Not iArgs(1) Then L = iArgs(2) - 1
        Case 4 'For
            If F(UBound(F)).LNum <> L Then
                ReDim F(UBound(F) + 1)
                With F(UBound(F))
                    .Var = iArgs(1)
                    .LB = iArgs(2)
                    .UB = iArgs(3)
                    .Stp = iArgs(4)
                    .Jump = iArgs(5)
                    .LNum = L
                    If .Stp = 0 Then GoTo ZeroStep
                    X = FuncDefs(VarList(.Var).FuncLink).Type
                    TypeCompare X, .LB
                    TypeCompare X, .Stp
                    VarList(.Var).Value(0) = .LB - .Stp
                End With
            End If
            With F(UBound(F))
                VarList(.Var).Value(0) = VarList(.Var).Value(0) + .Stp
                X = IIf(VarList(.Var).Value(0) * Sgn(.Stp) > .UB * Sgn(.Stp), 1, 0)
            End With
            If X = 1 Then
                L = F(UBound(F)).Jump - 1
                ReDim Preserve F(UBound(F) - 1)
            End If
        Case 5 'Local
            Call MakeLocal(iArgs(1), iArgs(2), iArgs(3), iArgs(4))
        Case 6 'Print
            Call ServerWindow.AddMessage(CStr(iArgs(1)))
        Case Else
            If C > SUBLIMIT Then
                Call Execute(1, C, iArgs, ErrorBuffer)
                If ErrorBuffer <> "" Then Exit Function
            Else
                Stop
            End If
        End Select
    Next L
    
    If R > 0 Then Execute = VarList(R).Value
    Call RestoreLocals
Exit Function
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
ErrorTrap:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "RTE " & CStr(Err.Number) & " - " & Err.Description
    Exit Function
NotMyFault:
    If SubType <> 2 Then
        ErrorBuffer = EErr(TempEvent.L(L).LineNum) & E
    Else
        ErrorBuffer = E
    End If
    Exit Function
TypeMismatch:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "Type mismatch"
    Exit Function
TooManyArgs:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "Wrong number of Arguments"
    Exit Function
NoIndex:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "Missing array index"
    Exit Function
NoArray:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "Variable is not an Array"
    Exit Function
ZeroStep:
    ErrorBuffer = EErr(TempEvent.L(L).LineNum) & "Step cannot be zero"
    Exit Function
End Function
Private Function EvalArg(T As TermType, ErrorBuffer As String) As Variant
    Dim U As Long
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim V() As Variant
    Dim O() As ValueType
    Dim ThisOne As ValueType
    Dim LastOne As ValueType
    On Error GoTo ErrorTrap
    O = T.V
    ReDim V(1 To UBound(O))
    For X = 1 To UBound(O)
        If T.V(X).Real Then
            V(X) = Values(O(X).Link)
        Else
            V(X) = EvalBlock(O(X).Link, ErrorBuffer)
            If ErrorBuffer <> "" Then Exit Function
        End If
    Next X
    U = UBound(V)
    For X = 1 To 18
        Y = 0
        While Y < U
            Y = Y + 1
            If X = nbsNegat Then
                If O(Y).UseNeg Then
                    If VarType(V(Y)) <> vbDouble And VarType(V(Y)) <> vbLong Then GoTo TypeMismatch
                    V(Y) = -V(Y)
                End If
            ElseIf X = nbsNot Then
                If O(Y).UseNot Then
                    If VarType(V(Y)) <> vbBoolean Then GoTo TypeMismatch
                    V(Y) = Not V(Y)
                End If
            ElseIf X = O(Y).Op Then
                Select Case X
                Case nbsExpon
                    V(Y - 1) = V(Y - 1) ^ V(Y)
                Case nbsMulti
                    If T.V(Y).UseAlt Then
                        V(Y - 1) = V(Y - 1) / V(Y)
                    Else
                        V(Y - 1) = V(Y - 1) * V(Y)
                    End If
                Case nbsFloatDiv
                    V(Y - 1) = V(Y - 1) \ V(Y)
                Case nbsMod
                    V(Y - 1) = V(Y - 1) Mod V(Y)
                Case nbsAdd
                    If T.V(Y).UseAlt Then
                        V(Y - 1) = V(Y - 1) - V(Y)
                    Else
                        V(Y - 1) = V(Y - 1) + V(Y)
                    End If
                Case nbsConcat
                    V(Y - 1) = V(Y - 1) & V(Y)
                Case nbsEqual
                    V(Y - 1) = (V(Y - 1) = V(Y))
                Case nbsInequal
                    V(Y - 1) = (V(Y - 1) <> V(Y))
                Case nbsLessThan
                    V(Y - 1) = (V(Y - 1) < V(Y))
                Case nbsGreaterThan
                    V(Y - 1) = (V(Y - 1) > V(Y))
                Case nbsLessEqual
                    V(Y - 1) = (V(Y - 1) <= V(Y))
                Case nbsGreaterEqual
                    V(Y - 1) = (V(Y - 1) >= V(Y))
                Case nbsAnd
                    V(Y - 1) = V(Y - 1) And V(Y)
                Case nbsOr
                    V(Y - 1) = V(Y - 1) Or V(Y)
                Case nbsXor
                    V(Y - 1) = V(Y - 1) Xor V(Y)
                Case nbsEqv
                    V(Y - 1) = V(Y - 1) Eqv V(Y)
                End Select
                For Z = Y + 1 To UBound(V)
                    V(Z - 1) = V(Z)
                    O(Z - 1) = O(Z)
                Next Z
                U = U - 1
                Y = Y - 1
            End If
        Wend
    Next X
    If U <> 1 Then Stop
    If VarType(V(1)) = vbSingle Then V(1) = CDbl(V(1))
    EvalArg = V(1)
Exit Function
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
ErrorTrap:
    ErrorBuffer = "RTE " & CStr(Err.Number) & " - " & Err.Description
    Exit Function
    Resume
TypeMismatch:
    ErrorBuffer = "Type mismatch"
    Exit Function
End Function
Private Function EvalBlock(Index As Long, ErrorBuffer As String) As Variant
    Dim iArgs() As Variant
    Dim X As Long
    Dim Y As Long
    Dim F As Long
    Dim V As Variant
    On Error GoTo ErrorTrap
    With BlockList(Index)
        F = .FunctionNum
        If (Not .ArgList) <> -1 Then
            Y = UBound(.ArgList)
            ReDim iArgs(1 To Y)
            If (Not FuncDefs(F).Args) = -1 Then X = 0 Else X = UBound(FuncDefs(F).Args)
            If F <= FUNCLIMIT Then
                If Y > X Then
                    GoTo TooManyArgs
                ElseIf FuncDefs(F).Opt > 0 Then
                    If Y < FuncDefs(F).Opt Then GoTo TooManyArgs
                ElseIf Y < X Then
                    GoTo TooManyArgs
                End If
            Else
                If Y <> X Then GoTo TooManyArgs
            End If
            For X = 1 To UBound(.ArgList)
                iArgs(X) = EvalArg(.ArgList(X), ErrorBuffer)
                If ErrorBuffer <> "" Then Exit Function
                If Not TypeCompare(FuncDefs(F).Args(X).Type, iArgs(X)) Then GoTo TypeMismatch
            Next X
        ElseIf (Not FuncDefs(F).Args) <> -1 Then
            GoTo TooManyArgs
        End If
    End With
    Select Case F
    Case 0 '[No Function]
        EvalBlock = iArgs(1)
    Case 1 'CStr
        EvalBlock = CStr(iArgs(1))
    Case 2 'CInt
        EvalBlock = CLng(iArgs(1))
    Case 3 'CFloat
        EvalBlock = CDbl(iArgs(1))
    Case 4 'CBool
        EvalBlock = CBool(iArgs(1))
    Case 5 'Round
        If UBound(iArgs) < 2 Then
            ReDim Preserve iArgs(1 To 2)
            iArgs(2) = 0
        End If
        EvalBlock = CDbl(Round(iArgs(1), iArgs(2)))
    Case 6 'SysTimer
        EvalBlock = CDbl(Timer)
    Case 7 'Random
        If UBound(iArgs) < 2 Then
            ReDim Preserve iArgs(1 To 2)
            iArgs(2) = 0
        End If
        If iArgs(1) > iArgs(2) Then
            X = iArgs(2)
            iArgs(2) = iArgs(1)
            iArgs(1) = X
        End If
        EvalBlock = CLng(Int(Rnd * (iArgs(2) - iArgs(1) + 1)) + iArgs(1))
    Case 8 'Rnd
        EvalBlock = CDbl(Rnd)
    Case 9 'Asc
        EvalBlock = Asc(iArgs(1))
    Case 10 'Chr
        EvalBlock = Chr$(iArgs(1))
    Case 11 'Len
        EvalBlock = CLng(Len(iArgs(1)))
    Case 12 'Search
        If UBound(iArgs) < 3 Then
            ReDim Preserve iArgs(1 To 3)
            iArgs(3) = 1
        End If
        If UBound(iArgs) < 4 Then
            ReDim Preserve iArgs(1 To 4)
            iArgs(3) = False
        End If
        EvalBlock = CLng(InStr(iArgs(3), iArgs(1), iArgs(2), IIf(iArgs(4), vbBinaryCompare, vbTextCompare)))
    Case 13 'Left
        EvalBlock = left$(iArgs(1), iArgs(2))
    Case 14 'Right
        EvalBlock = Right$(iArgs(1), iArgs(2))
    Case 15 'Mid
        EvalBlock = Mid$(iArgs(1), iArgs(2), iArgs(3))
    Case 16 'LCase
        EvalBlock = LCase$(iArgs(1))
    Case 17 'UCase
        EvalBlock = UCase$(iArgs(1))
    Case 18 'Trim
        EvalBlock = Trim$(iArgs(1))
    Case Else
        If F <= FUNCLIMIT Then Stop
        If FuncDefs(F).IsVar Then
            With VarList(FuncDefs(F).Index)
                If .UB = -1 Then
                    If (Not iArgs) <> -1 Then GoTo NoArray
                    ReDim iArgs(1 To 1)
                    iArgs(1) = 0
                Else
                    If (Not iArgs) = -1 Then GoTo NoIndex
                End If
                V = VarList(FuncDefs(F).Index).Value(iArgs(1))
                If V = Empty Then
                    Select Case FuncDefs(F).Type
                    Case vbString: V = ""
                    Case vbBoolean: V = False
                    Case Else: V = 0
                    End Select
                End If
                EvalBlock = V
            End With
        Else
            EvalBlock = Execute(3, FuncDefs(F).Index, iArgs, ErrorBuffer)
            If ErrorBuffer <> "" Then Exit Function
        End If
    End Select
    Exit Function
'--------------------------------------'
'*********** ERROR MESSAGES ***********'
'--------------------------------------'
ErrorTrap:
    If Err.Number = 16 And InVBMode Then Resume
    ErrorBuffer = "RTE " & CStr(Err.Number) & " - " & Err.Description
    Exit Function
    Resume
TypeMismatch:
    ErrorBuffer = "Type mismatch"
    Exit Function
TooManyArgs:
    ErrorBuffer = "Wrong number of Arguments"
    Exit Function
NoIndex:
    ErrorBuffer = "Missing array index"
    Exit Function
NoArray:
    ErrorBuffer = "Variable is not an Array"
    Exit Function
End Function
Private Sub SaveLocals()
    Dim X As Long
    Dim Y As Long
    If (Not Stack) = -1 Then
        ReDim Stack(0)
        Exit Sub
    End If
    ReDim Preserve Stack(UBound(Stack) + 1)
    With Stack(UBound(Stack))
        X = UBound(VarList) - VarNum
        If UBound(FuncDefs) - FuncNum <> X Then Stop
        If X = 0 Then Exit Sub
        ReDim .Vars(1 To X)
        ReDim .Funcs(1 To X)
        For Y = 1 To X
            .Vars(Y) = VarList(VarNum + Y)
            .Funcs(Y) = FuncDefs(FuncNum + Y)
        Next Y
        ReDim Preserve VarList(VarNum)
        ReDim Preserve FuncList(FuncNum)
    End With
End Sub
Private Sub RestoreLocals()
    Dim X As Long
    Dim Y As Long
    ReDim Preserve VarList(VarNum)
    ReDim Preserve FuncList(FuncNum)
    If UBound(Stack) = 0 Then
        Erase Stack
        Exit Sub
    End If
    With Stack(UBound(Stack))
        If (Not .Vars) <> -1 Then
            ReDim Preserve VarList(VarNum + UBound(.Vars))
            ReDim Preserve FuncList(FuncNum + UBound(.Funcs))
            For X = 1 To UBound(.Vars)
                VarList(VarNum + X) = .Vars(X)
                FuncDefs(FuncNum + X) = .Funcs(X)
            Next X
        End If
    End With
    ReDim Preserve Stack(UBound(Stack) - 1)
End Sub
Private Function TypeCompare(ByVal Target As VbVarType, ByRef Source As Variant) As Boolean
    Dim X As VbVarType
    On Error GoTo ETrap
    X = VarType(Source)
    TypeCompare = True
    If Target = vbDouble And X = vbLong Then
        Source = CDbl(Source)
    ElseIf Target = vbLong And X = vbDouble Then
        Source = CLng(Source)
    ElseIf X <> Target And Target <> vbEmpty Then
        TypeCompare = False
    End If
    Exit Function
ETrap:
    'DoEvents
    Resume
End Function
Public Sub Autofix(ByRef TheLine As String)
    Static sDone As Boolean
    Static Pad1() As String
    Static Cap1() As String
    Static Cap2() As String
    Dim Temp As String
    Dim X As Long
    If Not sDone Then
        'These will have a 1 space on each side
        Pad1 = Split("^|*|/|\|+|&|<>|<=|>=|<|>|=", "|")
        'These have to have an exisiting space on both sides to be Capped.
        Cap1 = Split("Mod|Not|And|Or|Xor|Eqv|To|Step", "|")
        For X = 0 To UBound(Cap1)
            Cap1(X) = " " & Cap1(X) & " "
        Next X
        'These have to be at the start of the line
        Cap2 = Split("Event|Command|Function|End Event|End Command|End Function|End If|For|Local|Dim|While|Wend|Next|If", "|")
    End If
    For X = 0 To UBound(Pad1)
        TheLine = CharBuffer(TheLine, Pad1(X))
    Next X
    For X = 0 To UBound(Cap1)
        TheLine = Replace(TheLine, Cap1(X), Cap1(X), , , vbTextCompare)
    Next X
    For X = 0 To UBound(Cap2)
        If LCase(left$(TheLine, Len(Cap2(X)))) = LCase(Cap2(X)) Then
            If Len(TheLine) = Len(Cap2(X)) Then
                TheLine = Cap2(X)
            ElseIf Mid$(TheLine, Len(Cap2(X)) + 1, 1) = " " Then
                Mid(TheLine, 1, Len(Cap2(X))) = Cap2(X)
            End If
        End If
    Next X
End Sub
Private Function CharBuffer(Source As String, Char As String, Optional LBuffer As Integer = 1, Optional RBuffer As Integer = 1) As String
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Build As String
    Build = Source
    X = InStr(1, Build, Char)
    Do Until X = 0
        Z = X + Len(Char)
        If RBuffer <> -1 Then
            For Y = Z To Len(Build)
                If Mid(Build, Y, 1) <> " " Then Exit For
            Next Y
            Build = left(Build, Z - 1) & nSpace(RBuffer) & Right(Build, Len(Build) - Y + 1)
        End If
        If LBuffer <> -1 Then
            For Y = X - 1 To 1 Step -1
                If Mid(Build, Y, 1) <> " " Then Exit For
            Next Y
            Build = left(Build, Y) & nSpace(LBuffer) & Right(Build, Len(Build) - X + 1)
            X = InStr(Y + LBuffer + 2, Build, Char)
        Else
            X = InStr(X + 1, Build, Char)
        End If
    Loop
    CharBuffer = Build
End Function

