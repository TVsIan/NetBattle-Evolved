Attribute VB_Name = "OldScriptMod"

'ScriptEvent Numbers
'-----------------------------------------------------------
'Num   Event            Source       Target        Arg
'1/2   NewMessage       N/A          N/A           [Message]
'1=Before message displayed, 2=After.
'
'3/4   PlayerSignOn     [PNum]       N/A           N/A
'3=Before Sign On message is sent, 4=After.
'
'5/6   PlayerSignOff    [PNum]       N/A           N/A
'5=Before Sign Off message is sent, 6=After.
'
'7/8   BattleOver       [Winner]     [Loser]       {"WIN"|"TIE"}
'7=Before Battle results are sent, 8=After
'
'9/10  BattleBegin      [Challenger] [Challengee]  N/A
'9=Before battle starts, 10=after.
'
'11/12 ChatMessage      [Sender]     N/A           [Message]
'11=Before message is sent, 12=After.
'
'13/14 PlayerAway       [PNum]       N/A           N/A
'13=Before player goes Away, 14=After.
'
'15/16 PlayerKick       [Kicker]     [KickedPNum]  N/A
'15=Before Kick occurs, 16=After.
'
'17/18 PlayerBan        [Banner]     [BannedPNum]  N/A
'17=Before Ban occurs, 18=After.
'
'19/20 TeamChange       [PNum]       N/A           N/A
'19=Before Team change is sent to other users, 20=After
'
'21/22 ChallengeIssued  [Challenger] [Challengee]  N/A
'21=Before challenge is sent, 22=After
'
'23    ServerStartup    N/A          N/A           N/A
'24+   Timer            N/A          N/A           N/A
'
'Variable Symbols
'------------------------------------------------------------
'$ - Text
'# - Number
'@ - Keyword
'! - Varient
'
'Special Commands
'------------------------------------------------------------
'Marker #Num
'^^^Marks the spot for a GoTo statement.  #Num must be an
'   actual value, not a variable.
'
'GoTo #MarkerNum
'^^^Starts execution directly after the specified Marker.
'   #MarkerNum may be a variable.
'
'If [Condition]
'  [Statements1]
'Else
'  [Statements2]
'End If
'^^^Checks the [Condition] for validity.  If it is determined
'   true, [Statements1] executes.  If false, then [Statements2]
'   executs.  If Else and [Statements2] are omitted and
'   [Condition] is false, excution jumps to the line after the
'   EndIf statement.
'
'Conditional Symbols
'-------------------------------------------------------------
'  =       Equal to.
'  ==      Equal to.  (Case Sensitive when dealing with Text)
'  <>      Not equal to.
'  >       Greater than.
'  <       Less than.
'  >=      Greater than or equal to.
'  <=      Less than or equal to.
'
'Condition Operators
'------------------------------------------------------------
'  AND     Returns True if both conditions are True
'  OR      Returns True if either of the conditions are True
'  XOR     Returns True if only one of the conditions is True
'  EQV     Returns True if both conditions are True or False
'
'Regular Functions
'------------------------------------------------------------
'#HasPoke(#PNum, #PokeNum)
'^^^Returns 1 if Player PNum has the Pokemon with
'   the No. PokeNum in his/her team.  Returns 0
'   otherwise.
'
'#HasPokeMove(#PNum, #PokeNum, #MoveNum)
'^^^Returns 1 if Player PNum has the Pokemon with
'   the No. PokeNum and having the move MoveNum in
'   his/her team.  Returns 0 otherwise.
'
'#GetPlayerInfo(#PNum, @Info)
'^^^Returns the specified Player's information
'   depending on the value of @Info.  (Number
'   only)  See below for valid values.
'
'$GetPlayerInfo(#PNum, @Info)
'^^^Returns the specified Player's information
'   depending on the value of @Info.  (Text
'   only)  See below for valid values.
'
'#LineNum
'^^^Returns the number of lines of text in the
'   main message box.
'
'#TrainersNum
'^^^Returns the number of connected Players
'
'#SysTimer
'^^^Returns the number of seconds past midnight.
'
'$Time
'^^^Returns the current time in the form: HH:MM:SS AM/PM
'
'$Date
'^^^Returns the current date in the form: MM/DD/YY
'
'$WeekDay
'^^^Returns a string containing the current day of the week.
'
'$Month
'^^^Returns a string containing the current month.
'
'#PNumber($PName)
'^^^Returns the number of the player matching then
'   name $PName
'
'#Math($MathString)
'^^^Returns the answer to the mathmatical string
'   contained in $MathString
'
'#Rand(#UpperLimit, #LowerLimit)
'^^^Returns a random integer between #LowerLimit
'   and #UpperLimit, inclusive.  If #LowerLimit
'   is omitted, 0 is assumed.
'
'#RandPlayer
'^^^Returns a random player number.  Returns 0 if
'   no players are connected.
'
'#GetValue($Key)
'^^^Retrieves a number from the Windows Registry.
'   An error occurs if the value is not a number.
'
'$GetValue($Key)
'^^^Retrieves a number from the Windows Registry.
'   If the value is a number, it is coverted to
'   text.
'
'$Msg(#Index)
'^^^Returns a Predefined Message, set in the Script
'   Window on the Messages tab.
'
'Settable Functions
'-------------------------------------------------
'#MaxUsers
'^^^Returns/Sets the maximum number of players.
'
'#FloodTol
'^^^Returns/Sets the server's flood tolerance.
'
'$WelcomeMsg
'^^^Returns/Sets the server's welcome message.
'
'String Manipulation Functions:
'-------------------------------------------------
'$Left($Text, #Number)
'^^^Returns the specified number of characters from
'   the left of the text.
'
'$Right($Text, #Number)
'^^^Returns the specified number of characters from
'   the right of the text.
'
'$Mid($Text, #Start, #Length)
'^^^Returns a portion of the text $Text starting
'   at the character specified in #Start as long
'   as #Length characters.
'
'
'#IsIn($Text, $Check, #Case)
'^^^Checks if $Check is located anywhere in $Text.
'   If so, returns the number of characters into
'   $Text that $Check is found.  If not, returns 0.
'   #Case specifies whether or not the check is
'   case sensitive.  0=Not CS, 1=CS.  If omitted,
'   Not CS is assumed.
'
'#Len($Text)
'^^^Returns the number of characters in the text.
'
'$Replace $SourceText, $Find, $Replace
'^^^Searchs for the text $Find in $SourceText and
'   replaces it with $Replace.  Returns the result.
'
'$LCase($Text)
'^^^Puts all the letters in $Text in lower case.
'
'$UCase($Text)
'^^^Puts all the letters in $Text in upper case.
'
'$Chr(#Code)
'^^^Returns the character specified by the ASCII
'   code #Code.  Valid value for #Code are 0 to
'   255.  NOTE: $Chr(1) is reserved for system
'   use.  If you try to use it, it will be
'   replaced with $Chr(2).  Both are meaningless
'   characters.
'
'#Asc($Character)
'^^^Returns the ASCII code for the character.  If
'   the length of $Character is more than 1, the
'   first character is used.
'
'$Str(#Number)
'^^^Returns the specified number in text format.
'
'#Val($Text)
'^^^Returns the numbers in a text statement in
'   number format.
'
'Subroutines:
'-------------------------------------------------
'/? !val
'^^^Show the value of !val in the messages box.
'
'/Set !var, !val
'^^^Set the value of variable !var to !val.  If the
'   variable does not exist, it is created.  If !val
'   is omitted, 0 or "" is assumed, depending on the
'   variable type.
'
'/Unset !var
'^^^Destroys the variable !var.
'
'/Inc #var, #num
'^^^Increments the value of #var by #num.  If #num
'   is omitted, 1 is assumed.  To decrease a number,
'   use a negitive value for #num.
'
'/Clear
'^^^Clears the main messages box
'
'/SendPM #PNum, $Message
'^^^Sends a message to only the player #PNum.
'
'/SendAll $Message
'^^^Sends a message to all connected players.
'
'/Kick #PNum
'^^^Disconnects player #PNum.
'
'/Ban #PNum
'^^^Disconnects player #PNum and adds his/her IP to
'   the banned IP list.
'
'/SIDBan #PNum
'^^^Disconnects player #PNum and adds his/her SID to
'   the banned SID list.
'
'/StopEvent
'^^^Stops the current event from taking place.  Only
'   valid in certain BE event scripts.
'
'/Run $Path
'^^^Runs a the program located at $Path
'
'/SaveVal $Key, !Val
'^^^Saves a value to the Windows Registry.  Advanced
'   users can use this to extend the script's power.
'   Keys are saved to the Visual Basic SaveSetting
'   directory under /NetBattle/Script Values/[$Key]
'
'/SetPlayerInfo #PNum, @Info, !NewVal
'^^^Sets the specified player's infomation as !NewVal.
'   Only certain values can be set.  See below for
'   valid @Info values.
'
'Values for Get/SetPlayerValue's @Info
'------------------------------------------------------
'Value  Type  Notes
'NAME    $    Name.
'IPAD    $    IP Address.  Cannot be Set.
'PSID    $    SID.  Cannot be Set.
'DNSA    $    DNS Address.  Cannot be Set.
'EXTR    $    Extra information.
'VERS    $    NetBattle Version number.  Cannot be Set.
'AUTH    #    Authority.  0=Player, 1=Mod, 2=Admin
'BWTH    #    Battling With (Player Number)  Cannot be Set.
'SPED    #    Connection speed.  Cannot be Set.
'HIDE    #    0=Team Hidden, 1=Team Shown
'WINS    #    Wins.  Cannot be Set.
'LOSE    #    Loses.  Cannot be Set.
'TIES    #    Ties.  Cannot be Set.
'DISC    #    Disconnects.  Cannot be Set.
'
'
'Scripting Module v1.0
'by MasamuneXGP
'
'The scripting module consists of five major functions:
'1. Script Translator (ReRead), 2. Statement Execution (Exec),
'3. Function Evaluation (Eval), 4. Condition Evaluation (IfEval),
'and 5. Block Execution (BlockExec).
'As well as one minor function, Event Initializing (ScriptEvent).

Option Explicit
Option Compare Text
Public Type CallType
    Source As Integer
    Target As Integer
    Arg As String
End Type
Private Type LineType
    Text As String
    LineNum As Long
End Type
Public Type MarkerType
    mName As String
    mLine As Long
End Type
Public Type EventType
    sLine() As LineType
    Marker() As MarkerType
    Counter As Long
    Trigger As Long
End Type
Public Type VarType1
    vName As String
    Value As Single
End Type
Public Type VarType2
    vName As String
    Value As String
End Type
Public Type VarType3
    vName As String
    Value() As Single
End Type
Public Type VarType4
    vName As String
    Value() As String
End Type
Global Const TimerLimit As Integer = 24
Global MainScript As String
Global ProcessScript As Boolean
Global sSource As Integer
Global sTarget As Integer
Global sArg As String
Global PDM() As String
Global AppPath As String
Global VarChange As Boolean
Global NumVar() As VarType1
Global TxtVar() As VarType2
Global PANum() As VarType3
Global PATxt() As VarType4
Global sEvent() As EventType
Global BlankSource As CallType


'***************************************************************'
'-------------BEGIN MAJOR #1: SCRIPT TRANSLATION----------------'
'***************************************************************'
'This function reads the script and puts the lines into
'variables ready for processing.  It returns an error
'description, if there is one.

Public Function Reread(Script As String) As String
    Const FATAL As String = "FATAL SCRIPT ERROR IN LINE #"
    Dim Xs As Single
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim E As String
    Dim EVENTNUM As Integer
    Dim iLine() As String
    Dim cLine As String
    Dim Temp As String
    Dim IfNest() As Integer
    ProcessScript = False
    On Error GoTo ErrorTrap
    ReDim sEvent(TimerLimit - 1)
    For X = 1 To TimerLimit - 1
        ReDim sEvent(X).sLine(0)
        ReDim sEvent(X).Marker(0)
    Next X
    ReDim IfNest(0)
    ReDim sTimer(0)
    iLine() = Split(Script, vbNewLine)
    If UBound(iLine) = -1 Then Exit Function
    For X = 1 To UBound(iLine) + 1
        iLine(X - 1) = LineCheck(Trim(iLine(X - 1)), E)
        If E <> "" Then GoTo NotMyFault
        cLine = iLine(X - 1)
        If left(cLine, 6) = "Event " Then
            If EVENTNUM <> 0 Then GoTo MissingEndEvent
            Temp = Trim(Right(cLine, Len(cLine) - 6))
            Y = 0
            Z = 0
            If left(Temp, 1) = "-" Then
                Y = 1
                Z = 1
            ElseIf left(Temp, 1) = "+" Then
                Z = 1
            End If
            EVENTNUM = GetEventNum(Right(Temp, Len(Temp) - Z))
            If EVENTNUM = TimerLimit Then
                Xs = CSng(Right(Temp, Len(Temp) - 6))
                If Xs <> Int(Xs) Then GoTo BadTimerVal
                Z = CInt(Xs)
                If Z < 1 Or Z > 86400 Then GoTo BadTimerVal2
                For Y = TimerLimit To UBound(sEvent)
                    If sEvent(Y).Trigger = Z Then GoTo RedundantEvent
                Next Y
                ReDim Preserve sEvent(Y)
                ReDim sEvent(Y).sLine(0)
                ReDim sEvent(Y).Marker(0)
                sEvent(Y).Trigger = Z
                sEvent(Y).Counter = 0
                EVENTNUM = Y
            Else
                EVENTNUM = EVENTNUM - Y
                If EVENTNUM < 1 Then GoTo UnknownEvent
                If UBound(sEvent(EVENTNUM).sLine) <> 0 Then GoTo RedundantEvent
            End If
        ElseIf cLine = "EndEvent" Then
            If EVENTNUM = 0 Then GoTo MisplacedEE
            If UBound(IfNest) <> 0 Then GoTo IfWithout
            EVENTNUM = 0
        ElseIf left(cLine, 1) = ":" Then
            If EVENTNUM = 0 Then GoTo ComOutside
            Temp = Right(cLine, Len(cLine) - 1)
            If Not IsAlpha(Temp) Then GoTo BadMarkerVal
            If GetMarkerNum(EVENTNUM, Temp) <> 0 Then GoTo RedundantMarker
            Z = UBound(sEvent(EVENTNUM).Marker) + 1
            ReDim Preserve sEvent(EVENTNUM).Marker(Z)
            sEvent(EVENTNUM).Marker(Z).mName = Temp
            sEvent(EVENTNUM).Marker(Z).mLine = UBound(sEvent(EVENTNUM).sLine)
        ElseIf left(cLine, 3) = "If " Then
            If EVENTNUM = 0 Then GoTo ComOutside
            Y = UBound(sEvent(EVENTNUM).sLine) + 1
            ReDim Preserve sEvent(EVENTNUM).sLine(Y)
            sEvent(EVENTNUM).sLine(Y).Text = cLine
            sEvent(EVENTNUM).sLine(Y).LineNum = X
            Z = UBound(IfNest) + 1
            ReDim Preserve IfNest(Z)
            IfNest(Z) = Y
        ElseIf cLine = "Else" Then
            If EVENTNUM = 0 Then GoTo ComOutside
            If UBound(IfNest) = 0 Then GoTo ElseWithout
            Y = UBound(IfNest)
            Z = UBound(sEvent(EVENTNUM).sLine) + 1
            sEvent(EVENTNUM).sLine(IfNest(Y)).Text = sEvent(EVENTNUM).sLine(IfNest(Y)).Text & "|" & CStr(Z)
            IfNest(Y) = Z
            ReDim Preserve sEvent(EVENTNUM).sLine(Z)
            sEvent(EVENTNUM).sLine(Z).Text = "Else"
            sEvent(EVENTNUM).sLine(Z).LineNum = X
        ElseIf cLine = "EndIf" Then
            If EVENTNUM = 0 Then GoTo ComOutside
            If UBound(IfNest) = 0 Then GoTo EndIfWithout
            Y = UBound(IfNest)
            Z = UBound(sEvent(EVENTNUM).sLine)
            sEvent(EVENTNUM).sLine(IfNest(Y)).Text = sEvent(EVENTNUM).sLine(IfNest(Y)).Text & "|" & CStr(Z)
            ReDim Preserve IfNest(Y - 1)
        ElseIf left(cLine, 2) = "//" Or cLine = "" Then 'do nothing
        Else
            If EVENTNUM = 0 Then GoTo ComOutside
            Y = UBound(sEvent(EVENTNUM).sLine) + 1
            ReDim Preserve sEvent(EVENTNUM).sLine(Y)
            sEvent(EVENTNUM).sLine(Y).Text = cLine
            sEvent(EVENTNUM).sLine(Y).LineNum = X
        End If
    Next X
    If EVENTNUM <> 0 Then GoTo MissingEndEvent
    ProcessScript = True
    Reread = "Script Check: OK!"
Exit Function

'--------------Errors---------------
ErrorTrap:
    Reread = FATAL & X & ": RTE " & Err.Number & ": " & Err.Description & "."
    Exit Function
NotMyFault:
    Reread = FATAL & X & ": " & E
    Exit Function
MissingEndEvent:
    Reread = FATAL & X & ": Missing EndEvent."
    Exit Function
UnknownEvent:
    Reread = FATAL & X & ": Unknown Event."
    Exit Function
RedundantEvent:
    Reread = FATAL & X & ": Redundant Event."
    Exit Function
BadTimerVal:
    Reread = FATAL & X & ": Timer Interval Must Be Integer."
    Exit Function
BadTimerVal2:
    Reread = FATAL & X & ": Timer Interval Out of Bounds."
    Exit Function
MisplacedEE:
    Reread = FATAL & X & ": Misplaced EndEvent."
    Exit Function
ComOutside:
    Reread = FATAL & X & ": Statement Outside Event."
    Exit Function
BadMarkerVal:
    Reread = FATAL & X & ": Marker Name Must Consist of Letters Only."
    Exit Function
RedundantMarker:
    Reread = FATAL & X & ": Redundant Marker."
    Exit Function
MissingEndIf:
    Reread = FATAL & X & ": Missing EndIf."
    Exit Function
IfWithout:
    Reread = FATAL & X & ": If Without EndIf."
    Exit Function
ElseWithout:
    Reread = FATAL & X & ": Else Without If."
    Exit Function
EndIfWithout:
    Reread = FATAL & X & ": EndIf Without If."
    Exit Function
End Function
    
'***************************************************************'
'-------------BEGIN MAJOR #2: STATEMENT EXECUTION---------------'
'***************************************************************'
'This function executes the statement given to it.  If there are
'no errors, this returns "".  If there is an error, this returns
'the error description.

Public Function Exec(Statement As String, CallInfo As CallType, CancelEvent As Boolean) As String
    Dim TextVal() As String
    Dim EvalArg() As String
    Dim Arg() As String
    Dim Build As String
    Dim Build2 As String
    Dim Temp As String
    Dim Temp2 As String
    Dim Char As String
    Dim Pre As String
    Dim E As String
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Xs As Single
    Dim Ys As Single
    On Error GoTo ErrorTrap
    If left(Statement, 1) <> "/" Then GoTo InvalidCommand
    Build = Right(Statement, Len(Statement) - 1)
    E = ""
    
    'Temporarily cut out the text values so they don't interfere.
    TextVal = RemoveText(Build, E)
    If E <> "" Then GoTo MissingQuote
    
    'Parse the arguments
    ReDim EvalArg(0)
    W = 0
    Y = InStr(1, Build, " ")
    If Y = 0 Then Y = Len(Build)
    Pre = Trim(left(Build, Y))
    Temp = Right(Build, Len(Build) - Y)
    Build2 = Temp
    Z = InStr(1, Temp, "(")
    If Z = 0 Then Z = Len(Temp) + 1
    For X = Z To Len(Temp)
        Char = Mid(Temp, X, 1)
        If Char = "(" Then
            W = W + 1
            If W = 1 Then Y = X
        End If
        If Char = ")" Then
            W = W - 1
            If W = 0 Then
                Temp2 = Mid(Temp, Y, X - Y + 1)
                Z = UBound(EvalArg) + 1
                ReDim Preserve EvalArg(Z)
                EvalArg(Z) = Temp2
                Build2 = Replace(Build2, Temp2, "{" & CStr(Z) & "}", 1, 1)
            End If
        End If
    Next X
    If W <> 0 Then GoTo MissingPar
    Arg = ParseArgs(Build2)
    
    'Put the text values back in and evaluate
    For X = 0 To UBound(Arg)
        For Y = 1 To UBound(EvalArg)
            Arg(X) = Replace(Arg(X), "{" & CStr(Y) & "}", EvalArg(Y), 1, 1)
        Next Y
        For Y = 1 To UBound(TextVal)
            Z = InStrRev(Arg(X), Chr(34))
            Arg(X) = left(Arg(X), Z) & Replace(Arg(X), "[" & CStr(Y) & "]", TextVal(Y), Z + 1, 1)
        Next Y
        If X = 1 And Pre = "SetPlayerInfo" Then
            'Do Nothing, it's a constant.
        ElseIf X <> 0 Or (Pre <> "Set" And Pre <> "Inc" And Pre <> "Unset" And Pre <> "SetPA") Then
            Arg(X) = Eval(Arg(X), CallInfo, E)
            If E <> "" Then
                Exec = E
                Exit Function
            End If
            Arg(X) = Replace(Dequote(Arg(X)), Chr(1), Chr(34))
        ElseIf Pre <> "Unset" And Pre <> "SetPA" Then
            Y = InStr(1, Arg(0), "(")
            Z = InStrRev(Arg(0), ")")
            If Y <> 0 Then
                If Z = 0 Then GoTo MissingPar
                Temp = Eval(Mid(Arg(0), Y + 1, Z - Y - 1), CallInfo, E)
                If E <> "" Then
                    Exec = E
                    Exit Function
                End If
                Arg(0) = Replace(Arg(0), Mid(Arg(0), Y + 1, Z - Y - 1), Temp)
            Else
                If Z <> 0 Then GoTo MissingPar
            End If
        End If
    Next X
    
    'Now that all the functions are evaluated,
    'it's time to execute!
    Select Case Pre
    Case "?"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        ServerWindow.AddMessage Arg(0)
    Case "Set"
        If UBound(Arg) = 0 Then
            ReDim Preserve Arg(1)
            If left(Arg(0), 1) = "#" Then
                Arg(1) = "0"
            Else
                Arg(1) = Chr(34) & Chr(34)
            End If
        End If
        If UBound(Arg) <> 1 Then GoTo WrongArgNum
        If Len(Arg(0)) < 2 Then GoTo InvalidArg
        Select Case left(Arg(0), 1)
        Case "$"
            Temp = Enquote(Replace(Arg(1), Chr(34), Chr(1)))
            Y = GetTxtVarNum(Arg(0))
            If Y > UBound(TxtVar) Then
                Z = GetPATxt(Arg(0))
                If Z = 0 Then
                    If Not IsAlpha(Right(Arg(0), Len(Arg(0)) - 1)) Then GoTo BadVar
                    ReDim Preserve TxtVar(Y)
                    TxtVar(Y).vName = Arg(0)
                    TxtVar(Y).Value = Temp
                Else
                    Y = InStr(1, Arg(0), "(")
                    W = InStrRev(Arg(0), ")")
                    If W = 0 Or Y = 0 Then GoTo InvalidArg
                    Temp2 = Eval(Mid(Arg(0), Y + 1, W - Y - 1), CallInfo, E)
                    If E <> "" Then
                        Exec = E
                        Exit Function
                    End If
                    Y = CInt(Temp2)
                    PATxt(Z).Value(Y) = Temp
                End If
            Else
                TxtVar(Y).Value = Temp
            End If
            VarChange = True
        Case "#"
            Xs = CSng(Arg(1))
            Y = GetNumVarNum(Arg(0))
            If Y > UBound(NumVar) Then
                Z = GetPANum(Arg(0))
                If Z = 0 Then
                    If Not IsAlpha(Right(Arg(0), Len(Arg(0)) - 1)) Then GoTo BadVar
                    ReDim Preserve NumVar(Y)
                    NumVar(Y).vName = Arg(0)
                    NumVar(Y).Value = Xs
                Else
                    Y = InStr(1, Arg(0), "(")
                    W = InStrRev(Arg(0), ")")
                    If W = 0 Or Y = 0 Then GoTo InvalidArg
                    Temp2 = Eval(Mid(Arg(0), Y + 1, W - Y - 1), CallInfo, E)
                    If E <> "" Then
                        Exec = E
                        Exit Function
                    End If
                    Y = CInt(Temp2)
                    PANum(Z).Value(Y) = Xs
                End If
            Else
                NumVar(Y).Value = Xs
            End If
            VarChange = True
        Case Else: GoTo InvalidArg
        End Select
    Case "SetPA"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        If Not IsAlpha(Right(Arg(0), Len(Arg(0)) - 1)) Then GoTo BadVar
        Select Case Mid(Arg(0), 1, 1)
        Case "#"
            X = UBound(PANum) + 1
            ReDim Preserve PANum(X)
            ReDim PANum(X).Value(MaxUsers)
            PANum(X).vName = Arg(0)
        Case "$"
            X = UBound(PATxt) + 1
            ReDim Preserve PATxt(X)
            ReDim PATxt(X).Value(MaxUsers)
            For Y = 1 To MaxUsers
                PATxt(X).Value(Y) = Chr(34) & Chr(34)
            Next Y
            PATxt(X).vName = Arg(0)
        Case Else
            GoTo InvalidArg
        End Select
    Case "Unset"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Select Case left(Arg(0), 1)
        Case "#"
            Y = GetNumVarNum(Arg(0))
            W = UBound(NumVar)
            If Y <> W + 1 Then
                For Z = Y To W - 1
                    NumVar(Z) = NumVar(Z + 1)
                Next Z
                ReDim Preserve NumVar(W - 1)
            Else
                Y = GetPANum(Arg(0))
                If Y <> 0 Then
                    Z = UBound(PANum)
                    For X = Y To Z - 1
                        PANum(X) = PANum(X - 1)
                    Next X
                    ReDim Preserve PANum(Z - 1)
                End If
            End If
            VarChange = True
        Case "$"
            Y = GetTxtVarNum(Arg(0))
            W = UBound(TxtVar)
            If Y <> W + 1 Then
                For Z = Y To W - 1
                    TxtVar(Z) = TxtVar(Z + 1)
                Next Z
                ReDim Preserve TxtVar(W - 1)
            Else
                Y = GetPATxt(Arg(0))
                If Y <> 0 Then
                    Z = UBound(PATxt)
                    For X = Y To Z - 1
                        PATxt(X) = PATxt(X - 1)
                    Next X
                    ReDim Preserve PATxt(Z - 1)
                End If
            End If
            VarChange = True
        Case Else
            GoTo InvalidArg
        End Select
    Case "Inc"
        If UBound(Arg) = 0 Then ReDim Preserve Arg(1): Arg(1) = "1"
        If UBound(Arg) <> 1 Then GoTo WrongArgNum
        Xs = CSng(Arg(1))
        If left(Arg(0), 1) <> "#" Then Y = CInt("a")
        Y = GetNumVarNum(Arg(0))
        If Y = UBound(NumVar) + 1 Then
            Y = GetPANum(Arg(0))
            If Y = 0 Then GoTo NoSuchVar
            Z = InStrRev(Arg(0), ")")
            W = InStr(1, Arg(0), "(")
            If Z = 0 Or W = 0 Then GoTo InvalidArg
            Z = CInt(Mid(Arg(0), W + 1, Z - W - 1))
            If E <> "" Then
                Exec = E
                Exit Function
            End If
            PANum(Y).Value(Z) = PANum(Y).Value(Z) + Xs
        Else
            NumVar(Y).Value = NumVar(Y).Value + Xs
        End If
        VarChange = True
    Case "Clear"
        ServerWindow.Messages.Text = ""
    Case "SendPM"
        If UBound(Arg) <> 1 Then GoTo WrongArgNum
        Y = CInt(Arg(0))
        If Not IsLoaded(Y) Then GoTo PlayerNotConnected
        ServerWindow.AddToQueue Y, "CMSG:" & Arg(1)
    Case "SendAll"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        ServerWindow.SendAll "CMSG:" & Arg(0)
        ServerWindow.AddMessage Arg(0)
    Case "Kick"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Y = CInt(Arg(0))
        If Not IsLoaded(Y) Then GoTo PlayerNotConnected
        Player(Y).DCReason = "Scripted kick."
        ServerWindow.AddToQueue Y, "BOOT:"
        If CallInfo.Source = Y Then CancelEvent = True
    Case "Ban"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Y = CInt(Arg(0))
        If Not IsLoaded(Y) Then GoTo PlayerNotConnected
        ServerDB.BanUser Y
        If CallInfo.Source = Y Then CancelEvent = True
    Case "SIDBan"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Y = CInt(Arg(0))
        If Not IsLoaded(Y) Then GoTo PlayerNotConnected
        ServerDB.SIDBanUser Y
        If CallInfo.Source = Y Then CancelEvent = True
    Case "TempBan"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Y = CInt(Arg(0))
        If Not IsLoaded(Y) Then GoTo PlayerNotConnected
        ServerWindow.Tempban Y
        If CallInfo.Source = Y Then CancelEvent = True
    Case "Run"
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Shell Arg(0), vbNormalFocus
    Case "SaveValue"
        If UBound(Arg) <> 1 Then GoTo WrongArgNum
        SaveSetting "NetBattle", "Script Values", Arg(0), Arg(1)
    Case "SetPlayerInfo"
        If UBound(Arg) <> 2 Then GoTo WrongArgNum
        X = CInt(Arg(0))
        If Not IsLoaded(X) Then GoTo PlayerNotConnected
        Select Case Arg(1)
        Case "EXTR"
            Player(X).Extra = CorrectText(Arg(2))
        Case "AUTH"
            If CInt(Arg(2)) > 2 Or CInt(Arg(2)) < 0 Then GoTo InvalidArg
            Player(X).Authority = CInt(Arg(2)) + 1
            ServerWindow.SendAll "AUTH:" & FixedHex(X, 3) & Player(X).Authority
            ServerDB.ChgAuth Player(X).Name, Player(X).Authority
        Case "HIDE"
            Player(X).ShowTeam = Not (CInt(Arg(2)) = 0)
        Case Else
            GoTo InvalidArg
        End Select
        ServerWindow.RefreshListing
        ServerWindow.SendAll ServerWindow.PreparePlayerData(X, True)
    Case Else
        GoTo InvalidCommand
    End Select
        
Exit Function
'--------------Errors---------------
ErrorTrap:
    Exec = "RTE " & Err.Number & ": " & Err.Description & "."
    Exit Function
NoSuchVar:
    Exec = "No Such Variable."
    Exit Function
BadVar:
    Exec = "Variable Name Must Include Only Letters."
    Exit Function
MissingPar:
    Exec = "Missing Parenthesis."
    Exit Function
MissingQuote:
    Exec = "Missing: " & Chr(34)
    Exit Function
WrongArgNum:
    Exec = "Wrong Number of Arguments."
    Exit Function
PlayerNotConnected:
    Exec = "No Such Player Number."
    Exit Function
InvalidArg:
    Exec = "Invalid Argument."
    Exit Function
TextOutside:
    Exec = "Text Outside Quotes."
    Exit Function
InvalidCommand:
    Exec = "Invalid Command."
    Exit Function
End Function

'***************************************************************'
'-------------BEGIN MAJOR #3: FUNCTION EVALUATION---------------'
'***************************************************************'
'This function evaluates all the script functions and math, as
'well as all nested functions, and returns the final value.  If
'there is an error while processing, it is stored in ErrorBuffer.

Public Function Eval(Source As String, CallInfo As CallType, ErrorBuffer As String, Optional DoVars As Boolean = True) As String
    Dim B As Boolean
    Dim V As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Xl As Long
    Dim Yl As Long
    Dim Xs As Single
    Dim Ys As Single
    Dim Zs As Single
    Dim E As String
    Dim A() As Integer
    Dim F() As String
    Dim Arg() As String
    
    Dim Pre As String
    Dim Temp As String
    Dim Build As String
    Dim NewVal As String
    Dim TextVal() As String
    Dim FinalVal As String
    'This is the main processing function.  It evaluates all script functions.
    On Error GoTo ErrorTrap
    If Source = Enquote("") Then
        Eval = Source
        Exit Function
    End If
    If IsStrictNumeric(Source) Then
        Eval = Source
        Exit Function
    End If
    If left(Source, 1) = Chr(34) And InStr(2, Source, Chr(34)) = Len(Source) Then
        Eval = Source
        Exit Function
    End If
    Build = Source
    E = ""
    X = InStr(1, Build, "$Message")
    Do Until X = 0
        If Not (IsInQuotes(Build, X) Or IsAlpha(Mid(Build, X + 8, 1))) Then
            Temp = Enquote(Replace(CallInfo.Arg, Chr(34), Chr(1)))
            Build = left(Build, X - 1) & Replace(Build, "$Message", Temp, X, 1)
        End If
        X = InStr(X + 1, Build, "$Message")
    Loop
    
    'Temporarily cut out the text values so they don't interfere.
    TextVal = RemoveText(Build, E)
    If E <> "" Then GoTo MissingQuote
    
    If DoVars Then
        'Replace any variables with their current values.
        For X = 1 To UBound(NumVar)
            Z = Len(NumVar(X).vName)
            Y = InStr(1, Build, NumVar(X).vName)
            Do Until Y = 0
                If Not IsAlpha(Mid(Build, Y + Z, 1)) Then
                    Build = left(Build, Y - 1) & Replace(Build, NumVar(X).vName, CStr(NumVar(X).Value), Y, 1)
                End If
                Y = InStr(Y + 1, Build, NumVar(X).vName)
            Loop
        Next X
        For X = 1 To UBound(TxtVar)
            Z = Len(TxtVar(X).vName)
            Y = InStr(1, Build, TxtVar(X).vName)
            Do Until Y = 0
                If Not IsAlpha(Mid(Build, Y + Z, 1)) Then
                    W = UBound(TextVal) + 1
                    ReDim Preserve TextVal(W)
                    TextVal(W) = TxtVar(X).Value
                    Build = left(Build, Y - 1) & Replace(Build, TxtVar(X).vName, "[" & CStr(W) & "]", Y, 1)
                End If
                Y = InStr(Y + 1, Build, TxtVar(X).vName)
            Loop
        Next X
        Build = Replace(Build, "#Source", CStr(CallInfo.Source))
        Build = Replace(Build, "#Target", CStr(CallInfo.Target))
        Arg = NoArgFunctions
        For X = 1 To UBound(Arg)
            Build = Replace(Build, Arg(X), Arg(X) & "()")
        Next X
        Erase Arg
    End If
    
    'Get rid of those pesky parenthesis...
    If Mid(Build, Len(Build), 1) <> ")" Or (left(Build, 1) <> "$" And left(Build, 1) <> "#" And left(Build, 1) <> "(") Then
        Build = "(" & Build & ")"
    End If
    E = ""
    Do Until InStr(1, Build, ")") = Len(Build)
        Z = InStr(1, Build, ")")
        Y = InStrRev(Build, "(", Z)
        If Y = 0 Then GoTo MissingPar
        X = InStrRev(Build, " ", Y)
        W = InStrRev(Build, "(", Y - 1)
        If W > X Then X = W
        Temp = Mid(Build, X + 1, Z - X)
        Pre = Temp
        ReDim A(0)
        For X = 1 To UBound(TextVal)
            If InStr(1, Temp, "[" & CStr(X) & "]") Then
                ReDim A(UBound(A) + 1)
                A(UBound(A)) = X
                Temp = Replace(Temp, "[" & CStr(X) & "]", TextVal(X), 1, 1)
                TextVal(X) = ""
            End If
        Next X
        If Temp = Source And Not DoVars Then GoTo SyntaxError 'Kills those nasty StackSpace errors
        E = ""
        NewVal = Eval(Temp, CallInfo, E, False)
        If E <> "" Then
            ErrorBuffer = E
            Exit Function
        End If
        If left(NewVal, 1) = Chr(34) Then
            For X = 1 To UBound(TextVal)
                If TextVal(X) = "" Then Exit For
            Next X
            If X > UBound(TextVal) Then ReDim Preserve TextVal(X)
            TextVal(X) = NewVal
            NewVal = "[" & CStr(X) & "]"
        End If
        Build = Replace(Build, Pre, NewVal, 1, 1)
        If Mid(Build, Len(Build), 1) <> ")" Or (left(Build, 1) <> "$" And left(Build, 1) <> "#" And left(Build, 1) <> "(") Then
            Build = "(" & Build & ")"
        End If
    Loop
    
    
    'Now that that's taken care of, we've got a single function with
    'actual values for the arguments (if any).  So lets get to work.
    Y = InStr(1, Build, "(")
    If Y = 0 Then GoTo UnknownFunc
    Pre = left(Build, Y - 1)
    Arg = ParseArgs(Mid(Build, Y + 1, Len(Build) - Y - 1))
    
    'Put the text values back in.  This one in particular is tougher
    'than in the other functions, since the [X] markers can be out of
    'order.
    For X = 0 To UBound(Arg)
        For Y = 1 To UBound(TextVal)
            W = 0
            Do
                W = InStr(W + 1, Arg(X), "[" & CStr(Y) & "]")
                B = Not IsInQuotes(Arg(X), W)
            Loop Until B Or W = 0
            If W <> 0 Then
                Z = InStrRev(Arg(X), Chr(34), W)
                Arg(X) = left(Arg(X), Z) & Replace(Arg(X), "[" & CStr(Y) & "]", TextVal(Y), Z + 1, 1)
            End If
        Next Y
    Next X

    'Now we do any math or string joins
    For X = 0 To UBound(Arg)
        If left(Arg(X), 1) = Chr(34) Then
            Arg(X) = Replace(Arg(X), Enquote(" & "), "")
            If left(Build, 1) = "#" Or left(Build, 1) = "$" Then Arg(X) = Dequote(Arg(X))
        Else
            ReDim F(4)
            F(1) = "*"
            F(2) = "/"
            F(3) = "+"
            F(4) = "-"
            For W = 1 To 3 Step 2
                Y = InStr(3, Arg(X), F(W))
                Z = InStr(3, Arg(X), F(W + 1))
                Do Until Y = 0 And Z = 0
                    If Y = 0 Then Y = Len(Arg(X)) + 1
                    If Z = 0 Then Z = Len(Arg(X)) + 1
                    If Y > Z Then Y = Z
                    If Y < 3 Then GoTo InvalidArg
                    Z = InStrRev(Arg(X), " ", Y - 2)
                    Temp = Mid(Arg(X), Z + 1, Y - Z - 2)
                    Xs = CSng(Temp)
                    Z = InStr(Y + 2, Arg(X), " ")
                    If Z = 0 Then Z = Len(Arg(X)) + 1
                    NewVal = Mid(Arg(X), Y + 2, Z - Y - 2)
                    Ys = CStr(NewVal)
                    Select Case Mid(Arg(X), Y, 1)
                    Case "*"
                        Zs = Xs * Ys
                        Arg(X) = Replace(Arg(X), Temp & " * " & NewVal, CStr(Zs))
                    Case "/"
                        Zs = Xs / Ys
                        Arg(X) = Replace(Arg(X), Temp & " / " & NewVal, CStr(Zs))
                    Case "+"
                        Zs = Xs + Ys
                        Arg(X) = Replace(Arg(X), Temp & " + " & NewVal, CStr(Zs))
                    Case "-"
                        Zs = Xs - Ys
                        Arg(X) = Replace(Arg(X), Temp & " - " & NewVal, CStr(Zs))
                    End Select
                    Y = InStr(3, Arg(X), F(W))
                    Z = InStr(3, Arg(X), F(W + 1))
                Loop
            Next W
        End If
    Next X
    
    Select Case Mid(Build, 1, 1)
    Case "#" 'Number Function
        Select Case Pre
        Case "#IsLoaded"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            FinalVal = CStr(Abs(CInt(IsLoaded(Z))))
        Case "#HasPoke"
            Z = CInt(Arg(0))
            Y = CInt(Arg(1))
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            FinalVal = "0"
            For X = 1 To 6
                If Player(Z).PKMN(X) = Y Then FinalVal = "1"
            Next X
        Case "#HasPokeMove"
            If UBound(Arg) <> 2 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            Y = CInt(Arg(1))
            W = CInt(Arg(2))
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            FinalVal = "0"
            For X = 1 To 6
                If Player(Z).PKMN(X) = Y Then Exit For
            Next X
            If X <> 7 Then
                For V = 1 To 4
                    If Player(Z).PokeData(X).Move(V) = W Then FinalVal = "1"
                Next V
            End If
        Case "#GetTeamPoke"
            Z = CInt(Arg(0))
            Y = CInt(Arg(1))
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            FinalVal = CStr(Player(Z).PKMN(Y))
        Case "#GetPlayerInfo"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            Select Case Arg(1)
            Case "AUTH": FinalVal = CStr(Player(Z).Authority - 1)
            Case "BWTH": FinalVal = IIf(Player(Z).BattlingWith <> 1025, CStr(Player(Z).BattlingWith), "0")
            Case "SPED": FinalVal = CStr(Player(Z).Speed)
            Case "HIDE": FinalVal = CStr(Abs(CInt(Player(Z).ShowTeam)))
            Case "WINS": FinalVal = CStr(Player(Z).Wins)
            Case "LOSE": FinalVal = CStr(Player(Z).Losses)
            Case "TIES": FinalVal = CStr(Player(Z).Ties)
            Case "DISC": FinalVal = CStr(Player(Z).Disconnect)
            Case Else: GoTo InvalidArg
            End Select
        Case "#GetCompat"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            Select Case Player(Z).GameVersion
            Case nbTrueRBY: FinalVal = "0"
            Case nbRBYTrade: FinalVal = "1"
            Case nbTrueGSC: FinalVal = "2"
            Case nbGSCTrade: FinalVal = "3"
            Case nbTrueRuSa: FinalVal = "4"
            Case nbFullAdvance: FinalVal = "5"
            End Select
        Case "#LineNum"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = CStr(UBound(Split(ServerWindow.Messages.Text, vbNewLine)))
        Case "#TrainersNum"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = CStr(ServerWindow.ListView1.ListItems.count)
        Case "#SysTimer"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = CStr(Timer)
        Case "#PNumber"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(ServerWindow.GetNumber(Arg(0)))
        Case "#Rand"
            If UBound(Arg) = 0 Then ReDim Preserve Arg(1): Arg(1) = "0"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            Xl = CLng(Arg(0))
            Yl = CLng(Arg(1))
            FinalVal = CStr(Int(Rnd * (Xl - Yl + 1) + Yl))
        Case "#RandPlayer"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            Z = ServerWindow.ListView1.ListItems.count
            If Z = 0 Then
                FinalVal = "0"
            Else
                Z = Int(Rnd * Z) + 1
                Temp = ServerWindow.ListView1.ListItems(Z).Key
                FinalVal = Right(Temp, Len(Temp) - 5)
            End If
        Case "#MaxUsers"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = CStr(MaxUsers)
        Case "#FloodTol"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = CStr(FloodTolerance)
        Case "#IsIn"
            If UBound(Arg) = 1 Then ReDim Preserve Arg(2): Arg(2) = "0"
            If UBound(Arg) <> 2 Then GoTo WrongArgNum
            If Arg(2) <> "0" And Arg(2) <> "1" Then GoTo InvalidArg
            X = Abs(Not -CInt(Arg(2)))
            FinalVal = CStr(InStr(1, Arg(0), Arg(1), X))
        Case "#Len"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(Len(Arg(0)))
        Case "#Asc"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(Asc(left(Arg(0), 1)))
        Case "#Val"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(Val(Arg(0)))
        Case "#GetValue"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = GetSetting("NetBattle", "Script Values", Arg(0), 0)
        Case "#Round"
            If UBound(Arg) = 0 Then ReDim Preserve Arg(1): Arg(1) = "0"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            FinalVal = CStr(Round(CSng(Arg(0)), Val(Arg(1))))
        Case "#PokeNum"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(GetPokeNum(Arg(0)))
        Case "#MoveNum"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(GetMoveNum(Arg(0)))
        Case "#ItemNum"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(GetItemNum(Arg(0)))
        Case "#GetPokeLevel"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            FinalVal = CStr(Player(CInt(Arg(0))).PokeData(CInt(Arg(1))).Level)
        Case "#GetPokeItem"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            FinalVal = CStr(Player(CInt(Arg(0))).PokeData(CInt(Arg(1))).Item)
        
        Case Else
            Y = GetPANum(Pre)
            If Y = 0 Then GoTo UnknownFunc
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            FinalVal = CStr(PANum(Y).Value(Z))
        End Select
        Eval = FinalVal
        
        
    Case "$" 'Text Function
        Select Case Pre
        Case "$GetPlayerInfo"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            If Not IsLoaded(Z) Then GoTo PlayerNotConnected
            Select Case Arg(1)
            Case "NAME": FinalVal = Player(Z).Name
            Case "IPAD": FinalVal = Player(Z).Address
            Case "PSID": FinalVal = Player(Z).SID
            Case "DNSA": FinalVal = Player(Z).DNSAddress
            Case "EXTR": FinalVal = Player(Z).Extra
            Case "VERS": FinalVal = Player(Z).Version
            Case Else: GoTo InvalidArg
            End Select
        Case "$Name"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            FinalVal = Player(Z).Name
        Case "$Time"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = Format(Now, "HH:MM:SS AMPM")
        Case "$Date"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = Format(Now, "MM/DD/YY")
        Case "$WeekDay"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = Format(Now, "DDDD")
        Case "$Month"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = Format(Now, "MMMM")
        Case "$WelcomeMsg"
            If UBound(Arg) <> -1 Then GoTo WrongArgNum
            FinalVal = ServerMessage
        Case "$Left"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            X = CInt(Arg(1))
            FinalVal = left(Arg(0), X)
        Case "$Right"
            If UBound(Arg) <> 1 Then GoTo WrongArgNum
            X = CInt(Arg(1))
            FinalVal = Right(Arg(0), X)
        Case "$Mid"
            If UBound(Arg) <> 2 Then GoTo WrongArgNum
            X = CInt(Arg(1))
            Y = CInt(Arg(2))
            FinalVal = Mid(Arg(0), X, Y)
        Case "$Replace"
            If UBound(Arg) <> 2 Then GoTo WrongArgNum
            FinalVal = Replace(Arg(0), Arg(1), Arg(2))
        Case "$LCase"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = LCase(Arg(0))
        Case "$UCase"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = UCase(Arg(0))
        Case "$Chr"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            X = CInt(Arg(0))
            If X = 1 Then X = 2  'Due to their nature of screwing things up,
            If X = 34 Then X = 1 'quotes will be treated as Chr(1) for now.
            FinalVal = Chr$(X)
        Case "$Str"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = CStr(CSng(Arg(0)))
        Case "$Msg"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            X = CInt(Arg(0))
            FinalVal = Eval(PDM(X), CallInfo, E)
            If E <> "" Then
                ErrorBuffer = E
                Exit Function
            End If
            FinalVal = Dequote(FinalVal)
        Case "$GetValue"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = GetSetting("NetBattle", "Script Values", Arg(0), "")
        Case "$Pokemon"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            If Z = 0 Then Z = -1
            FinalVal = BasePKMN(Z).Name
        Case "$Move"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = Moves(CInt(Arg(0))).Name
        Case "$Item"
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            FinalVal = Item(CInt(Arg(0)))
        Case Else
            Y = GetPATxt(Pre)
            If Y = 0 Then GoTo UnknownFunc
            If UBound(Arg) <> 0 Then GoTo WrongArgNum
            Z = CInt(Arg(0))
            FinalVal = Dequote(PATxt(Y).Value(Z))
        End Select
        Eval = Enquote(FinalVal)
    
    
    
    Case Else 'No Function
        If UBound(Arg) <> 0 Then GoTo WrongArgNum
        Eval = Arg(0)
    End Select
    
    If InStr(1, Eval, Chr(34)) <> 0 Then
        If left(Eval, 1) <> Chr(34) Or InStr(2, Eval, Chr(34)) <> Len(Eval) Then
            Eval = ""
            GoTo TextOutside
        End If
    Else
        If Not IsNumeric(Eval) Then
            Eval = ""
            GoTo SyntaxError
        Else
            Eval = CStr(CSng(Eval))
        End If
    End If
Exit Function

'--------------Errors---------------
ErrorTrap:
    ErrorBuffer = "RTE " & Err.Number & ": " & Err.Description & "."
    Exit Function
SyntaxError:
    ErrorBuffer = "Syntax Error."
    Exit Function
MissingPar:
    ErrorBuffer = "Missing: ("
    Exit Function
MissingQuote:
    ErrorBuffer = "Missing: " & Chr(34)
    Exit Function
UnrecTag:
    ErrorBuffer = "Unrecognized Function Tag."
    Exit Function
UnknownFunc:
    ErrorBuffer = "Unrecongnized Function."
    Exit Function
WrongArgNum:
    ErrorBuffer = "Wrong Number of Arguments."
    Exit Function
PlayerNotConnected:
    ErrorBuffer = "No Such Player Number."
    Exit Function
InvalidArg:
    ErrorBuffer = "Invalid Argument."
    Exit Function
TextOutside:
    ErrorBuffer = "Text Outside Quotes."
    Exit Function
End Function

'***************************************************************'
'-------------BEGIN MAJOR #4: CONDITION EVALUATION--------------'
'***************************************************************'
'This function is actually two functions, IfEval and IfEvalKernel.
'IfEval parses the entire string of conditions and then applies
'the AND/OR/XOR/EQV operators.  IfEvalKernel actually evaluates
'the individual conditions for IfEval.

Public Function IfEval(Source As String, CallInfo As CallType, ErrorBuffer As String) As Boolean
    Dim Build As String
    Dim Build2 As String
    Dim TextVal() As String
    Dim Arg() As String
    Dim Results() As Boolean
    Dim E As String
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Op As String
    On Error GoTo ErrorTrap
    Build = Source
    TextVal = RemoveText(Build, E)
    If E <> "" Then
        ErrorBuffer = E
        Exit Function
    End If
    Op = " " & Chr(1) & " "
    Build = Replace(Build, " OR ", " ORO ") 'Oro? It's a lot easier when they're
    Build2 = Replace(Build, " ORO ", Op)  'all three letters, that it is. =P
    Build2 = Replace(Build2, " AND ", Op)
    Build2 = Replace(Build2, " XOR ", Op)
    Build2 = Replace(Build2, " EQV ", Op)
    Arg = Split(Build2, Op)
    ReDim Results(UBound(Arg))
    For X = 0 To UBound(Arg)
        For Y = 1 To UBound(TextVal)
            Z = InStrRev(Arg(X), Chr(34))
            Arg(X) = left(Arg(X), Z) & Replace(Arg(X), "[" & CStr(Y) & "]", TextVal(Y), Z + 1, 1)
        Next Y
        Results(X) = IfEvalKernel(Arg(X), CallInfo, E)
        If E <> "" Then
            ErrorBuffer = E
            Exit Function
        End If
    Next X
    
    Z = Len(Build2) + 1
    For X = UBound(Arg) To 1 Step -1
        Z = InStrRev(Build2, Chr(1), Z - 1)
        Op = Mid(Build, Z + (X - 1) * 2, 3)
        Select Case Op
        Case "AND"
            Results(X - 1) = Results(X - 1) And Results(X)
        Case "ORO"
            Results(X - 1) = Results(X - 1) Or Results(X)
        Case "XOR"
            Results(X - 1) = Results(X - 1) Xor Results(X)
        Case "EQV"
            Results(X - 1) = Results(X - 1) Eqv Results(X)
        End Select
    Next X
    IfEval = Results(0)
    Exit Function
ErrorTrap:
    ErrorBuffer = "RTE " & Err.Number & ": " & Err.Description & "."
    Exit Function
End Function
Private Function IfEvalKernel(Source As String, CallInfo As CallType, ErrorBuffer As String) As Boolean
    Dim Build As String
    Dim Build2 As String
    Dim TextVal() As String
    Dim Arg() As String
    Dim Op As String
    Dim E As String
    Dim Xs As Single
    Dim Ys As Single
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    On Error GoTo ErrorTrap
    Build = Source
    TextVal = RemoveText(Build, E)
    If E <> "" Then
        ErrorBuffer = E
        Exit Function
    End If
    Build2 = Replace(Build, "==", Chr(1))
    Build2 = Replace(Build2, ">=", Chr(1))
    Build2 = Replace(Build2, "<=", Chr(1))
    Build2 = Replace(Build2, "<>", Chr(1))
    Build2 = Replace(Build2, "=", Chr(1))
    Build2 = Replace(Build2, ">", Chr(1))
    Build2 = Replace(Build2, "<", Chr(1))
    Z = InStr(1, Build2, Chr(1))
    If Z = 0 Then GoTo Invalid
    Op = Trim(Mid(Build, Z, 2))
    ReDim Arg(1)
    Arg(0) = Trim(left(Build2, Z - 1))
    Arg(1) = Trim(Right(Build2, Len(Build2) - Z))
    For X = 0 To UBound(Arg)
        For Y = 1 To UBound(TextVal)
            Z = InStrRev(Arg(X), Chr(34))
            Arg(X) = left(Arg(X), Z) & Replace(Arg(X), "[" & CStr(Y) & "]", TextVal(Y), Z + 1, 1)
        Next Y
    Next X
    Arg(1) = Eval(Arg(1), CallInfo, E)
    Arg(0) = Eval(Arg(0), CallInfo, E)
    If E <> "" Then
        ErrorBuffer = E
        Exit Function
    End If
    If (left(Arg(0), 1) = Chr(34) Xor left(Arg(1), 1) = Chr(34)) Then Y = CInt("a")
    If left(Arg(0), 1) = Chr(34) Then
        Select Case Op
        Case "="
            IfEvalKernel = Not CBool(StrComp(Arg(0), Arg(1), vbTextCompare))
        Case "=="
            IfEvalKernel = Not CBool(StrComp(Arg(0), Arg(1), vbBinaryCompare))
        Case "<>"
            IfEvalKernel = (Arg(0) <> Arg(1))
        Case "<", ">", ">=", "<="
            Err.Raise 13
        Case Else
            GoTo Invalid
        End Select
    Else
        X = CSng(Arg(0))
        Y = CSng(Arg(1))
        Select Case Op
        Case "=", "=="
            IfEvalKernel = (X = Y)
        Case ">"
            IfEvalKernel = (X > Y)
        Case "<"
            IfEvalKernel = (X < Y)
        Case "<>"
            IfEvalKernel = (X <> Y)
        Case ">="
            IfEvalKernel = (X >= Y)
        Case "<="
            IfEvalKernel = (X <= Y)
        End Select
    End If
    Exit Function
ErrorTrap:
    ErrorBuffer = "RTE " & Err.Number & ": " & Err.Description & "."
    Exit Function
Invalid:
    ErrorBuffer = "Invalid Comparison."
    Exit Function
End Function

'***************************************************************'
'-------------BEGIN MAJOR #5: BLOCK EXECUTION-------------------'
'***************************************************************'
'This is the main routine that executes the blocks of code.

Public Sub BlockExec(ByVal EVENTNUM As Integer, Optional ByRef Cancel As Boolean, Optional ByVal Source As Integer = 0, Optional ByVal Target As Integer = 0, Optional ByVal Arg As String = "")
    Const NF As String = "SCRIPT ERROR IN LINE #"
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Temp As String
    Dim E As String
    Dim Result As Boolean
    Dim cLine As String
    Dim TextVal() As String
    Dim CallInfo As CallType
    On Error GoTo ErrorTrap
    If Not ProcessScript Then Exit Sub
    CallInfo.Source = Source
    CallInfo.Target = Target
    CallInfo.Arg = Arg
    sArg = Arg
    With sEvent(EVENTNUM)
        For X = 1 To UBound(.sLine)
            cLine = .sLine(X).Text
            If left(cLine, 5) = "GoTo " Then
                Y = GetMarkerNum(EVENTNUM, Right(cLine, Len(cLine) - 5))
                If Y = 0 Then GoTo NoMarker
                X = .Marker(Y).mLine
            ElseIf cLine = "/StopEvent" Then
                Select Case EVENTNUM
                Case 1, 11, 13, 15, 17, 21
                Case Else
                    GoTo InvalidSE
                End Select
                Cancel = True
            ElseIf cLine = "/Exit" Then
                Exit Sub
            ElseIf left(cLine, 3) = "If " Then
                Y = InStrRev(cLine, "|")
                If Y = 0 Then Stop
                Temp = Mid(cLine, 4, Y - 4)
                Z = CInt(Right(cLine, Len(cLine) - Y))
                Result = IfEval(Temp, CallInfo, E)
                If E <> "" Then GoTo NotMyFault
                If Not Result Then X = Z
            ElseIf left(cLine, 5) = "Else|" Then
                Z = CInt(Right(cLine, Len(cLine) - 5))
                X = Z
            ElseIf left(cLine, 1) = "/" Then
                E = Exec(cLine, CallInfo, Cancel)
                If E <> "" Then GoTo NotMyFault
            Else
                GoTo NoCommand
            End If
        Next X
    End With
    Exit Sub
ErrorTrap:
    Temp = NF & sEvent(EVENTNUM).sLine(X).LineNum & ": "
    ServerWindow.AddMessage (Temp & "RTE " & Err.Number & ": " & Err.Description & ".")
    Exit Sub
NotMyFault: 'Bet you were wondering where all those ErrorBuffers were leading eh?
    Temp = NF & sEvent(EVENTNUM).sLine(X).LineNum & ": "
    ServerWindow.AddMessage (Temp & E)
    Exit Sub
NoMarker:
    Temp = NF & sEvent(EVENTNUM).sLine(X).LineNum & ": "
    ServerWindow.AddMessage (Temp & "No Such Marker.")
    Exit Sub
InvalidSE:
    Temp = NF & sEvent(EVENTNUM).sLine(X).LineNum & ": "
    ServerWindow.AddMessage (Temp & "Invalid /StopEvent")
    Exit Sub
NoCommand:
    Temp = NF & sEvent(EVENTNUM).sLine(X).LineNum & ": "
    ServerWindow.AddMessage (Temp & "Unknown Command.")
    Exit Sub
End Sub


'***************************************************************'
'-------------BEGIN MINOR #1: SCRIPT INITIALIZATION-------------'
'***************************************************************'
'This is called at the start of the server, to read the script
'from the saved file and reset all the variables.
Public Function ScriptInit() As String
    Dim B As Boolean
    Dim Z As Integer
    Dim Y As Long
    Dim X As Integer
    Dim Temp As String
    Dim iLine() As String
    AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    Z = FreeFile
    If Dir(AppPath & "Script.ini") = "" Then
        Open AppPath & "Script.ini" For Output As #Z
        Close #Z
    End If
    ReDim iLine(0)
    X = 0
    Open AppPath & "Script.ini" For Input As #Z
    Do Until EOF(Z)
        ReDim Preserve iLine(X)
        Line Input #Z, iLine(X)
        X = X + 1
    Loop
    Close #Z
    MainScript = Join(iLine, vbNewLine)
    ReDim NumVar(0)
    ReDim TxtVar(0)
    ReDim PANum(0)
    ReDim PATxt(0)
    ReDim PANum(0).Value(MaxUsers)
    ReDim PANum(0).Value(MaxUsers)
    If Dir(AppPath & "NVariables.ini") = "" Then
        Open AppPath & "NVariables.ini" For Output As #Z
        Close #Z
    End If
    X = 0
    Y = 0
    Open AppPath & "NVariables.ini" For Input As #Z
    Do Until EOF(Z)
        Line Input #Z, Temp
        If Temp = "" Then
            Y = 1
        Else
            Select Case Y
            Case 0
                Select Case left(Temp, 1)
                Case "#"
                    ReDim Preserve PANum(UBound(PANum) + 1)
                    ReDim PANum(UBound(PANum)).Value(MaxUsers)
                    PANum(UBound(PANum)).vName = Trim(left(Temp, 15))
                Case "$"
                    ReDim Preserve PATxt(UBound(PATxt) + 1)
                    ReDim PATxt(UBound(PATxt)).Value(MaxUsers)
                    PATxt(UBound(PATxt)).vName = Trim(left(Temp, 15))
                End Select
            Case 1
                X = X + 1
                ReDim Preserve NumVar(X)
                NumVar(X).vName = Trim(left(Temp, 15))
                NumVar(X).Value = CSng(Right(Temp, Len(Temp) - 15))
            End Select
        End If
    Loop
    Close #Z
    If Dir(AppPath & "TVariables.ini") = "" Then
        Open AppPath & "TVariables.ini" For Output As #Z
        Close #Z
    End If
    X = 0
    Y = 0
    Open AppPath & "TVariables.ini" For Binary As #Z
    Do Until 1 = 2
        X = X + 1
        Temp = nSpace(4)
        Get #Z, , Temp
        Temp = nSpace(Val("&H" & Temp))
        If Len(Temp) = 0 Then Exit Do
        Get #Z, , Temp
        ReDim Preserve TxtVar(X)
        TxtVar(X).vName = Trim(left(Temp, 15))
        TxtVar(X).Value = Right(Temp, Len(Temp) - 15)
    Loop
    Close #Z
    ReDim PDM(0)
    If Dir(AppPath & "Messages.ini") = "" Then
        Open AppPath & "Messages.ini" For Output As #Z
        'Print #Z, ""
        Close #Z
    End If
    X = 0
    Open AppPath & "Messages.ini" For Input As #Z
    Do Until EOF(Z)
        ReDim Preserve PDM(UBound(PDM) + 1)
        Line Input #Z, PDM(UBound(PDM))
    Loop
    Close #Z
    For X = UBound(PDM) To 1 Step -1
        If PDM(X) = "" Then ReDim Preserve PDM(X - 1) Else Exit For
    Next X
    ServerWindow.AddMessage "Script files loaded successfully."
    Temp = Reread(MainScript)
    If Temp <> "" Then ServerWindow.AddMessage Temp
End Function
'***************************************************************'
'-------------BEGIN MINOR #2: AUTOMATIC LINE CHECK--------------'
'***************************************************************'
'This function takes a line of code and automatically fixes
'certain syntax errors, such as too many or too few spaces
'between statements. Just like a REAL programming editor! =D
Public Function LineCheck(Source As String, ErrorBuffer As String, Optional DoCaps As Boolean = False) As String
    Dim X As Integer
    Dim Y As Integer
    Dim TextVal() As String
    Dim Build As String
    Dim E As String
    Dim Char As String
    If Source = "" Then
        LineCheck = ""
        Exit Function
    End If
    If left(Trim(Source), 2) = "//" Or left(Trim(Source), 1) = ":" Then
        'Comment or Marker
        LineCheck = Source
        Exit Function
    End If
    Build = Source
    TextVal = RemoveText(Build, E)
    If E <> "" Then
        ErrorBuffer = E
        Exit Function
    End If
    Build = CharBuffer(Build, "[", 1, -1)
    Build = CharBuffer(Build, "]", -1, 1)
    Build = CharBuffer(Build, ",", 0, 1)
    Build = CharBuffer(Build, ")", 0, -1)
    Build = CharBuffer(Build, "(", -1, 0)
    Build = left(Build, 1) & CharBuffer(Right(Build, Len(Build) - 1), "/")
    Build = CharBuffer(Build, "&")
    Build = CharBuffer(Build, "*")
    Build = CharBuffer(Build, "+")
    Build = CharBuffer(Build, "-")
    Build = CharBuffer(Build, "=")
    Build = CharBuffer(Build, ">")
    Build = CharBuffer(Build, "<")
    Build = Replace(Build, "= =", "==")
    Build = Replace(Build, "< =", "<=")
    Build = Replace(Build, "> =", ">=")
    Build = Replace(Build, "< >", "<>")
    Build = Replace(Build, "> <", "<>")
    Build = Replace(Build, "+ -", "-")
    Build = Replace(Build, "- -", "+")
    Build = Replace(Build, "Event - ", "Event -")
    Build = Replace(Build, "Event + ", "Event +")
    If DoCaps Then
        If left(Build, 6) = "Event " Then
            Mid(Build, 1, 6) = "Event "
            Build = Replace(Build, "newmessage", "NewMessage")
            Build = Replace(Build, "playersignon", "PlayerSignOn")
            Build = Replace(Build, "playersignoff", "PlayerSignOff")
            Build = Replace(Build, "battleover", "BattleOver")
            Build = Replace(Build, "battlebegin", "BattleBegin")
            Build = Replace(Build, "chatmessage", "ChatMessage")
            Build = Replace(Build, "playeraway", "PlayerAway")
            Build = Replace(Build, "playerkick", "PlayerKick")
            Build = Replace(Build, "playerban", "PlayerBan")
            Build = Replace(Build, "teamchange", "TeamChange")
            Build = Replace(Build, "challengeissued", "ChallengeIssued")
            Build = Replace(Build, "serverstartup", "ServerStartup")
            Build = Replace(Build, "timer", "Timer")
        ElseIf Build = "end if" Or Build = "endif" Then
            Build = "EndIf"
        ElseIf Build = "else" Then
            Build = "Else"
        ElseIf Build = "endevent" Or Build = "end event" Then
            Build = "EndEvent"
        Else
            If left(Build, 5) = "goto " Then Mid(Build, 1, 4) = "GoTo "
            If left(Build, 3) = "if " Then Mid(Build, 1, 3) = "If "
            Build = Replace(Build, " and ", " AND ")
            Build = Replace(Build, " or ", " OR ")
            Build = Replace(Build, " xor ", " XOR ")
            Build = Replace(Build, " eqv ", " EQV ")
            Build = Replace(Build, "#isloaded", "#IsLoaded")
            Build = Replace(Build, "#haspoke", "#HasPoke")
            Build = Replace(Build, "#haspokemove", "#HasPokeMove")
            Build = Replace(Build, "#getteampoke", "#GetTeamPoke")
            Build = Replace(Build, "#getplayerinfo", "#GetPlayerInfo")
            Build = Replace(Build, "#linenum", "#LineNum")
            Build = Replace(Build, "#trainersnum", "#TrainersNum")
            Build = Replace(Build, "#systimer", "#SysTimer")
            Build = Replace(Build, "#pnumber", "#PNumber")
            Build = Replace(Build, "#rand", "#Rand")
            Build = Replace(Build, "#randplayer", "#RandPlayer")
            Build = Replace(Build, "#maxusers", "#MaxUsers")
            Build = Replace(Build, "#floodtol", "#FloodTol")
            Build = Replace(Build, "#isin", "#IsIn")
            Build = Replace(Build, "#len", "#Len")
            Build = Replace(Build, "#asc", "#Asc")
            Build = Replace(Build, "#val", "#Val")
            Build = Replace(Build, "#getvalue", "#GetValue")
            Build = Replace(Build, "#getcompat", "#GetCompat")
            Build = Replace(Build, "#source", "#Source")
            Build = Replace(Build, "#target", "#Target")
            Build = Replace(Build, "$message", "$Message")
            Build = Replace(Build, "$getplayerinfo", "$GetPlayerInfo")
            Build = Replace(Build, "$name", "$Name")
            Build = Replace(Build, "$time", "$Time")
            Build = Replace(Build, "$date", "$Date")
            Build = Replace(Build, "$weekday", "$WeekDay")
            Build = Replace(Build, "$month", "$Month")
            Build = Replace(Build, "$welcomemsg", "$WelcomeMsg")
            Build = Replace(Build, "$left", "$Left")
            Build = Replace(Build, "$right", "$Right")
            Build = Replace(Build, "$mid", "$Mid")
            Build = Replace(Build, "$replace", "$Replace")
            Build = Replace(Build, "$lcase", "$LCase")
            Build = Replace(Build, "$ucase", "$UCase")
            Build = Replace(Build, "$chr", "$Chr")
            Build = Replace(Build, "$str", "$Str")
            Build = Replace(Build, "$msg", "$Msg")
            Build = Replace(Build, "$getvalue", "$GetValue")
            Build = Replace(Build, "$pokemon", "$Pokemon")
            Build = Replace(Build, "/set", "/Set")
            Build = Replace(Build, "/setpa", "/SetPA")
            Build = Replace(Build, "/unset", "/Unset")
            Build = Replace(Build, "/inc", "/Inc")
            Build = Replace(Build, "/clear", "/Clear")
            Build = Replace(Build, "/sendpm", "/SendPM")
            Build = Replace(Build, "/sendall", "/SendAll")
            Build = Replace(Build, "/kick", "/Kick")
            Build = Replace(Build, "/ban", "/Ban")
            Build = Replace(Build, "/sidban", "/SIDBan")
            Build = Replace(Build, "/tempban", "/TempBan")
            Build = Replace(Build, "/run", "/Run")
            Build = Replace(Build, "/savevalue", "/SaveValue")
            Build = Replace(Build, "/setplayerinfo", "/SetPlayerInfo")
            Build = Replace(Build, "/stopevent", "/StopEvent")
            Build = Replace(Build, "/exit", "/Exit")
            Build = Replace(Build, ", name", ", NAME")
            Build = Replace(Build, ", ipad", ", IPAD")
            Build = Replace(Build, ", psid", ", PSID")
            Build = Replace(Build, ", dnsa", ", DNSA")
            Build = Replace(Build, ", extr", ", EXTR")
            Build = Replace(Build, ", vers", ", VERS")
            Build = Replace(Build, ", auth", ", AUTH")
            Build = Replace(Build, ", bwth", ", BWTH")
            Build = Replace(Build, ", sped", ", SPED")
            Build = Replace(Build, ", hide", ", HIDE")
            Build = Replace(Build, ", wins", ", WINS")
            Build = Replace(Build, ", lose", ", LOSE")
            Build = Replace(Build, ", ties", ", TIES")
            Build = Replace(Build, ", disc", ", DISC")
        End If
    End If
    For X = 1 To UBound(TextVal)
        Y = InStrRev(Build, Chr(34))
        Build = left(Build, Y) & Replace(Build, "[" & CStr(X) & "]", TextVal(X), Y + 1, 1)
    Next X
    LineCheck = Trim(Build)
End Function





'--------------------------------------------------------
'And down here are private helper functions.
Private Function CharBuffer(Source As String, Char As String, Optional LBuffer As Integer = 1, Optional RBuffer As Integer = 1)
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
Private Function Enquote(Text As String) As String
    Enquote = Chr(34) & Text & Chr(34)
End Function
Private Function NoArgFunctions() As String()
    Dim T(11) As String
    T(1) = "#LineNum"
    T(2) = "#TrainersNum"
    T(3) = "$Time"
    T(4) = "#SysTimer"
    T(5) = "$WeekDay"
    T(6) = "$Month"
    T(7) = "#RandPlayer"
    T(8) = "#MaxUsers"
    T(9) = "#FloodTol"
    T(10) = "$WelcomeMsg"
    T(11) = "$Date"
    NoArgFunctions = T
End Function
Private Function ParseArgs(Source As String) As String()
    Dim Args() As String
    Dim X As Integer
    Args = Split(Source, ",")
    For X = 0 To UBound(Args)
        Args(X) = Trim(Args(X))
    Next X
    ParseArgs = Args
End Function
Private Function GetEventNum(eName As String) As Integer
    Dim X As Integer
    '1, 11, 13, 15, 17, 21
    Select Case eName
        Case "newmessage": X = 2
        Case "playersignon": X = 4
        Case "playersignoff": X = 6
        Case "battleover": X = 8
        Case "battlebegin": X = 10
        Case "chatmessage": X = 12
        Case "playeraway": X = 14
        Case "playerkick": X = 16
        Case "playerban": X = 18
        Case "teamchange": X = 20
        Case "challengeissued": X = 22
        Case "serverstartup": X = 23
        Case Else
            If left(eName, 6) = "timer " Then X = TimerLimit Else X = 0
    End Select
    GetEventNum = X
End Function
Private Function GetTxtVarNum(iName As String) As Integer
    Dim X As Integer
    For X = 1 To UBound(TxtVar)
        If TxtVar(X).vName = iName Then Exit For
    Next X
    GetTxtVarNum = X
End Function
Private Function GetNumVarNum(iName As String) As Integer
    Dim X As Integer
    For X = 1 To UBound(NumVar)
        If NumVar(X).vName = iName Then Exit For
    Next X
    GetNumVarNum = X
End Function
Private Function GetPANum(iName As String) As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim T As String
    T = iName
    X = InStr(1, iName, "(")
    Y = InStr(1, iName, "{")
    If Y <> 0 Or X <> 0 Then
        If Y = 0 Then Y = X + 1
        If X = 0 Then X = Y + 1
        If X > Y Then X = Y
        T = left(T, X - 1)
    End If
    For Z = 1 To UBound(PANum)
        If PANum(Z).vName = T Then Exit For
    Next Z
    If Z = UBound(PANum) + 1 Then Z = 0
    X = InStrRev(iName, ")")
    Y = InStrRev(iName, "}")
    If X < Y Then X = Y
    If X <> 0 And Len(iName) <> X Then Z = 0
    GetPANum = Z
End Function
Private Function GetPATxt(iName As String) As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim T As String
    T = iName
    X = InStr(1, iName, "(")
    Y = InStr(1, iName, "{")
    If Y <> 0 Or X <> 0 Then
        If Y = 0 Then Y = X + 1
        If X = 0 Then X = Y + 1
        If X > Y Then X = Y
        T = left(T, X - 1)
    End If
    For X = 1 To UBound(PATxt)
        If PATxt(X).vName = T Then Exit For
    Next X
    If X = UBound(PATxt) + 1 Then X = 0
    GetPATxt = X
End Function
Private Function IsAlpha(Text As String) As Boolean
    Dim X As Integer
    Dim Y As Integer
    If Text = "" Then IsAlpha = False Else IsAlpha = True
    For X = 1 To Len(Text)
        Y = Asc(Mid(Text, X, 1))
        If Y < 65 Or Y > 122 Then IsAlpha = False
        If Y < 97 And Y > 90 Then IsAlpha = False
    Next X
End Function
Public Function IsStrictNumeric(Text As String) As Boolean
    Dim X As Double
    Dim Y As Double
    On Error Resume Next
    Err.Number = 0
    X = CDbl(Text)
    Y = Val(Text)
    IsStrictNumeric = (Err.Number = 0 And X = Y)
End Function

Private Function RemoveText(Build As String, ErrorBuffer As String) As String()
    Dim TextVal() As String
    Dim Temp As String
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    ReDim TextVal(0)
    X = InStr(1, Build, Chr(34))
    Do Until X = 0
        Y = InStr(X + 1, Build, Chr(34))
        If Y = 0 Then
            ErrorBuffer = "Missing: " & Chr(34)
            RemoveText = TextVal
            Exit Function
        End If
        Z = UBound(TextVal) + 1
        Temp = Mid(Build, X, Y - X + 1)
        ReDim Preserve TextVal(Z)
        TextVal(Z) = Temp
        Build = Replace(Build, Temp, "[" & CStr(Z) & "]", 1, 1)
        X = InStr(1, Build, Chr(34))
    Loop
    RemoveText = TextVal
End Function
Private Function IsInQuotes(Source As String, Number As Integer) As Boolean
    Dim X As Integer
    Dim B As Boolean
    B = True
    Do
        X = InStr(X + 1, Source, Chr(34))
        B = Not B
    Loop Until X = 0 Or X > Number
    IsInQuotes = B
End Function

Public Sub SaveVariables()
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim Temp As String
    Dim T2 As String
    Dim vName As String
    On Error GoTo Failed
    Z = FreeFile
    Open AppPath & "NVariables.ini" For Output As #Z
    For X = 1 To UBound(PANum)
        Print #Z, PANum(X).vName
    Next X
    For X = 1 To UBound(PATxt)
        Print #Z, PATxt(X).vName
    Next X
    Print #Z, ""
    For X = 1 To UBound(NumVar)
        Print #Z, Pad(NumVar(X).vName, 15) & CStr(NumVar(X).Value)
    Next X
    Close #Z
    Temp = ""
    For X = 1 To UBound(TxtVar)
        T2 = FixedHex(15 + Len(TxtVar(X).Value), 4)
        T2 = T2 & TxtVar(X).vName & nSpace(15 - Len(TxtVar(X).vName)) & TxtVar(X).Value
        Temp = Temp & T2
    Next X
    Open AppPath & "TVariables.ini" For Output As #Z
    Print #Z, Temp
    Close #Z
Failed:
End Sub
Private Function GetMarkerNum(EVENTNUM As Integer, iName As String) As Integer
    Dim X As Integer
    Dim Y As Integer
    Y = 0
    With sEvent(EVENTNUM)
        For X = 1 To UBound(.Marker)
            If .Marker(X).mName = iName Then Y = X: Exit For
        Next X
    End With
    GetMarkerNum = Y
End Function


