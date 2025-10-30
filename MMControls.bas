Attribute VB_Name = "ModuleMultimedia"
'Version 6.1
'This copy which sent to MSDN library
'Module MPEG,AVI,sequencer,audio source code
'This code written by abdullah al-ahdal e-mail:a_ahdal@yahoo.com
'for planet source code & for All who want the best and easist deal
'with Multimedia
'I written this code (standard code) to make the best and the easist
'dealing with multimedia file (All types) By pure Windows API.
'In this Module ready functions to use it in your projects
'Just add this code to your project and you will have the
'easist way to Controlling with multimedia files.just you
'must know how can call these functions from this module
'and how you can deal with it "if it success or not"
'All Functions in this Module will return a value
'if the Function success or not.

'Special Thanks to:
'1-Janet because he solved the problem of File name and the
'Path
'2-also to Alex because he told me for some bugs in this code and all of it was repaired.
'3-also to Aaron Wilkes.
'4-also to Hans de Vries for notice me about the bug when playing rmi files

'For any request Contact to me at : a_ahdal@yahoo.com

Option Explicit

'Private Declares
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Private types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'variable just in the module
Dim glo_from As Long
Dim glo_to As Long
Dim glo_AliasName As String
Dim glo_hWnd As Long

Public Function OpenMultimedia(hWnd As Long, AliasName As String, FileName As String, typeDevice As String) As String
'Callig OpenMultimedia will open the multimedia file
'Parameters
'hWnd
'[in]handle of the window
'which you want to play in. you can put handle for
'your desktop if you want to playing movie in your desktop.

'AliasName
'[in]Specifies name for every multimedia file and it
'should be difference  e.g.:
'you want to play two multimedia files the first maybe
'named "audio1" then you should name the other difference.

'filename
'[in]Specifies file name and the path it can contain any space
'which you want to play.

'typeDevice
'[in] Specifies a type of MCI device and it could be from the following:
'Type MCI       description                     driver file
'sequencer      dealing with mid                mciseq.drv
'               files
'MPEGVideo      dealing with most multimedia    mciqtz.drv
'               like mpg,mp3,mp2..
'               au,aiff,..etc also support
'               avi,vob(for DVD),midi,mid
'               and rmi files.because of this
'               my advice to you to use
'               type "MPEGVideo" to playing
'               MOST FILES even avi!!
'               I got this info from my
'               experiment when I opened
'               System.ini in section MCI
'               Then I must share others.
'avivideo       deling with avi movie           mciavi.drv

'the following types if you had ATI RAGE II or Later
'(This VGA Card to Support DVD Video)

'DvdVideo       This support DVD's Video        MciCinem.drv DVD
'ATIMPEGVIDEO   to playing MPEG Video           mciatim1.drv

'But my advice to you to not use type "ATIMPEGVIDEO" & "DvdVideo" because
'Type MPEGVideo can support most Multimedia files and also support DVD's
'Video if you had ATI RAGE II or LATER.
'last note for DVD Video: you must have a fast computer

'note : Type "MpegVideo" support these extensions:
'qt , mov, dat,snd, mpg, mpa, mpv, enc, m1v, mp2,mp3, mpe, mpeg, mpm
'au , snd, aif, aiff, aifc,wav,wmv,wma,avi,midi,mid,rmi,avi,etc.

'Note if there are any new type in (system.ini in windows 98 or in registry in windows 2000)
'it will supported by Type "MPEGVideo" because of this use type "MPEGVideo" to playing
'Most Files and remember you can use sequencer for mid and avivideo for avi,,etc.

'Now you must note using Type "MPEGVideo" can playing all Multimedia files

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

'Okay make sure if you used this function don't forget to use function
'CloseMultimedia or CloseAll When you will end your program or you
'will got error message

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFile As String
Const WS_CHILD = &H40000000

lenShort = GetShortPathName(FileName, tmp, 255)
ShortPathAndFile = Left$(tmp, lenShort) 'cut short path from buffer


cmdToDo = "open " & ShortPathAndFile & " type " & typeDevice & " Alias " & AliasName & " parent " & hWnd & " Style " & WS_CHILD
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    OpenMultimedia = ret: Exit Function
End If

'Success
OpenMultimedia = "Success"
End Function

Public Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String
'Calling PlayMultimedia will playing the multimedia file.
'Parameters

'AliasName
'[in]Specifies name alias name which you want play it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'from_where
'[in] Specifies the first frame in playing

'to_where
'[in]Specifies the last frame in playing

'if from_where is vbNullString and the to_where is vbNullString the Function will:
'playing from the beginning to end.

'if from_where is 10 and to_where is 100 the Function will:
'playing from 10 to 100 and stop.

'if from_where is vbNullString and to_where is 100 the Function will:
'playing from the beginning to 100 and stop.

'if from_where is 104 and to_where is vbNullString the Function will:
'playing from 104 to end.

'Note :the numbers 10,100,104 is an example for from where start playing to where end playing

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)

'Improtant for auto repeat
If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128

cmdToDo = "play " & AliasName & " from " & from_where & " to " & to_where

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&) 'play

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'get the error
    PlayMultimedia = ret
    Exit Function
End If

'Success
PlayMultimedia = "Success"
End Function

Public Function CloseMultimedia(AliasName As String) As String
'Calling CloseMultimedia will close the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want Close it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'you must call this function if you called OpenMultimedia
'And want to close your program or you will get an error message

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close " & AliasName, 0&, 0&, 0&) 'close

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    CloseMultimedia = ret
    Exit Function
End If

'Success
If AliasName = glo_AliasName Then 'if alias the same
'this mean the user close this alias then we must delete
'the timer Function
KillTimer glo_hWnd, 500
End If

CloseMultimedia = "Success"
End Function

Public Function PauseMultimedia(AliasName As String) As String
'Calling PauseMultimedia will pause the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want Pause it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause " & AliasName, 0&, 0&, 0&) 'pause

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    PauseMultimedia = ret
    Exit Function
End If

'Success
PauseMultimedia = "Success"
End Function

Public Function StopMultimedia(AliasName As String) As String
'Calling StopMultimedia will Stop the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want Stop it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop " & AliasName, 0&, 0&, 0&) 'stop

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    StopMultimedia = ret
    Exit Function
End If

'Success
StopMultimedia = "Success"
End Function

Public Function ResumeMultimedia(AliasName As String) As String
'Calling ResumeMultimedia will Resume the multimedia file
'note: if you paused or stopped the file call this function to Continue
'( don't call PlayMultimedia function to Continue)

'Parameters

'AliasName
'[in]Specifies name alias name which you want Resume it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume " & AliasName, 0&, 0&, 0&) 'Resume

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    ResumeMultimedia = ret
    Exit Function
End If

'Success
ResumeMultimedia = "Success"
End Function

Public Function GetStatusMultimedia(AliasName As String) As String
'Calling Function GetStatusMultimedia will tell if the multimedia file
'now is playing or stopped or paused

'Parameters

'AliasName
'[in]Specifies name alias name which you want Get status for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value string
'(the status of multimedia file) if it "playing" or "paused" or "stopped"
'or if not will return value string "ERROR"


'also you can exame the status like this: you can copy it
'Dim Result As String
'Result = GetStatusMultimedia("aliasname")'alias name for e.g. movie
'If Result = "ERROR" Then 'this mean failed then write your commands here
''.....
''....
''..
'ElseIf Result = "playing" Then 'this mean it now playing .ok write your commands here
''....
''...
''..
'ElseIf Result = "stopped" Then 'this mean it now stopped .ok write your commands here
''....
''...
''..
'ElseIf Result = "paused" Then 'this mean it now paused .ok write your commands here
''....
''...
''..

'End If


Dim dwReturn As Long
Dim status As String * 128
Dim ret As String * 128

dwReturn = mciSendString("status " & AliasName & " mode", status, 128, 0&)  'Get status

If Not dwReturn = 0 Then  'not success
    GetStatusMultimedia = "ERROR"
    Exit Function
End If

'Extract just the string
Dim i As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(status, 1)
For i = 1 To Len(status)
    CharA = Mid(status, i, 1)
    If CharA = RChar Then Exit For
    GetStatusMultimedia = GetStatusMultimedia + CharA
Next i
End Function

Public Function GetTotalframes(AliasName As String) As Long
'Calling GetTotalframes will Get the Total frames for
'the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want Get Total frames for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value long
'is "number of total frames"
'or if not will return value long is -1

Dim dwReturn As Long
Dim Total As String * 128

dwReturn = mciSendString("set " & AliasName & " time format frames", Total, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", Total, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetTotalframes = -1
    Exit Function
End If

'Success
GetTotalframes = Val(Total)
End Function

Public Function GetTotalTimeByMS(AliasName As String) As Long
'Calling GetTotalTimeByMS will Get the Total time by
'millisecond for the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want Get Total time for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Note : if this Function success will return value long
'is "the Total time by millisecond" divid by 1000 if you want the time by second
'or if not will return value long is -1

Dim dwReturn As Long
Dim TotalTime As String * 128


dwReturn = mciSendString("set " & AliasName & " time format ms", TotalTime, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", TotalTime, 128, 0&)

mciSendString "set " & AliasName & " time format frames", 0&, 0&, 0& ' return focus to frames not to time

If Not dwReturn = 0 Then  'not success
    GetTotalTimeByMS = -1
    Exit Function
End If

'Success
GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMultimedia(AliasName As String, to_where As Long) As String
'Calling MoveMultimedia will seek (change the position)for
'the multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want change position for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'to_where
'[in]Specifies number frame which you want jump to it

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128

dwReturn = mciSendString("seek " & AliasName & " to " & to_where, 0&, 0&, 0&)
mciSendString "Play " & AliasName, 0&, 0&, 0&

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    MoveMultimedia = ret
    Exit Function
End If

'Success
MoveMultimedia = "Success"
End Function

Public Function GetCurrentMultimediaPos(AliasName As String) As Long
'Calling Function GetCurrentMultimediaPos will get the current frame

'Parameters

'AliasName
'[in]Specifies name alias name which you want Get current frame for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'the returned value from this function is number of current frame
'and if the function failed will return value -1


Dim dwReturn As Long
Dim Pos As String * 128

dwReturn = mciSendString("status " & AliasName & " position", Pos, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetCurrentMultimediaPos = -1
    Exit Function
End If

'Success
GetCurrentMultimediaPos = Val(Pos)
End Function

Public Function PutMultimedia(hWnd As Long, AliasName As String, Left As Long, Top As Long, Width As Long, Height As Long) As String
'Calling PutMultimedia will resize the movie

'Parameters

'hWnd
'Specifies the handle of the window.
'note: don't think this handle to put movie on it, this handle to get the size from it.

'AliasName
'[in]Specifies name alias name which you want to resize the movie
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Left
'Specifies the new position of the left side of the window.

'Top
'Specifies the new position of the top of the window.

'Width
'Specifies the new width of the window.

'Height
'Specifies the new height of the window.


'if you are set parameter width or Height zero
'the function will get the actual size of the window which
'want to play in and resize the movie to fit the window(hWnd)

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128

If Width = 0 Or Height = 0 Then
    'Get Window Size
    Dim rec As RECT
    Call GetWindowRect(hWnd, rec)
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.Top
End If

dwReturn = mciSendString("put " & AliasName & " window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    PutMultimedia = ret
    Exit Function
End If

'Success
PutMultimedia = "Success"
End Function
Public Function GetPercent(AliasName As String) As Long
'Calling Function GetPercent will get the percent of plying file

'Parameters

'AliasName
'[in]Specifies name alias name which you want to Get percent for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'the returned value from this function is Percent "Progress"
'if it successed and if the function failed will return value -1

On Error Resume Next
Dim TotalFrames As Long
Dim currframe As Long
TotalFrames = GetTotalframes(AliasName)
currframe = GetCurrentMultimediaPos(AliasName)

If TotalFrames = -1 Or currframe = -1 Then 'Not success
    GetPercent = -1
    Exit Function
End If

'Success
GetPercent = currframe * 100 / TotalFrames
End Function
Public Function GetFramesPerSecond(AliasName As String) As Long
'Calling Function GetFramesPerSecond will get amount frames per second

'Parameters

'AliasName
'[in]Specifies name alias name which you want to Get number frames
'per second for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success


'this Function Will return amount frames per second if it
'Success or if not will return value -1

Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTimeByMS(AliasName)
TotalFrames = GetTotalframes(AliasName)
If TotalFrames = -1 Or TotalTime = -1 Then 'Not success
    GetFramesPerSecond = -1
    Exit Function
End If

'Success
GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function
Public Function GetSize(AliasName As String, CxOrCy As String) As Long
'Calling GetSize will get current width(cx) or height(cy)

'Parameters

'AliasName
'[in]Specifies name alias name which you want to get the current size for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'cxOrcy
'Specifies the width or height and you must note if you want to get the current width
'set this pararmeter ="cx"
'and if you want to get the current height set this parameter = "cy"

'Important Note:
'if you want to get the actual size you (must) call this function after Calling
'Function OpenMultimedia (directly)before resize the movie.
'and note if you resized the movie and after that called this function then you will
'get the current size.


'Note : if this Function success will return value long (width  or height )
'or if not will return value long is -1


If Not LCase(CxOrCy) = "cx" And Not LCase(CxOrCy) = "cy" Then GetSize = -1: Exit Function
Dim dwReturn As Long
Dim Size As String * 128
Dim s1, s2, s3, Width, Height As Long

dwReturn = mciSendString("Where " & AliasName & " destination", Size, 128, 0&)


If Not dwReturn = 0 Then  'not success
    GetSize = -1
    Exit Function
End If

s1 = InStr(1, Size, " "): s2 = InStr(s1 + 1, Size, " "): s1 = InStr(s2 + 1, Size, " ")
Width = Mid(Size, s2, s1 - s2): Height = Mid(Size, s1 + 1)

'Success
If LCase(CxOrCy) = "cx" Then 'get the width
GetSize = Width
ElseIf LCase(CxOrCy) = "cy" Then 'Get the height
GetSize = Height
End If

End Function
Public Function CloseAll() As String
'This Fucntion will close all multimedia files.
'use it when you want to end your program

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close All", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    CloseAll = ret
    Exit Function
End If

'Success
CloseAll = "Success"
End Function
Public Function ChannelsControl(AliasName As String, Channel As String, OnOrOFF As String) As String
'Callig ChannelsControl will make controls for channels audio (left and right)

'Parameters

'AliasName
'[in]Specifies name alias name which you want to make controls for channels audio
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'channel
'[in]Specifies name for channel which you want to make control for it
'this parameter must be from the following:
'channel             Description
'"left"              to make control for left audio channel
'"right"             to make control for right audio channel
'"all"               to make control for both audio channels (left & right)

'OnOrOFF
'[in] Specifies the channel control. This parameter must be from the following:
'Type Control           Description
'"on"                   to turn the channel on
'"off"                  to turn the channel off

'Important Note:
'To make control for every channel work effectly like turn off channel and turn on
'the another channel BE sure the audio or movie file has two channels(Stereo)

'Note: Be sure if you played a Stereo file (has two channels)and you turned off one
'of the channels, the sound which in this channel will not appear,JUST will appear the sound
'which in the other channel
'for Example:
'you played a mp3 file and you listened the person in the left channel say "Oh yeah"
'and you listened the person on the right channel say "Okay" then :
'if you turned off the right channel you JUST hear "oh yeah"
'if you turned off the left channel you JUST hear "Okay"

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur



Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128

cmdToDo = "set " & AliasName & " audio " & Channel & " " & OnOrOFF

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    ChannelsControl = ret
    Exit Function
End If

'Success
ChannelsControl = "Success"

End Function

Public Function SetVolume(AliasName As String, Channel As String, VolumeValue As Long) As String
'Callig SetVolume will make control for volume channels

'Parameters

'AliasName
'[in]Specifies name alias name which you want to make control for volume channels audio
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Channel
'[in]Specifies name for channel which you want to make volume control for it
'this parameter must be from the following:
'channel                Description
'"left"                 to make control for volume left audio channel
'"right"                to make control for volume right audio channel
'any value like "all"   to make control for volume both audio channels (left & right)

'VolumeValue
'[in]Specifies value for Volume and this parameter must be from 0 to 100

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
Dim VolumeV As Long
VolumeV = VolumeValue

If VolumeV < 0 Or VolumeV > 100 Then
    SetVolume = "out of volume"
    Exit Function
End If

VolumeV = VolumeV * 10

If LCase(Channel) = "left" Or LCase(Channel) = "right" Then
    cmdToDo = "setaudio " & AliasName & " " & Channel & " Volume to " & VolumeV
Else
    cmdToDo = "setaudio " & AliasName & " Volume to " & VolumeV
End If

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    SetVolume = ret
    Exit Function
End If

'Success
SetVolume = "Success"
End Function


Public Function GetVolume(AliasName As String, Channel As String) As Long
'Callig GetVolume will get the volume for Specified channels (left or right) or both channels

'Parameters

'AliasName
'[in]Specifies name alias name which you want to get volume for channels audio
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Channel
'[in]Specifies name for channel which you want to get volume for it
'this parameter must be from the following:
'channel                Description
'"left"                 to get volume left audio channel
'"right"                to get volume right audio channel
'any value like "all"   to get volume both audio channels (left & right)

'Note : if this Function success will return value long
'is "volume for specified channel"
'or if not will return value long is -1

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim Volume As String * 128

If LCase(Channel) = "left" Or LCase(Channel) = "right" Then
    cmdToDo = "status " & AliasName & " " & Channel & " Volume"
Else
    cmdToDo = "status " & AliasName & " Volume"
End If

dwReturn = mciSendString(cmdToDo, Volume, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetVolume = -1
    Exit Function
End If

'Success
GetVolume = Val(Volume) / 10
End Function

Public Function SetRate(AliasName As String, RateValue As Long) As String
'Callig SetRate will increase or decrease speed playing for Multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want to increase or decrease speed for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'Rate
'[in]Specifies value for speed playing Multimedia file, this parameter must be from 0 to 200
'the following:
'Rate                   description
'100                    playing Multimedia file as normal speed
'more than 100          will increase speed playing file
'less than 100          will decrease speed playing file

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur


Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
Dim RateV As Long

RateV = RateV
If RateV < 0 Or RateV > 200 Then
   SetRate = "out of rate"
   Exit Function
End If


RateV = RateValue * 10


cmdToDo = "set " & AliasName & " speed " & RateV

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128 'Get the error
    SetRate = ret
    Exit Function
End If

'Success
SetRate = "Success"
End Function

Public Function GetRate(AliasName As String) As Long
'Callig GetRate will get current rate for Multimedia file

'Parameters

'AliasName
'[in]Specifies name alias name which you want to get current rate for it
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success


'Note : if this Function success will return value long
'is "the current rate for Multimedia file"
'or if not will return value long is -1

Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim Rate As String * 128

cmdToDo = "status " & AliasName & " speed"

dwReturn = mciSendString(cmdToDo, Rate, 128, 0&)

If Not dwReturn = 0 Then  'not success
    GetRate = -1
    Exit Function
End If

'Success
GetRate = Val(Rate) / 10
End Function


Public Function AreMultimediaAtEnd(AliasName As String, lastFrame As Long) As Boolean
'Calling Function AreMultimediaAtEnd will let you know if the File at
'the end now and this benefit you if you want to plays a list of files or make auto repeat
'(play the file again}

'Parameters

'AliasName
'[in]Specifies name alias name which you want to know if it at the end now
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'lastFrame
'[in]Specifies the last frame you want to play to
'if this parameter is zero (0) this function will get the last frame

'This Function will tell if multimedia file now at end
'To use this Function do the following:
'1-put it in a timer and set Interval for a timer = 100
'2-make the timer false
'3-after Play Multimedia files Successfully set the timer true.
'4-The Commands Which you will put it in a timer the Following:

'Copy the Following in a timer

'If AreMultimediaAtEnd("aliasname") = True Then' alias name for e.g.:"movie"
    ''this mean  file multimedia at the end now then
    ''write your commnads here or call you favourit Fucntion
    ''or even you can play the file again or play the next file
    ''if you had a list of multimedia files.
    '.....
    '...
    '..
    'if you want to know if the multimedia file
    'at the end now don't use option Auto Repeat
    'you must do auto repeat by yourself by the following commands:
    
    'Result = PlayMultimedia("aliasname",txtFrom, TxtTo)

    ''or you have choice to close this File and open
    ''another file and play it( this if had a list of files):
    
    'Dim Result As String
    'Result = CloseMultimedia("aliasname")
    'Result = OpenMultimedia(FrameVideo.hwnd,"aliasname", filename, typeDevice) 'call now function openMultimedia
    'Result = PlayMultimedia("aliasname",txtFrom, TxtTo)

    
'Else
    'this mean result calling function false and this mean the
    'multimedia file not at the end now
    '....
    '...
    '..

'End If


Dim currpos As Long

'if last frame is zero then get actaul last frame
If lastFrame = 0 Then lastFrame = GetTotalframes(AliasName)

currpos = Val(GetCurrentMultimediaPos(AliasName))

If currpos = -1 Or lastFrame = -1 Then 'there are an error then not resume
    AreMultimediaAtEnd = False
    Exit Function
End If
    
If lastFrame = currpos Or (lastFrame - 1) < currpos Then
AreMultimediaAtEnd = True ' ok we reach to last frame
Else
AreMultimediaAtEnd = False ' we not reach to last frame
End If
End Function
Public Function SetAutoRepeat(hWnd As Long, AliasName As String, first_frame As String, last_frame As String, autoTrueOrFalse As Boolean) As Boolean
'Calling this Function will set Specifies multimedia auto repeat or not

'Improtant:
'1-you can not use this function to set auto repeat for more one multimedia file.
'2-keep in your mind if you want to use this function call it after calling OpenMultimedia function not else.

'Parameters

'hWnd
'Specifies the handle of the window (this window we will create timer in).

'AliasName
'[in]Specifies name alias name which you want to Set auto repeat
'Note : you must let this parameter the alias which you
'used it OpenMultimedia Function or this function not Success

'firstFrame
'[in]Specifies the first frame you want to play  from
'if this parameter is vbNullString then the first frame be 0  .

'lastFrame
'[in]Specifies the last frame you want to play  to
'if this parameter is vbNullString then the last frame be the actual last frame

'autoTrueOrFalse
'Specifies if you want auto repeat or kill auto repeat.
'if this parameter true this mean you want to set auto repeat.
'if this parameter false this mean you want to kill auto repeat.

'if this Function success will return true or if not will return false.

Dim Result As String

If first_frame = vbNullString Then first_frame = 0
If last_frame = vbNullString Then last_frame = GetTotalframes(AliasName)

glo_from = first_frame 'store it in global to use it TimerFunction
glo_to = last_frame ' store it in global to use it TimerFunction

glo_hWnd = hWnd
If autoTrueOrFalse = True Then
    glo_AliasName = AliasName
    Result = SetTimer(hWnd, 500, 100, AddressOf TimerFunction)
Else
    glo_AliasName = vbNullString
    Result = KillTimer(hWnd, 500)
End If

If Result = 0 Then
    SetAutoRepeat = False
Else
    SetAutoRepeat = True
End If
End Function

Sub TimerFunction()
'Important for auto repeat
Dim currpos As Long
Dim Result As String
currpos = Val(GetCurrentMultimediaPos(glo_AliasName))
If currpos = -1 Then Exit Sub   'if  function get cuurent pos not success then exit
'
If Val(glo_to) = currpos Or (Val(glo_to) - 1) < currpos Then
    Result = PlayMultimedia(glo_AliasName, Str(glo_from), Str(glo_to))
    If Not Result = "Success" Then KillTimer glo_hWnd, 500 'if  function play not success then kill timer
End If
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
'this sub is very important to set the default MCI device
'maybe xing mpeg installed in your computer and it not support
'all multimedia files
'because of this you can rest the default device of MCI to
'drivers microsft
'which came with windows or you when install Microsft media player
'ok any way the default device Following:
'Device Type        Driver
'MPEGVideo          mciqtz.drv          this is the most important
'sequencer          mciseq.drv
'avivideo           mciavi.drv
'waveaudio          mciwave.drv
'videodisc          mcipionr.drv
'cdaudio            mcicda.drv

'the following for ATI all in Wonder 128 VGA card
'DvdVideo           MciCinem.drv DVD
'ATIMPEGVIDEO       mciatim1.drv

'e.g. :
'SetDefaultDevice "MPEGVideo", "mciqtz.drv" ' this the most
'improtant device and it will receives calls mci
'Some programs change this device like xing mpeg
'and if this occur you can not play all mutimedia files
'and will occur unexpected errors
'because of this write this line when your program loaded
'SetDefaultDevice "MPEGVideo", "mciqtz.drv"
'to set the strongest default device

'Note: Windows 2000 not use system.ini to set drivers.it use registry.

Dim Res As String
Dim tmp As String * 255
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
'this Function help you if you want to know the default device
'the parameter must be the device type like:
'MPEGVideo
'sequencer
'avivideo
'waveaudio
'videodisc
'cdaudio
'and the returned value is a string for the default device
'Please read the description of sub SetDefaultDevice

Dim tmp As String * 255
Dim Res As String
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
GetDefaultDevice = Left$(tmp, Res)
End Function

'Okay I hope you Enjoyed
'You can use this module in your own projects if you wanna
'the easist deal with multimedia.
'Using API is more stronger than using controls and not take a space
'for any request, suggestions,Devlopment or bugs   e-mail at
'a_ahdal@yahoo.com
'Thank you
