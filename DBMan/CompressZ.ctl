VERSION 5.00
Begin VB.UserControl CompressZIt 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   540
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "CompressZ.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "CompressZIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_CompressedSize = 0
Const m_def_OriginalSize = 0
'Property Variables:
Dim m_CompressedSize As Long
Dim m_OriginalSize As Long

'Declares
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Enum CZErrors
[Insufficient Buffer] = -5
End Enum

Public Function About() As Boolean
Attribute About.VB_UserMemId = -552

    i = MsgBox("Compress-Z-It" & Chr$(10) & Chr$(10) & "Data compression ActiveX component module." & Chr$(10) & Chr$(10) & "Custom control written and compiled by Benjamin Dowse. Portions written by other external 'zLib' compression software library authors." & Chr$(10) & Chr$(10) & "Special thanks and honor to the authors of the zLib DLL.", vbInformation + vbOKOnly, "DowseWare - Compress-Z-It ActiveX Control")

End Function

Public Function CompressData(TheData() As Byte) As Long
Attribute CompressData.VB_Description = "Compress a byte array or raw binary data."

OriginalSize = UBound(TheData) + 1

'Allocate memory for byte array
Dim BufferSize As Long
Dim TempBuffer() As Byte

BufferSize = UBound(TheData) + 1
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'Compress byte array (data)
Result = compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

'Truncate to compressed size
ReDim Preserve TheData(BufferSize - 1)
CopyMemory TheData(0), TempBuffer(0), BufferSize

'Cleanup
Erase TempBuffer

'Set properties if no error occurred
If Result = 0 Then CompressedSize = UBound(TheData) + 1

'Return error code (if any)
CompressData = Result

End Function

Public Function CompressString(TheString As String) As Long
Attribute CompressString.VB_Description = "Compresses a string. String may contain Null characters."

OriginalSize = Len(TheString)

'Allocate string space for the buffers
Dim CmpSize As Long
Dim TBuff As String
orgSize = Len(TheString)
TBuff = String(orgSize + (orgSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)

'Compress string (temporary string buffer) data
ret = compress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))

'Set original value
OriginalSize = Len(TheString)

'Crop the string and set it to the actual string.
TheString = left$(TBuff, CmpSize)

'Set compressed size of string.
CompressedSize = CmpSize

'Cleanup
TBuff = ""

'Return error code (if any)
CompressString = ret

End Function

Public Function DecompressData(TheData() As Byte, OrigSize As Long) As Long
Attribute DecompressData.VB_Description = "Decompresses a compressed byte array or raw binary data."

'Allocate memory for buffers
Dim BufferSize As Long
Dim TempBuffer() As Byte

BufferSize = OrigSize
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'Decompress data
Result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

'Truncate buffer to compressed size
ReDim Preserve TheData(BufferSize - 1)
CopyMemory TheData(0), TempBuffer(0), BufferSize

'Reset properties
If Result = 0 Then
CompressedSize = 0
OriginalSize = 0
End If

'Return error code (if any)
DecompressData = Result

End Function

Public Function DecompressString(TheString As String, OrigSize As Long) As Long
Attribute DecompressString.VB_Description = "Decompresses a compressed string. String may contain Null characters."

'Allocate string space
Dim CmpSize As Long
Dim TBuff As String
TBuff = String(OriginalSize + (OriginalSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)

'Decompress
Result = uncompress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))

'Make string the size of the uncompressed string
TheString = left$(TBuff, CmpSize)

'Reset properties
If Result = 0 Then
CompressedSize = 0
OriginalSize = 0
End If

'Return error code (if any)
DecompressString = ret

End Function

Public Property Get CompressedSize() As Long
Attribute CompressedSize.VB_Description = "Determine compressed size of last compressed data or string."
Attribute CompressedSize.VB_MemberFlags = "400"
    CompressedSize = m_CompressedSize
End Property

Public Property Let CompressedSize(ByVal New_CompressedSize As Long)
    If Ambient.UserMode = False Then Exit Property
    m_CompressedSize = New_CompressedSize
    PropertyChanged "CompressedSize"
End Property

Public Property Get OriginalSize() As Long
Attribute OriginalSize.VB_Description = "Determines the original size of the last compressed data or string."
Attribute OriginalSize.VB_MemberFlags = "400"
    OriginalSize = m_OriginalSize
End Property

Public Property Let OriginalSize(ByVal New_OriginalSize As Long)
    If Ambient.UserMode = False Then Exit Property
    m_OriginalSize = New_OriginalSize
    PropertyChanged "OriginalSize"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CompressedSize = m_def_CompressedSize
    m_OriginalSize = m_def_OriginalSize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
End Sub

