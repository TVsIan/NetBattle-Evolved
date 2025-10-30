VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerialGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
'IDE registers
Private Type IDEREGS
   bFeaturesReg     As Byte 'Used for specifying SMART "commands"
   bSectorCountReg  As Byte 'IDE sector count register
   bSectorNumberReg As Byte 'IDE sector number register
   bCylLowReg       As Byte 'IDE low order cylinder value
   bCylHighReg      As Byte 'IDE high order cylinder value
   bDriveHeadReg    As Byte 'IDE drive/head register
   bCommandReg      As Byte 'Actual IDE command
   bReserved        As Byte 'reserved for future use - must be zero
End Type
'Status returned from driver
Private Type DRIVERSTATUS
   bDriverError  As Byte          'Error code from driver, or 0 if no error
   bIDEStatus    As Byte          'Contents of IDE Error register
                                  'Only valid when bDriverError is SMART_IDE_ERROR
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type
'SENDCMDINPARAMS contains the input parameters for the
'Send Command to Drive function
Private Type SENDCMDINPARAMS
   cBufferSize     As Long     'Buffer size in bytes
   irDriveRegs     As IDEREGS  'Structure with drive register values.
   bDriveNumber    As Byte     'Physical drive number to send command to (0,1,2,3).
   bReserved(2)    As Byte     'Bytes reserved
   dwReserved(3)   As Long     'DWORDS reserved
   bBuffer()       As Byte      'Input buffer.
End Type
Private Type SENDCMDOUTPARAMS
  cBufferSize       As Long         'Size of Buffer in bytes
  DRIVERSTATUS      As DRIVERSTATUS 'Driver status structure
  bBuffer(512)      As Byte          'Buffer of arbitrary length for data read from drive
End Type
'Valid values for the bCommandReg memb
Private Type GETVERSIONOUTPARAMS
   bVersion       As Byte 'Binary driver version.
   bRevision      As Byte 'Binary driver revision
   bReserved      As Byte 'Not used
   bIDEDeviceMap  As Byte 'Bit map of IDE devices
   fCapabilities  As Long 'Bit mask of driver capabilities
   dwReserved(3)  As Long 'For future use
End Type
Private Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type
Private Type STORAGE_DEVICE_DESCRIPTOR
  Version               As Long
  Size                  As Long
  DeviceType            As Byte
  DeviceTypeModifier    As Byte
  RemovableMedia        As Byte
  CommandQueueing       As Byte
  VendorIdOffset        As Long
  ProductIdOffset       As Long
  ProductRevisionOffset As Long
  SerialNumberOffset    As Long
  BusType               As STORAGE_BUS_TYPE
  RawPropertiesLength   As Long
  'RawDeviceProperties() As Byte
End Type
Private Type StoragePropertyQuery
    PropertyID              As StoragePropertyID
    QueryType               As StorageQueryType
    AdditionalParameters()  As Byte
End Type
Private Type SRB_IO_CONTROL
    HeaderLength            As Long
    Signature(7)            As Byte
    Timeout                 As Long
    ControlCode             As Long
    ReturnCode              As Long
    Length                  As Long
End Type
Private Type RT_IdeDInfo
    IDEExists(3)            As Byte
    DiskExists(7)           As Byte
    DisksRawInfo(2047)      As Integer
End Type
Private Type RT_DiskInfo
   DiskExists           As Byte
   ATAdevice            As Byte
   RemovableDevice      As Byte
   TotLogCyl            As Integer
   TotLogHeads          As Integer
   TotLogSPT            As Integer
   SerialNumber(19)     As Byte
   FirmwareRevision(7)  As Byte
   ModelNumber(39)      As Byte
   CurLogCyl            As Integer
   CurLogHeads          As Integer
   CurLogSPT            As Integer
End Type


Private Enum ePlatform
   VER_PLATFORM_WIN32_WINDOWS = 1
   VER_PLATFORM_WIN32_NT = 2
End Enum
Private Enum DESIREDFLAGS
  GENERIC_READ = &H80000000
  GENERIC_WRITE = &H40000000
End Enum
Private Enum SHAREFLAGS
  FILE_SHARE_READ = &H1
  FILE_SHARE_WRITE = &H2
End Enum
Private Enum DISPOSITIONFLAGS
  OPEN_EXISTING = 3
End Enum
Private Enum IOCOMMANDS
  IOCTL_STORAGE_CHECK_VERIFY = &H2D4800
  IOCTL_STORAGE_CHECK_VERIFY2 = &H2D0800
  IOCTL_STORAGE_MEDIA_REMOVAL = &H2D4804
  IOCTL_STORAGE_EJECT_MEDIA = &H2D4808
  IOCTL_STORAGE_LOAD_MEDIA = &H2D480C
  IOCTL_STORAGE_LOAD_MEDIA2 = &H2D080C
  IOCTL_STORAGE_RESERVE = &H2D4810
  IOCTL_STORAGE_RELEASE = &H2D4814
  IOCTL_STORAGE_FIND_NEW_DEVICES = &H2D4818
  IOCTL_STORAGE_EJECTION_CONTROL = &H2D0940
  IOCTL_STORAGE_MCN_CONTROL = &H2D0944
  IOCTL_STORAGE_GET_MEDIA_TYPES = &H2D0C00
  IOCTL_STORAGE_GET_MEDIA_TYPES_EX = &H2D0C04
  IOCTL_STORAGE_GET_MEDIA_SERIAL_NUMBER = &H2D0C10
  IOCTL_STORAGE_GET_HOTPLUG_INFO = &H2D0C14
  IOCTL_STORAGE_SET_HOTPLUG_INFO = &H2DCC18
  IOCTL_STORAGE_RESET_BUS = &H2D5000
  IOCTL_STORAGE_RESET_DEVICE = &H2D5004
  IOCTL_STORAGE_BREAK_RESERVATION = &H2D5014
  IOCTL_STORAGE_GET_DEVICE_NUMBER = &H2D1080
  IOCTL_STORAGE_PREDICT_FAILURE = &H2D1100
  IOCTL_STORAGE_QUERY_PROPERTY = &H2D1400
  IOCTL_SCSI_PASS_THROUGH = &H4D004
  IOCTL_SCSI_MINIPORT = &H4D008
  IOCTL_SCSI_GET_INQUIRY_DATA = &H4100C
  IOCTL_SCSI_GET_CAPABILITIES = &H41010
  IOCTL_SCSI_PASS_THROUGH_DIRECT = &H4D014
  IOCTL_SCSI_GET_ADDRESS = &H41018
  IOCTL_SCSI_RESCAN_BUS = &H4101C
  IOCTL_SCSI_GET_DUMP_POINTERS = &H41020
  IOCTL_SCSI_FREE_DUMP_POINTERS = &H41024
  IOCTL_IDE_PASS_THROUGH = &H4D028
End Enum
Enum StoragePropertyID
    StorageDeviceProperty = 0
    StorageAdapterProperty
End Enum
Enum StorageQueryType
    PropertyStandardQuery = 0          ' Retrieves the descriptor
    PropertyExistsQuery                ' Used to test whether the descriptor is supported
    PropertyMaskQuery                  ' Used to retrieve a mask of writeable fields in the descriptor
    PropertyQueryMaxDefined     ' use to validate the value
End Enum
Private Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum

Private Const INVALID_HANDLE_VALUE = -1
Private Const IDE_ATAPI_IDENTIFY = &HA1  ' Returns ID sector for ATAPI.
Private Const IDE_ATA_IDENTIFY = &HEC  ' Returns ID sector for ATA.
Private Const DFP_GET_VERSION = &H74080
Private Const DFP_SEND_DRIVE_COMMAND = &H7C084
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088
Private Const IDENTIFY_BUFFER_SIZE = 512
Private Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16
Private Const FILE_DEVICE_SCSI = &H1B
Private Const IOCTL_SCSI_MINIPORT_IDENTIFY = (FILE_DEVICE_SCSI * &H10000) + &H501
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    OSVI As OSVERSIONINFO) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" ( _
    ByVal lpFilename As String, _
    ByVal dwDesiredAccess As DESIREDFLAGS, _
    ByVal dwShareMode As SHAREFLAGS, _
    ByRef lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As DISPOSITIONFLAGS, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" ( _
    ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    ByRef lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    ByRef lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    ByRef lpBytesReturned As Long, _
    ByRef lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    ByRef Dest As Any, _
'    ByRef Source As Any, _
'    ByVal Count As Long)
Private Declare Function SetPriorityClass Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private SerialBuffer As String
Private OSType As ePlatform

Public VXDFile As String

Public Function GetSerial() As String
    Dim Temp As String
    On Error Resume Next
    If OSType = VER_PLATFORM_WIN32_NT Then
        WriteDebugLog "Starting XPA Read"
        Temp = ReadUsingXPAdminRights
        If Len(Temp) = 0 Then
            WriteDebugLog "Starting SCSI Read"
            Temp = ReadUsingSCSIBackdoor
        End If
        If Len(Temp) = 0 Then
            WriteDebugLog "Starting XPZ Read"
            Temp = ReadUsingXPZeroRights
        End If
        GetSerial = Temp
    Else
        WriteDebugLog "SIDG Starting Win9X Read"
        GetSerial = ReadUsingVxDinWin9X
    End If
End Function

Private Sub Class_Initialize()
    'First let's find out if we're running NT/XP or not.
    
    Dim OSV As OSVERSIONINFO
    OSV.dwOSVersionInfoSize = Len(OSV)
    GetVersionEx OSV
    OSType = OSV.dwPlatformId
    SerialBuffer = vbNullString
    VXDFile = SlashPath & "virtual.drv" '"\\.\IDE21201.VXD"
    WriteDebugLog "SIDG Init"
End Sub

Public Function ReadUsingXPAdminRights(Optional DeviceID As Long = 0) As String
    Dim DriveName As String
    Dim hDrive As Long
    Dim VersionParams As GETVERSIONOUTPARAMS
    Dim SCIP As SENDCMDINPARAMS
    Dim SCOP As SENDCMDOUTPARAMS
    Dim cbBytesReturned As Long
    Dim DevMap As Byte
    Dim bIDCmd As Byte
    Dim DiskData(255) As Long
    Dim bArrOut() As Byte
    Dim IDSec As IDSECTOR
    Dim X As Long
    Debug.Print "ReadXPA " & DeviceID
    DriveName = "\\.\PhysicalDrive" & CStr(DeviceID)
    hDrive = CreateFile(DriveName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    If hDrive <> INVALID_HANDLE_VALUE Then 'If error here, we've failed, try a different method
        DeviceIoControl hDrive, DFP_GET_VERSION, ByVal 0&, 0, VersionParams, Len(VersionParams), cbBytesReturned, ByVal 0&
        DevMap = VersionParams.bIDEDeviceMap
        If (ShiftRight(DevMap, DeviceID) And 16) <> 0 And DevMap > 0 Then
            bIDCmd = IDE_ATAPI_IDENTIFY
        Else
            bIDCmd = IDE_ATA_IDENTIFY
        End If
        If DoIdentify(hDrive, SCIP, bArrOut, bIDCmd, DeviceID, cbBytesReturned) Then
            CopyMemory IDSec, bArrOut(16), Len(IDSec)
            ReadUsingXPAdminRights = StrConv(SwapBytes(IDSec.sSerialNumber), vbUnicode)
            For X = 1 To Len(ReadUsingXPAdminRights)
                If Asc(Mid(ReadUsingXPAdminRights, X, 1)) = 0 Then
                    ReadUsingXPAdminRights = Left$(ReadUsingXPAdminRights, X - 1)
                    Exit For
                End If
            Next X
        End If
'        Else
'            bIDCmd = 0
'        End If
        CloseHandle hDrive
    Else
        'MsgBox "VxD: INVALID_HANDLE_VALUE"
    End If
End Function
Public Function ReadUsingXPZeroRights(Optional DeviceID As Long = 0) As String
    Dim DriveName As String
    Dim hDrive As Long
    Dim DevMap As Byte
    Dim Query As StoragePropertyQuery
    Dim cbBytesReturned As Long
    Dim Buffer(0 To 9999) As Byte
    Dim Serial() As Byte
    Dim SDD As STORAGE_DEVICE_DESCRIPTOR
    Dim X As Long
    Debug.Print "ReadXPZ " & DeviceID
    DriveName = "\\.\PhysicalDrive" & CStr(DeviceID)
    hDrive = CreateFile(DriveName, 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
    If hDrive <> INVALID_HANDLE_VALUE Then 'If error here, we've failed, try a different method
        Query.PropertyID = StorageDeviceProperty
        Query.QueryType = PropertyStandardQuery
        If DeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, Query, Len(Query), Buffer(0), 10000, cbBytesReturned, 0) Then
            CopyMemory ByVal VarPtr(SDD), ByVal VarPtr(Buffer(0)), Len(SDD)
            If SDD.SerialNumberOffset > -1 Then
                For X = SDD.SerialNumberOffset To cbBytesReturned
                'For X = 0 To cbBytesReturned
                    'Debug.Print Buffer(X) & " - " & Chr$(Buffer(X))
                    If Buffer(X) = 0 Then Exit For
                Next X
                If X > SDD.SerialNumberOffset + 1 Then
                    ReDim Serial(0 To X - SDD.SerialNumberOffset - 1)
                    For X = 0 To UBound(Serial)
                        Serial(X) = Buffer(X + SDD.SerialNumberOffset)
                    Next X
                    ReadUsingXPZeroRights = StrConv(FlipAndCodeBytes(Serial), vbUnicode)
                Else
                    'WriteDebugLog "X = SDD.SNO"
                End If
            Else
                'MsgBox "XPZR: SDD.SerialNumberOffset = -1"
            End If
        Else
            'MsgBox "XPZR: DeviceIoControl False"
        End If
        CloseHandle hDrive
    Else
        'MsgBox "XPZR: INVALID_HANDLE_VALUE"
    End If
End Function
Public Function ReadUsingSCSIBackdoor() As String

    Dim DriveName As String
    Dim hDrive As Long
    Dim Controller As Long
    Dim DriveNum As Long
    Dim Buffer() As Byte
    Dim p As SRB_IO_CONTROL
    Dim SCIP As SENDCMDINPARAMS
    Dim SCOP As SENDCMDOUTPARAMS
    Dim Dummy As Long
    Dim IDSec As IDSECTOR
    Dim Temp As String
    Dim X As Long
    Debug.Print "ReadSCSI"
    ReDim Buffer(Len(p) + Len(SCOP) + IDENTIFY_BUFFER_SIZE - 1)
    
    For Controller = 0 To 17
        DriveName = "\\.\Scsi" & CStr(Controller) & ":"
        hDrive = CreateFile(DriveName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
        If hDrive <> INVALID_HANDLE_VALUE Then
            For DriveNum = 0 To 1
                With p
                    .HeaderLength = Len(p)
                    .Timeout = 10000
                    .Length = Len(SCOP) + IDENTIFY_BUFFER_SIZE
                    .ControlCode = IOCTL_SCSI_MINIPORT_IDENTIFY
                    CopyMemory .Signature(0), ByVal "SCSIDISK", 8
                End With
                With SCIP
                    .irDriveRegs.bCommandReg = IDE_ATA_IDENTIFY
                    .bDriveNumber = DriveNum
                End With
                CopyMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(p), Len(p)
                CopyMemory ByVal VarPtr(Buffer(Len(p))), ByVal VarPtr(SCIP), Len(SCIP)
                If DeviceIoControl(hDrive, IOCTL_SCSI_MINIPORT, Buffer(0), Len(p) + Len(SCIP) - 1, Buffer(0), Len(p) + Len(SCOP) + IDENTIFY_BUFFER_SIZE, Dummy, ByVal 0&) Then
                    CopyMemory ByVal VarPtr(SCOP), ByVal VarPtr(Buffer(Len(p))), Len(SCOP)
                    If SCOP.cBufferSize > 0 Then
                        CopyMemory ByVal VarPtr(IDSec), ByVal VarPtr(SCOP.bBuffer(0)), Len(IDSec)
                        Temp = StrConv(SwapBytes(IDSec.sSerialNumber), vbUnicode)
                        For X = 1 To Len(Temp)
                            If Asc(Mid(Temp, X, 1)) = 0 Then
                                Temp = Left$(Temp, X - 1)
                                Exit For
                            End If
                        Next X
                        If Len(Temp) <> 0 Then
                            ReadUsingSCSIBackdoor = Temp
                            CloseHandle hDrive
                            Exit Function
                        End If
                    End If
                End If
            Next DriveNum
            CloseHandle hDrive
        End If
    Next Controller
End Function
Public Function ReadUsingVxDinWin9X() As String
    Dim VxDHandle As Long
    Dim OutBufVxD As RT_IdeDInfo
    Dim BytesReturned As Long
    Dim i As Long
    Dim J As Long
    Dim DiskInfo(255) As Integer
    Dim bDiskInfo() As Byte
    Dim Temp As String
    Dim X As Long
    Dim Y As Long
    Dim Attempt As Long
    ReDim bDiskInfo(0 To 511)
    Debug.Print "ReadWin9x"
    SetPriorityClass GetCurrentProcess, REALTIME_PRIORITY_CLASS
    For Attempt = 1 To 10
        
        VxDHandle = CreateFile(VXDFile, 0, 0, 0, 0, FILE_FLAG_DELETE_ON_CLOSE, 0)
        If VxDHandle <> INVALID_HANDLE_VALUE Then
            DeviceIoControl VxDHandle, 1, 0, 0, OutBufVxD, Len(OutBufVxD), BytesReturned, 0
            CloseHandle VxDHandle
            For i = 0 To 7
                If OutBufVxD.DiskExists(i) And OutBufVxD.IDEExists(i \ 2) Then
                    For J = 0 To 255
                        DiskInfo(J) = OutBufVxD.DisksRawInfo(i * 256 + J)
                        'Form1.Text1.Text = Form1.Text1.Text & Hex(DiskInfo(j)) & " "
                    Next J
                    CopyMemory ByVal VarPtr(bDiskInfo(0)), ByVal VarPtr(DiskInfo(10)), 492
                    bDiskInfo = SwapBytes(bDiskInfo)
                    Temp = StrConv(bDiskInfo, vbUnicode)
                    For X = 1 To Len(Temp)
                        If Asc(Mid(Temp, X, 1)) = 0 Then
                            Temp = Left$(Temp, X - 1)
                            Exit For
                        End If
                    Next X
                    Temp = Trim$(Temp)
                    If Len(Temp) <> 0 Then
                        ReadUsingVxDinWin9X = Temp
                        SetPriorityClass GetCurrentProcess, NORMAL_PRIORITY_CLASS
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next Attempt
    SetPriorityClass GetCurrentProcess, NORMAL_PRIORITY_CLASS
End Function

Private Function DoIdentify(hPhysicalDriveIOCTL As Long, SCIP As SENDCMDINPARAMS, bOut() As Byte, ByVal bIDCmd As Byte, ByVal bDriveNum As Byte, ByVal lpcbBytesReturned As Long) As Boolean
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte

      ' Set up data structures for IDENTIFY command.
   SCIP.cBufferSize = IDENTIFY_BUFFER_SIZE
   SCIP.irDriveRegs.bFeaturesReg = 0
   SCIP.irDriveRegs.bSectorCountReg = 1
   SCIP.irDriveRegs.bSectorNumberReg = 1
   SCIP.irDriveRegs.bCylLowReg = 0
   SCIP.irDriveRegs.bCylHighReg = 0

      ' Compute the drive number.
   SCIP.irDriveRegs.bDriveHeadReg = &HA0 Or ShiftLeft((bDriveNum And 1), 4)

      ' The command can either be IDE identify or ATAPI identify.
   SCIP.irDriveRegs.bCommandReg = bIDCmd
   SCIP.bDriveNumber = bDriveNum
   SCIP.cBufferSize = IDENTIFY_BUFFER_SIZE

   DoIdentify = DeviceIoControl(hPhysicalDriveIOCTL, DFP_RECEIVE_DRIVE_DATA, _
               SCIP, _
               Len(SCIP) - 4, _
               bArrOut(0), _
               OUTPUT_DATA_SIZE, _
               lpcbBytesReturned, 0)
    bOut = bArrOut
End Function



'The following are bit shifting functions taken from VBSpeed
'(http:'www.xbeat.net/vbspeed/)
'Stupid VB and its lack of << and >>... =(

Private Function ShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011010
  Select Case ShiftCount
  Case 0&:  ShiftRight = Value
  Case 1&:  ShiftRight = (Value And &HFFFFFFFE) \ &H2&
  Case 2&:  ShiftRight = (Value And &HFFFFFFFC) \ &H4&
  Case 3&:  ShiftRight = (Value And &HFFFFFFF8) \ &H8&
  Case 4&:  ShiftRight = (Value And &HFFFFFFF0) \ &H10&
  Case 5&:  ShiftRight = (Value And &HFFFFFFE0) \ &H20&
  Case 6&:  ShiftRight = (Value And &HFFFFFFC0) \ &H40&
  Case 7&:  ShiftRight = (Value And &HFFFFFF80) \ &H80&
  Case 8&:  ShiftRight = (Value And &HFFFFFF00) \ &H100&
  Case 9&:  ShiftRight = (Value And &HFFFFFE00) \ &H200&
  Case 10&: ShiftRight = (Value And &HFFFFFC00) \ &H400&
  Case 11&: ShiftRight = (Value And &HFFFFF800) \ &H800&
  Case 12&: ShiftRight = (Value And &HFFFFF000) \ &H1000&
  Case 13&: ShiftRight = (Value And &HFFFFE000) \ &H2000&
  Case 14&: ShiftRight = (Value And &HFFFFC000) \ &H4000&
  Case 15&: ShiftRight = (Value And &HFFFF8000) \ &H8000&
  Case 16&: ShiftRight = (Value And &HFFFF0000) \ &H10000
  Case 17&: ShiftRight = (Value And &HFFFE0000) \ &H20000
  Case 18&: ShiftRight = (Value And &HFFFC0000) \ &H40000
  Case 19&: ShiftRight = (Value And &HFFF80000) \ &H80000
  Case 20&: ShiftRight = (Value And &HFFF00000) \ &H100000
  Case 21&: ShiftRight = (Value And &HFFE00000) \ &H200000
  Case 22&: ShiftRight = (Value And &HFFC00000) \ &H400000
  Case 23&: ShiftRight = (Value And &HFF800000) \ &H800000
  Case 24&: ShiftRight = (Value And &HFF000000) \ &H1000000
  Case 25&: ShiftRight = (Value And &HFE000000) \ &H2000000
  Case 26&: ShiftRight = (Value And &HFC000000) \ &H4000000
  Case 27&: ShiftRight = (Value And &HF8000000) \ &H8000000
  Case 28&: ShiftRight = (Value And &HF0000000) \ &H10000000
  Case 29&: ShiftRight = (Value And &HE0000000) \ &H20000000
  Case 30&: ShiftRight = (Value And &HC0000000) \ &H40000000
  Case 31&: ShiftRight = CBool(Value And &H80000000)
  End Select
End Function
Private Function ShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011001
  Select Case ShiftCount
  Case 0&
    ShiftLeft = Value
  Case 1&
    If Value And &H40000000 Then
      ShiftLeft = (Value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft = (Value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If Value And &H20000000 Then
      ShiftLeft = (Value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft = (Value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If Value And &H10000000 Then
      ShiftLeft = (Value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft = (Value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If Value And &H8000000 Then
      ShiftLeft = (Value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft = (Value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If Value And &H4000000 Then
      ShiftLeft = (Value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft = (Value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If Value And &H2000000 Then
      ShiftLeft = (Value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft = (Value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If Value And &H1000000 Then
      ShiftLeft = (Value And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft = (Value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If Value And &H800000 Then
      ShiftLeft = (Value And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft = (Value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If Value And &H400000 Then
      ShiftLeft = (Value And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft = (Value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If Value And &H200000 Then
      ShiftLeft = (Value And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft = (Value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If Value And &H100000 Then
      ShiftLeft = (Value And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft = (Value And &HFFFFF) * &H800&
    End If
  Case 12&
    If Value And &H80000 Then
      ShiftLeft = (Value And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft = (Value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If Value And &H40000 Then
      ShiftLeft = (Value And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft = (Value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If Value And &H20000 Then
      ShiftLeft = (Value And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft = (Value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If Value And &H10000 Then
      ShiftLeft = (Value And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft = (Value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If Value And &H8000& Then
      ShiftLeft = (Value And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft = (Value And &H7FFF&) * &H10000
    End If
  Case 17&
    If Value And &H4000& Then
      ShiftLeft = (Value And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft = (Value And &H3FFF&) * &H20000
    End If
  Case 18&
    If Value And &H2000& Then
      ShiftLeft = (Value And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft = (Value And &H1FFF&) * &H40000
    End If
  Case 19&
    If Value And &H1000& Then
      ShiftLeft = (Value And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft = (Value And &HFFF&) * &H80000
    End If
  Case 20&
    If Value And &H800& Then
      ShiftLeft = (Value And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft = (Value And &H7FF&) * &H100000
    End If
  Case 21&
    If Value And &H400& Then
      ShiftLeft = (Value And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft = (Value And &H3FF&) * &H200000
    End If
  Case 22&
    If Value And &H200& Then
      ShiftLeft = (Value And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft = (Value And &H1FF&) * &H400000
    End If
  Case 23&
    If Value And &H100& Then
      ShiftLeft = (Value And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft = (Value And &HFF&) * &H800000
    End If
  Case 24&
    If Value And &H80& Then
      ShiftLeft = (Value And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H7F&) * &H1000000
    End If
  Case 25&
    If Value And &H40& Then
      ShiftLeft = (Value And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H3F&) * &H2000000
    End If
  Case 26&
    If Value And &H20& Then
      ShiftLeft = (Value And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H1F&) * &H4000000
    End If
  Case 27&
    If Value And &H10& Then
      ShiftLeft = (Value And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft = (Value And &HF&) * &H8000000
    End If
  Case 28&
    If Value And &H8& Then
      ShiftLeft = (Value And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H7&) * &H10000000
    End If
  Case 29&
    If Value And &H4& Then
      ShiftLeft = (Value And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H3&) * &H20000000
    End If
  Case 30&
    If Value And &H2& Then
      ShiftLeft = (Value And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft = (Value And &H1&) * &H40000000
    End If
  Case 31&
    If Value And &H1& Then
      ShiftLeft = &H80000000
    Else
      ShiftLeft = &H0&
    End If
  End Select
End Function
Private Function SwapBytes(B() As Byte) As Byte()
    Dim bTemp As Byte
    Dim cnt As Long

    For cnt = LBound(B) To UBound(B) Step 2
        bTemp = B(cnt)
        B(cnt) = B(cnt + 1)
        B(cnt + 1) = bTemp
    Next cnt
      
    SwapBytes = B()
      
End Function
Private Function FlipAndCodeBytes(B() As Byte) As Byte()
    Dim Flipped() As Byte
    Dim Num As Long
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Dim Y As Long
    Dim Sum As Long
    Y = 0
    ReDim Flipped(0 To 10000)
    Num = UBound(B)
    For i = 0 To Num Step 4
        For J = 1 To 0 Step -1
            Sum = 0
            For k = 0 To 1
                Sum = Sum * 16
                Select Case B(i + J * 2 + k)
                Case 48: Sum = Sum + 0
                Case 49: Sum = Sum + 1
                Case 50: Sum = Sum + 2
                Case 51: Sum = Sum + 3
                Case 52: Sum = Sum + 4
                Case 53: Sum = Sum + 5
                Case 54: Sum = Sum + 6
                Case 55: Sum = Sum + 7
                Case 56: Sum = Sum + 8
                Case 57: Sum = Sum + 9
                Case 97: Sum = Sum + 10
                Case 98: Sum = Sum + 11
                Case 99: Sum = Sum + 12
                Case 100: Sum = Sum + 13
                Case 101: Sum = Sum + 14
                Case 102: Sum = Sum + 15
                End Select
            Next
            If Sum > 0 Then
                Flipped(Y) = CByte(Sum)
                'Flipped(y + 1) = 0
                Y = Y + 1
            End If
        Next
    Next
    If Y <> 0 Then
        ReDim Preserve Flipped(0 To Y - 1)
    Else
        Erase Flipped
    End If
    FlipAndCodeBytes = Flipped
End Function
