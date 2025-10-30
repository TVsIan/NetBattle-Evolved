Attribute VB_Name = "modFunctionDelegator"
Option Explicit
'This lovely module allows VB to call functions from function pointers.
'Source: http://www.fawcette.com/archives/premier/mgznarch/vbpj/2000/02feb00/mc0200/mc0200.asp


'Here's the magic asm for doing the function pointer call.
'The stack comes in with the following:
'esp: return address
'esp + 4: this pointer for FunctionDelegator
'All that we need to do is remove the this pointer from the
'stack, replace it with the return address, then jmp to the
'correct function.  In other words, we're just squeezing the
'this pointer completely out of the picture.
'The code is:
'pop ecx (stores return address)
'pop eax (gets the this pointer)
'push ecx (restores the return address)
'jmp DWORD PTR [eax + 4] (jump to address at this + 4, 3 byte instruction)
'The corresponding byte stream for this is: 59 58 51 FF 60 04
'We pad these six bytes with two int 3 commands (CC CC) to get eight
'bytes, which can be stored in a Currency constant.
'Note that the memory location of this constant is not executable, so
'it must be copied into a currency variable.  The address of the variable
'is then used as the forwarding function.


'Ubiquitous helper function
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const cDelegateASM As Currency = -368956918007638.6215@

Private m_DelegateASM As Currency

Private Type DelegatorVTables
    VTable(7) As Long 'OKQI vtable in 0 to 3, FailQI vtable in 4 to 7
End Type

'Structure for a stack allocated Delegator
Private m_VTables As DelegatorVTables
Private m_pVTableOKQI As Long       'Pointer to vtable, no allocation version
Private m_pVTableFailQI As Long     'Pointer to vtable, no allocation version
Public Type FunctionDelegator
    pVTable As Long  'This has to stay at offset 0
    FuncPtr As Long      'This has to stay at offset 4
End Type

'Functions to initialize a Delegator object on an existing FunctionDelegator
Public Function InitDelegator(Delegator As FunctionDelegator, Optional ByVal pfn As Long) As IUnknown
    If m_pVTableOKQI = 0 Then InitVTables
    With Delegator
        .pVTable = m_pVTableOKQI
        .FuncPtr = pfn
    End With
    CopyMemory InitDelegator, VarPtr(Delegator), 4
End Function
Private Sub InitVTables()
Dim pAddRefRelease As Long
    With m_VTables
        .VTable(0) = FuncAddr(AddressOf QueryInterfaceOK)
        .VTable(4) = FuncAddr(AddressOf QueryInterfaceFail)
        pAddRefRelease = FuncAddr(AddressOf AddRefRelease)
        .VTable(1) = pAddRefRelease
        .VTable(5) = pAddRefRelease
        .VTable(2) = pAddRefRelease
        .VTable(6) = pAddRefRelease
        m_DelegateASM = cDelegateASM
        .VTable(3) = VarPtr(m_DelegateASM)
        .VTable(7) = .VTable(3)
        m_pVTableOKQI = VarPtr(.VTable(0))
        m_pVTableFailQI = VarPtr(.VTable(4))
    End With
End Sub
Private Function QueryInterfaceOK(This As FunctionDelegator, riid As Long, pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.pVTable = m_pVTableFailQI
End Function
Private Function AddRefRelease(ByVal This As Long) As Long
    'Nothing to do, memory not refcounted
End Function

Private Function QueryInterfaceFail(ByVal This As Long, riid As Long, pvObj As Long) As Long
    pvObj = 0
    QueryInterfaceFail = &H80004002 'E_NOINTERFACE
End Function

Private Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

