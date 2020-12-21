Attribute VB_Name = "mdlDnsNameToIpAddr"
Option Explicit

'-----------------------------------------------------------------------------------------
' Win32 API definitions

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Public Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
    (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
  
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
    (ByVal szHost As String, ByVal dwHostLen As Long) As Long
  
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
    (ByVal szHost As String) As Long
  
Public Declare Sub CopyMemoryFromPtr Lib "KERNEL32" Alias "RtlMoveMemory" _
    (dest As Any, ByVal src As Any, ByVal numBytes As Long)

Public Function HiByte(ByVal wParam As Integer) As Byte
    HiByte = (wParam And &HFF00&) \ (&H100)
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte
    LoByte = wParam And &HFF&
End Function

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim sLoByte As String
    Dim sHiByte As String
    
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "Error to initialize 32-bit Windows Sockets library"
        SocketsInitialize = False
        Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
            (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
            HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "32-bit Windows Sockets does not support " & sLoByte & "." & sHiByte & " version"
        
        SocketsInitialize = False
        Exit Function
    End If
    
    SocketsInitialize = True
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Error to free 32-bit Windows Sockets library"
    End If
End Sub

Public Function GetIPAddressFromDNSName(ByVal sDNSName As String) As String
    Dim ptrHost As Long
    Dim vHostEnt As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    
    GetIPAddressFromDNSName = ""
    If Not SocketsInitialize() Then
        Exit Function
    End If
    
    ptrHost = gethostbyname(sDNSName)
    If ptrHost = 0 Then
        MsgBox "Error to call Windows API gethostbyname!"
        SocketsCleanup
        Exit Function
    End If
    
    CopyMemoryFromPtr vHostEnt, ptrHost, Len(vHostEnt)
    CopyMemoryFromPtr dwIPAddr, vHostEnt.hAddrList, 4
    
    ReDim tmpIPAddr(vHostEnt.hLen - 1)
    CopyMemoryFromPtr tmpIPAddr(0), dwIPAddr, vHostEnt.hLen
    
    For i = 0 To vHostEnt.hLen - 1
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    GetIPAddressFromDNSName = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function


