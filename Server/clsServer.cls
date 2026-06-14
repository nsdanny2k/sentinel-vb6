VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WS_Port As Long

'// socket code =================================================================
Private Const mbCleanupSock As Boolean = False

Private Remote_Buffer As sockaddr
Private Read_Buffer As String * 1024

Private sockBuffer As sockaddr
Private udtWSAData As WSADataType

Private Const FD_SETSIZE = 64
Private Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Private Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Const hostent_size = 16

Private Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Private Const servent_size = 14

Private Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Private Const protoent_size = 10

Private Const IPPROTO_TCP = 6
Private Const IPPROTO_UDP = 17

Private Const INADDR_NONE = &HFFFFFFFF
Private Const INADDR_ANY = &H0

Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Const sockaddr_size = 16
Private saZero As sockaddr


Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Private Type WSADataType
    wversion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1

Private Const BAD_SOCKET = -1
Private Const UNRESOLVED_HOST = -2
Private Const UNABLE_TO_BIND = -3
Private Const UNABLE_TO_CONNECT = -4

'Types
Private Const SOCK_STREAM = 1
Private Const SOCK_DGRAM = 2
Private Const SOCK_RAW = 3
Private Const SOCK_RDM = 4
Private Const SOCK_SEQPACKET = 5

Private Const MAXGETHOSTSTRUCT = 1024

Private Const AF_UNSPEC = 0             '  /* unspecified */
Private Const AF_UNIX = 1               '  /* local to host (pipes, portals) */
Private Const AF_INET = 2               '  /* internetwork: UDP, TCP, etc. */
Private Const AF_IMPLINK = 3            '  /* arpanet imp addresses */
Private Const AF_PUP = 4                '  /* pup protocols: e.g. BSP */
Private Const AF_CHAOS = 5              '  /* mit CHAOS protocols */
Private Const AF_IPX = 6                '  /* IPX and SPX */
Private Const AF_NS = 6                 '  /* XEROX NS protocols */
Private Const AF_ISO = 7                '  /* ISO protocols */
Private Const AF_OSI = AF_ISO           '  /* OSI is ISO */
Private Const AF_ECMA = 8               '  /* european computer manufacturers */
Private Const AF_DATAKIT = 9            '  /* datakit protocols */
Private Const AF_CCITT = 10             '  /* CCITT protocols, X.25 etc */
Private Const AF_SNA = 11               '  /* IBM SNA */
Private Const AF_DECnet = 12            '  /* DECnet */
Private Const AF_DLI = 13               '  /* Direct data link interface */
Private Const AF_LAT = 14               '  /* LAT */
Private Const AF_HYLINK = 15            '  /* NSC Hyperchannel */
Private Const AF_APPLETALK = 16         '  /* AppleTalk */
Private Const AF_NETBIOS = 17           '  /* NetBios-style addresses */

'Protocol families, same as address families for now
Private Const PF_UNSPEC = 0
Private Const PF_UNIX = 1
Private Const PF_INET = 2
Private Const PF_IMPLINK = 3
Private Const PF_PUP = 4
Private Const PF_CHAOS = 5
Private Const PF_IPX = 6
Private Const PF_NS = 6
Private Const PF_ISO = 7
Private Const PF_OSI = AF_ISO
Private Const PF_ECMA = 8
Private Const PF_DATAKIT = 9
Private Const PF_CCITT = 10
Private Const PF_SNA = 11
Private Const PF_DECnet = 12
Private Const PF_DLI = 13
Private Const PF_LAT = 14
Private Const PF_HYLINK = 15
Private Const PF_APPLETALK = 16
Private Const PF_NETBIOS = 17

Private Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

' Windows Sockets definitions of regular Microsoft C error constants
Private Const WSAEINTR = 10004
Private Const WSAEBADF = 10009
Private Const WSAEACCES = 10013
Private Const WSAEFAULT = 10014
Private Const WSAEINVAL = 10022
Private Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Private Const WSAEWOULDBLOCK = 10035
Private Const WSAEINPROGRESS = 10036
Private Const WSAEALREADY = 10037
Private Const WSAENOTSOCK = 10038
Private Const WSAEDESTADDRREQ = 10039
Private Const WSAEMSGSIZE = 10040
Private Const WSAEPROTOTYPE = 10041
Private Const WSAENOPROTOOPT = 10042
Private Const WSAEPROTONOSUPPORT = 10043
Private Const WSAESOCKTNOSUPPORT = 10044
Private Const WSAEOPNOTSUPP = 10045
Private Const WSAEPFNOSUPPORT = 10046
Private Const WSAEAFNOSUPPORT = 10047
Private Const WSAEADDRINUSE = 10048
Private Const WSAEADDRNOTAVAIL = 10049
Private Const WSAENETDOWN = 10050
Private Const WSAENETUNREACH = 10051
Private Const WSAENETRESET = 10052
Private Const WSAECONNABORTED = 10053
Private Const WSAECONNRESET = 10054
Private Const WSAENOBUFS = 10055
Private Const WSAEISCONN = 10056
Private Const WSAENOTCONN = 10057
Private Const WSAESHUTDOWN = 10058
Private Const WSAETOOMANYREFS = 10059
Private Const WSAETIMEDOUT = 10060
Private Const WSAECONNREFUSED = 10061
Private Const WSAELOOP = 10062
Private Const WSAENAMETOOLONG = 10063
Private Const WSAEHOSTDOWN = 10064
Private Const WSAEHOSTUNREACH = 10065
Private Const WSAENOTEMPTY = 10066
Private Const WSAEPROCLIM = 10067
Private Const WSAEUSERS = 10068
Private Const WSAEDQUOT = 10069
Private Const WSAESTALE = 10070
Private Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Private Const WSASYSNOTREADY = 10091
Private Const WSAVERNOTSUPPORTED = 10092
Private Const WSANOTINITIALISED = 10093
Private Const WSAHOST_NOT_FOUND = 11001
Private Const WSATRY_AGAIN = 11002
Private Const WSANO_RECOVERY = 11003
Private Const WSANO_DATA = 11004
Private Const WSANO_ADDRESS = 11004
'---ioctl Constants
    Private Const FIONREAD = &H8004667F
    Private Const FIONBIO = &H8004667E
    Private Const FIOASYNC = &H8004667D

#If Win16 Then
'---Windows System functions
    Private Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
    Private Declare Sub MemCopy Lib "Kernel" Alias "hmemcpy" (Dest As Any, src As Any, ByVal cb&)
    Private Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
'---async notification constants
    Private Const SOL_SOCKET = &HFFFF
    Private Const SO_LINGER = &H80
    Private Const FD_READ = &H1
    Private Const FD_WRITE = &H2
    Private Const FD_OOB = &H4
    Private Const FD_ACCEPT = &H8
    Private Const FD_CONNECT = &H10
    Private Const FD_CLOSE = &H20
'---SOCKET FUNCTIONS
    Private Declare Function accept Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, addrlen As Integer) As Integer
    Private Declare Function bind Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Private Declare Function closesocket Lib "Winsock.dll" (ByVal s As Integer) As Integer
    Private Declare Function connect Lib "Winsock.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
    Private Declare Function ioctlsocket Lib "Winsock.dll" (ByVal s As Integer, ByVal cmd As Long, argp As Long) As Integer
    Private Declare Function getpeername Lib "Winsock.dll" (ByVal s As Integer, sName As sockaddr, namelen As Integer) As Integer
    Private Declare Function getsockname Lib "Winsock.dll" (ByVal s As Integer, sName As sockaddr, namelen As Integer) As Integer
    Private Declare Function getsockopt Lib "Winsock.dll" (ByVal s As Integer, ByVal Level As Integer, ByVal optname As Integer, optval As Any, optlen As Integer) As Integer
    Private Declare Function htonl Lib "Winsock.dll" (ByVal hostlong As Long) As Long
    Private Declare Function htons Lib "Winsock.dll" (ByVal hostshort As Integer) As Integer
    Private Declare Function inet_addr Lib "Winsock.dll" (ByVal cp As String) As Long
    Private Declare Function inet_ntoa Lib "Winsock.dll" (ByVal inn As Long) As Long
    Private Declare Function listen Lib "Winsock.dll" (ByVal s As Integer, ByVal backlog As Integer) As Integer
    Private Declare Function ntohl Lib "Winsock.dll" (ByVal netlong As Long) As Long
    Private Declare Function ntohs Lib "Winsock.dll" (ByVal netshort As Integer) As Integer
    Private Declare Function recv Lib "Winsock.dll" (ByVal s As Integer, ByVal buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Private Declare Function recvfrom Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer, from As sockaddr, fromlen As Integer) As Integer
    Private Declare Function ws_select Lib "Winsock.dll" Alias "select" (ByVal nfds As Integer, readfds As Any, writefds As Any, exceptfds As Any, Timeout As timeval) As Integer
    Private Declare Function send Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
    Private Declare Function sendto Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer, to_addr As sockaddr, ByVal tolen As Integer) As Integer
    Private Declare Function setsockopt Lib "Winsock.dll" (ByVal s As Integer, ByVal Level As Integer, ByVal optname As Integer, optval As Any, ByVal optlen As Integer) As Integer
    Private Declare Function ShutDown Lib "Winsock.dll" Alias "shutdown" (ByVal s As Integer, ByVal how As Integer) As Integer
    Private Declare Function Socket Lib "Winsock.dll" Alias "socket" (ByVal af As Integer, ByVal s_type As Integer, ByVal protocol As Integer) As Integer
'---DATABASE FUNCTIONS
    Private Declare Function gethostbyaddr Lib "Winsock.dll" (addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer) As Long
    Private Declare Function gethostbyname Lib "Winsock.dll" (ByVal host_name As String) As Long
    Private Declare Function gethostname Lib "Winsock.dll" (ByVal host_name As String, ByVal namelen As Integer) As Integer
    Private Declare Function getservbyport Lib "Winsock.dll" (ByVal Port As Integer, ByVal proto As String) As Long
    Private Declare Function getservbyname Lib "Winsock.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Private Declare Function getprotobynumber Lib "Winsock.dll" (ByVal proto As Integer) As Long
    Private Declare Function getprotobyname Lib "Winsock.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Private Declare Function WSAStartup Lib "Winsock.dll" (ByVal wVR As Integer, lpWSAD As WSADataType) As Integer
    Private Declare Function WSACleanup Lib "Winsock.dll" () As Integer
    Private Declare Sub WSASetLastError Lib "Winsock.dll" (ByVal iError As Integer)
    Private Declare Function WSAGetLastError Lib "Winsock.dll" () As Integer
    Private Declare Function WSAIsBlocking Lib "Winsock.dll" () As Integer
    Private Declare Function WSAUnhookBlockingHook Lib "Winsock.dll" () As Integer
    Private Declare Function WSASetBlockingHook Lib "Winsock.dll" (ByVal lpBlockFunc As Long) As Long
    Private Declare Function WSACancelBlockingCall Lib "Winsock.dll" () As Integer
    Private Declare Function WSAAsyncGetServByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSAAsyncGetServByPort Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Port As Integer, ByVal proto As String, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSAAsyncGetProtoByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal proto_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSAAsyncGetProtoByNumber Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal Number As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSAAsyncGetHostByName Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal host_name As String, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSAAsyncGetHostByAddr Lib "Winsock.dll" (ByVal hWnd As Integer, ByVal wMsg As Integer, addr As Long, ByVal addr_len As Integer, ByVal addr_type As Integer, buf As Any, ByVal buflen As Integer) As Integer
    Private Declare Function WSACancelAsyncRequest Lib "Winsock.dll" (ByVal hAsyncTaskHandle As Integer) As Integer
    Private Declare Function WSAAsyncSelect Lib "Winsock.dll" (ByVal s As Integer, ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal lEvent As Long) As Integer
    Private Declare Function WSARecvEx Lib "Winsock.dll" (ByVal s As Integer, buf As Any, ByVal buflen As Integer, ByVal flags As Integer) As Integer
#ElseIf Win32 Then
'---Windows System Functions
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal cb&)
    Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
    Private Const SOL_SOCKET = &HFFFF&
    Private Const SO_LINGER = &H80&
    Private Const FD_READ = &H1&
    Private Const FD_WRITE = &H2&
    Private Const FD_OOB = &H4&
    Private Const FD_ACCEPT = &H8&
    Private Const FD_CONNECT = &H10&
    Private Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Private Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
    Private Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
    Private Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Private Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
    Private Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Private Declare Function getsockname Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Private Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Private Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Private Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Private Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Private Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
    Private Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, Timeout As timeval) As Long
    Private Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Private Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Private Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Private Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
    Private Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Private Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Private Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Private Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Private Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Private Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Private Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Private Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Private Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Private Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Private Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Private Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Private Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
#End If


'SOME STUFF I ADDED
Private MySocket%
Private SockReadBuffer$
Private Const WSA_NoName = "Unknown"
Private WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled
Private retf As Integer

Private Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetAsyncBufLen = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetAsyncBufLen = lParam And &HFFFF&
    End If
End Function

Private Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function

Private Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Private Function AddrToIP(ByVal AddrOrIP$) As String
    AddrToIP$ = GetAscIP(GetHostByNameAlias(AddrOrIP$))
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
    Private Function ConnectSock(ByVal HOST$, ByVal Port%, retIpPort$, ByVal HWndToMsg%, ByVal Async%) As Integer
    Dim s%, SelectOps%, dummy%
#ElseIf Win32 Then
    Private Function ConnectSock(ByVal HOST$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long
    Dim s&, SelectOps&, dummy&
#End If
    Dim sockin As sockaddr
    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(HOST$)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    retIpPort$ = GetAscIP$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    s = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(s, 1, 0) = SOCKET_ERROR Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If Not connect(s, sockin, sockaddr_size) = 0 Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If HWndToMsg <> 0 Then
            SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
            If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
                If s > 0 Then
                    dummy = closesocket(s)
                End If
                ConnectSock = INVALID_SOCKET
                Exit Function
            End If
        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If connect(s, sockin, sockaddr_size) <> -1 Then
            If s > 0 Then
                dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    ConnectSock = s
End Function

#If Win32 Then
    Private Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
#Else
    Private Function SetSockLinger(ByVal SockNum%, ByVal OnOff%, ByVal LingerTime%) As Integer
#End If
    Dim Linger As LingerType
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger
        End If
    End If
End Function

Private Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

Private Function GetAscIP(ByVal inn As Long) As String
    #If Win32 Then
        Dim nStr&
    #Else
        Dim nStr%
    #End If
    Dim lpStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "255.255.255.255"
    End If
End Function

Private Function GetHostByAddress(ByVal addr As Long) As String
    Dim phe&, ret&
    Dim heDestHost As HOSTENT
    Dim Hostname$
    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe Then
        MemCopy heDestHost, ByVal phe, hostent_size
        Hostname = String(256, 0)
        MemCopy ByVal Hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left(Hostname, InStr(Hostname, Chr(0)) - 1)
    Else
        GetHostByAddress = WSA_NoName
    End If
End Function

'returns IP as long, in network byte order
Private Function GetHostByNameAlias(ByVal Hostname$) As Long
    'Return IP address as a long, in network byte order
    Dim phe&
    Dim heDestHost As HOSTENT
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(Hostname$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(Hostname$)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function

'returns your local machines name
Private Function GetLocalHostName() As String
    Dim sName$
    sName = String(256, 0)
    If gethostname(sName, 256) Then
        sName = WSA_NoName
    Else
        If InStr(sName, Chr(0)) Then
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
        End If
    End If
    GetLocalHostName = sName
End Function

#If Win16 Then
    Private Function GetPeerAddress(ByVal s%) As String
    Dim addrlen%
#ElseIf Win32 Then
    Private Function GetPeerAddress(ByVal s&) As String
    Dim addrlen&
#End If
    Dim SA As sockaddr
    addrlen = sockaddr_size
    If getpeername(s, SA, addrlen) Then
        GetPeerAddress = ""
    Else
        GetPeerAddress = SockAddressToString(SA)
    End If
End Function

#If Win16 Then
    Private Function GetPortFromString(ByVal PortStr$) As Integer
#ElseIf Win32 Then
    Private Function GetPortFromString(ByVal PortStr$) As Long
#End If
    'sometimes users provide ports outside the range of a VB
    'integer, so this function returns an integer for a string
    'just to keep an error from happening, it converts the
    'number to a negative if needed
    If Val(PortStr$) > 32767 Then
        GetPortFromString = CInt(Val(PortStr$) - &H10000)
    Else
        GetPortFromString = Val(PortStr$)
    End If
    If Err Then GetPortFromString = 0
End Function

#If Win16 Then
    Private Function GetProtocolByName(ByVal protocol$) As Integer
    Dim tmpShort%
#ElseIf Win32 Then
    Private Function GetProtocolByName(ByVal protocol$) As Long
    Dim tmpShort&
#End If
    Dim ppe&
    Dim peDestProt As protoent
    ppe = getprotobyname(protocol)
    If ppe Then
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocolByName = peDestProt.p_proto
    Else
        tmpShort = Val(protocol)
        If tmpShort Then
            GetProtocolByName = htons(tmpShort)
        Else
            GetProtocolByName = SOCKET_ERROR
        End If
    End If
End Function

#If Win16 Then
    Private Function GetServiceByName(ByVal service$, ByVal protocol$) As Integer
    Dim serv%
#ElseIf Win32 Then
    Private Function GetServiceByName(ByVal service$, ByVal protocol$) As Long
    Dim serv&
#End If
    Dim pse&
    Dim seDestServ As servent
    pse = getservbyname(service, protocol)
    If pse Then
        MemCopy seDestServ, ByVal pse, servent_size
        GetServiceByName = seDestServ.s_port
    Else
        serv = Val(service)
        If serv Then
            GetServiceByName = htons(serv)
        Else
            GetServiceByName = INVALID_SOCKET
        End If
    End If
End Function

'this function DOES work on 16 and 32 bit systems
#If Win16 Then
    Private Function GetSockAddress(ByVal s%) As String
    Dim addrlen%
    Dim ret%
#ElseIf Win32 Then
    Private Function GetSockAddress(ByVal s&) As String
    Dim addrlen&
    Dim ret&
#End If
    Dim SA As sockaddr
    Dim szRet$
    szRet = String(32, 0)
    addrlen = sockaddr_size
    If getsockname(s, SA, addrlen) Then
        GetSockAddress = ""
    Else
        GetSockAddress = SockAddressToString(SA)
    End If
End Function

'this function should work on 16 and 32 bit systems
Private Function GetWSAErrorString(ByVal ErrNum&) As String
    On Error Resume Next
    Select Case ErrNum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else:
    End Select
End Function

'this function DOES work on 16 and 32 bit systems
Private Function IpToAddr(ByVal AddrOrIP$) As String
    On Error Resume Next
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))
    If Err Then IpToAddr = WSA_NoName
End Function

'this function DOES work on 16 and 32 bit systems
Private Function IrcGetAscIp(ByVal IPL$) As String
    'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
    'the kind that would be parsed out of a DCC command string
    On Error GoTo IrcGetAscIPError:
    Dim lpStr&
#If Win16 Then
    Dim nStr%
#ElseIf Win32 Then
    Dim nStr&
#End If
    Dim retString$
    Dim inn&
    If Val(IPL) > 2147483647 Then
        inn = Val(IPL) - 4294967296#
    Else
        inn = Val(IPL)
    End If
    inn = ntohl(inn)
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        IrcGetAscIp = "0.0.0.0"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    IrcGetAscIp = retString
    Exit Function
IrcGetAscIPError:
    IrcGetAscIp = "0.0.0.0"
    Exit Function
    Resume
End Function

'this function DOES work on 16 and 32 bit systems
Private Function IrcGetLongIp(ByVal AscIp$) As String
    'this function converts an ascii ip string into a long ip in network byte order
    'and stick it in a string suitable for use in a DCC command.
    On Error GoTo IrcGetLongIpError:
    Dim inn&
    inn = inet_addr(AscIp)
    inn = htonl(inn)
    If inn < 0 Then
        IrcGetLongIp = CVar(inn + 4294967296#)
        Exit Function
    Else
        IrcGetLongIp = CVar(inn)
        Exit Function
    End If
    Exit Function
IrcGetLongIpError:
    IrcGetLongIp = "0"
    Exit Function
    Resume
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
Private Function ListenForConnect(ByVal Port%, ByVal HWndToMsg%) As Integer
    Dim s%, dummy%
    Dim SelectOps%
#ElseIf Win32 Then
Private Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&) As Long
    Dim s&, dummy&
    Dim SelectOps&
#End If
    Dim sockin As sockaddr
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    s = Socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = s
End Function

'this function should work on 16 and 32 bit systems
#If Win16 Then
Private Function SendData(ByVal s%, vMessage As Variant) As Integer
#ElseIf Win32 Then
Private Function SendData(ByVal s&, vMessage As Variant) As Long
#End If
    Dim TheMsg() As Byte, sTemp$
    TheMsg = ""
    Select Case VarType(vMessage)
        Case 8209   'byte array
            sTemp = vMessage
            TheMsg = sTemp
        Case 8      'string, if we recieve a string, its assumed we are linemode
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
            #Else
                sTemp = vMessage
            #End If
        Case Else
            sTemp = CStr(vMessage)
            #If Win32 Then
                sTemp = StrConv(vMessage, vbFromUnicode)
            #Else
                sTemp = vMessage
            #End If
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendData = send(s, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function

Private Function SockAddressToString(SA As sockaddr) As String
    SockAddressToString = GetAscIP(SA.sin_addr) & ":" & ntohs(SA.sin_port)
End Function

Private Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            Debug.Print "wVersion="; StartupData.wversion, "wHighVersion="; StartupData.wHighVersion
            Debug.Print "If wVersion == 257 then everything is kewl"
            Debug.Print "szDescription="; StartupData.szDescription
            Debug.Print "szSystemStatus="; StartupData.szSystemStatus
            Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Private Function WSAMakeSelectReply(TheEvent%, TheError%) As Long
    WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)
End Function

'// socket code =================================================================

Public Function WS_Listen(ByRef lngSocketnumber As Long, lngLocalPort As Long, ByVal Handle As Long) As Long
    Dim lngReturn As Long
    
    '//Startup winsock
    Call WS_CleanUp

    lngReturn = WSAStartup(&H101, udtWSAData)
        
    '//An error occured
    If (lngReturn = SOCKET_ERROR) Then
        MsgBox Error
        Exit Function
    End If
    
    '//Try to create our TCP socket
    lngSocketnumber = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    '//An error occured while trying...
    If (lngSocketnumber = SOCKET_ERROR) Then
        Exit Function
    End If
        
    '//Configure our socket
    With sockBuffer
        .sin_family = AF_INET
        .sin_port = htons(WS_Port)
        .sin_addr = 0
        .sin_zero = String$(8, 0)
    End With
                        
    '//Bind socket...
    lngReturn = bind(lngSocketnumber, sockBuffer, sockaddr_size)
    
    '//An error occured while trying to bind
    If Not lngReturn = 0 Then
        Call WS_CleanUp
        Exit Function
    End If
    
    '//Start listening for connections
    lngReturn = listen(lngSocketnumber, 1)
    
    '//Perform Async functions, with our form's hwnd
    WS_Listen = WSAAsyncSelect(lngSocketnumber, Handle, &H202, FD_CONNECT Or FD_ACCEPT)
End Function

Public Function WS_Accept(ByVal ServerHandle As Long, ByVal ClientHandle As Long) As Long
    Dim lngSocketHandle As Long
    Dim lngReadSocket As Long
    Dim lngReturn As Long
    Dim nHwnd As Long
    lngReadSocket = accept(ServerHandle, Remote_Buffer, Len(Remote_Buffer))
    lngReturn = WSAAsyncSelect(lngReadSocket, ClientHandle, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
    WS_Accept = lngReadSocket
End Function

Public Function WS_Recv(ByVal lngSocket As Long, strBuffer As String) As Long
    WS_Recv = recv(lngSocket, strBuffer, 1024, 0)
End Function

Public Function WS_Send(ByVal lngSocket As Long, strMessage As String) As Long
    Dim TheMsg() As Byte, sTemp As String
    TheMsg = ""
    sTemp = StrConv(strMessage, vbFromUnicode)
    TheMsg = sTemp
    WS_Send = send(lngSocket, TheMsg(0), UBound(TheMsg) + 1, 0)
End Function

Public Function WS_SendByte(ByVal lngSocket As Long, TheMsg() As Byte) As Long
    WS_SendByte = send(lngSocket, TheMsg(0), UBound(TheMsg) + 1, 0)
End Function

Public Sub WS_CloseSocket(ByVal lngSocket As Long)
    closesocket lngSocket
End Sub

Public Sub WS_CleanUp()
    On Error Resume Next
    Dim lngReturn As Long
    If mbCleanupSock Then
        lngReturn = WSACleanup()
    End If
End Sub

Public Function WS_GetIP(lngSocket As Long) As String
    Dim lngPosition1 As Long
    Dim strIPAddress As String
    
    strIPAddress = GetPeerAddress(lngSocket)

    lngPosition1 = InStr(1, strIPAddress, ":")
    
    '//Return everything before the ':'
    WS_GetIP = Left$(strIPAddress, lngPosition1 - 1)
End Function

Public Function WS_GetServerIP() As String
'    Dim lngPosition1 As Long
'    Dim strIPAddress As String
'    strIPAddress = GetIPEndPointField(lngServerSock, LOCAL_HOST_IP)
'    lngPosition1 = InStr(1, strIPAddress, ":")
'    '//Return everything before the ':'
'    If lngPosition1 = 0 Then
'        oserver.ws_GetServerIP = strIPAddress
'    Else
'        oserver.ws_GetServerIP = Left$(strIPAddress, lngPosition1 - 1)
'    End If
End Function

Private Sub Class_Initialize()
    WS_Port = 80 '// Port 80 is HTTP
End Sub



