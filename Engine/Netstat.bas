Attribute VB_Name = "Netstat"

Option Explicit

'-------------------------------------------------------------------------------
'Types and function for the ICMP table:
Public lc
Public MIBICMPSTATS As MIBICMPSTATS

Public Type MIBICMPSTATS
    dwEchos As Long
    dwEchoReps As Long
End Type

Public MIBICMPINFO As MIBICMPINFO
Public Type MIBICMPINFO
    icmpOutStats As MIBICMPSTATS
End Type

Public MIB_ICMP As MIB_ICMP
Public Type MIB_ICMP
    stats As MIBICMPINFO
End Type

Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIBICMPINFO) As Long
Public Last_ICMP_Cnt As Integer 'ICMP count

'-------------------------------------------------------------------------------
'Types and functions for the TCP table:

Type MIB_TCPROW
  dwState As Long
  dwLocalAddr As Long
  dwLocalPort As Long
  dwRemoteAddr As Long
  dwRemotePort As Long
End Type

Type MIB_TCPTABLE
  dwNumEntries As Long
  table(2000) As MIB_TCPROW
End Type

Public MIB_TCPTABLE As MIB_TCPTABLE

Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function SetTcpEntry Lib "IPhlpAPI" (pTcpRow As MIB_TCPROW) As Long 'This is used to close an open port.
Public IP_States(13) As String
Private Last_Tcp_Cnt As Integer 'TCP connection count

'-------------------------------------------------------------------------------
'Types and functions for winsock:

Private Const AF_INET = 2
Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const SOCKET_ERROR As Long = -1
Private Const WS_VERSION_REQD As Long = &H101

Type HOSTENT
    h_name As Long        ' official name of host
    h_aliases As Long     ' alias list
    h_addrtype As Integer ' host address type
    h_length As Integer   ' length of address
    h_addr_list As Long   ' list of addresses
End Type

Type servent
  s_name As Long            ' (pointer to string) official service name
  s_aliases As Long         ' (pointer to string) alias list (might be null-seperated with 2null terminated)
  s_port As Long            ' port #
  s_proto As Long           ' (pointer to) protocol to use
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal CP As String) As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (Addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" (ByVal lpString As Any) As Integer
Private Blocked As Boolean

Private Type Ports_
RemA As String
RemP As String
State As String
LocP As String
End Type
Public Ports(0 To 2000) As Ports_



'other imputs

'================= TCP things ====================
'state of the connection
Public Const MIB_TCP_STATE_CLOSED = 0
Public Const MIB_TCP_STATE_LISTEN = 1
Public Const MIB_TCP_STATE_SYN_SENT = 2
Public Const MIB_TCP_STATE_SYN_RCVD = 3
Public Const MIB_TCP_STATE_ESTAB = 4
Public Const MIB_TCP_STATE_FIN_WAIT1 = 5
Public Const MIB_TCP_STATE_FIN_WAIT2 = 6
Public Const MIB_TCP_STATE_CLOSE_WAIT = 7
Public Const MIB_TCP_STATE_CLOSING = 8
Public Const MIB_TCP_STATE_LAST_ACK = 9
Public Const MIB_TCP_STATE_TIME_WAIT = 10
Public Const MIB_TCP_STATE_DELETE_TCB = 11

'================= UDP things ====================
Type MIB_UDPROW
  dwLocalAddr As String * 4 'address on local computer
  dwLocalPort As String * 4 'port number on local computer
End Type

Type MIB_UDPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_UDPROW   'table of MIB_UDPROW structs
End Type

Declare Function GetUdpTable Lib "IPhlpAPI" _
  (pUdpTable As MIB_UDPTABLE, pdwSize As Long, bOrder As Long) As Long

'================= Statistics ====================
Type MIB_IPSTATS
  dwForwarding As Long       ' IP forwarding enabled or disabled
  dwDefaultTTL As Long       ' default time-to-live
  dwInReceives As Long       ' datagrams received
  dwInHdrErrors As Long      ' received header errors
  dwInAddrErrors As Long     ' received address errors
  dwForwDatagrams As Long    ' datagrams forwarded
  dwInUnknownProtos As Long  ' datagrams with unknown protocol
  dwInDiscards As Long       ' received datagrams discarded
  dwInDelivers As Long       ' received datagrams delivered
  dwOutRequests As Long      '
  dwRoutingDiscards As Long  '
  dwOutDiscards As Long      ' sent datagrams discarded
  dwOutNoRoutes As Long      ' datagrams for which no route
  dwReasmTimeout As Long     ' datagrams for which all frags didn't arrive
  dwReasmReqds As Long       ' datagrams requiring reassembly
  dwReasmOks As Long         ' successful reassemblies
  dwReasmFails As Long       ' failed reassemblies
  dwFragOks As Long          ' successful fragmentations
  dwFragFails As Long        ' failed fragmentations
  dwFragCreates As Long      ' datagrams fragmented
  dwNumIf As Long           ' number of interfaces on computer
  dwNumAddr As Long         ' number of IP address on computer
  dwNumRoutes As Long       ' number of routes in routing table
End Type

Declare Function GetIpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_IPSTATS) As Long


Type MIB_TCPSTATS
  dwRtoAlgorithm As Long    ' timeout algorithm
  dwRtoMin As Long          ' minimum timeout
  dwRtoMax As Long          ' maximum timeout
  dwMaxConn As Long         ' maximum connections
  dwActiveOpens As Long     ' active opens
  dwPassiveOpens As Long    ' passive opens
  dwAttemptFails As Long    ' failed attempts
  dwEstabResets As Long     ' establised connections reset
  dwCurrEstab As Long       ' established connections
  dwInSegs As Long          ' segments received
  dwOutSegs As Long         ' segment sent
  dwRetransSegs As Long     ' segments retransmitted
  dwInErrs As Long          ' incoming errors
  dwOutRsts As Long         ' outgoing resets
  dwNumConns As Long        ' cumulative connections
End Type

Declare Function GetTcpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_TCPSTATS) As Long

Type MIB_UDPSTATS
  dwInDatagrams As Long    ' received datagrams
  dwNoPorts As Long        ' datagrams for which no port
  dwInErrors As Long       ' errors on received datagrams
  dwOutDatagrams As Long   ' sent datagrams
  dwNumAddrs As Long       ' number of entries in UDP listener table
End Type

Declare Function GetUdpStatistics Lib "IPhlpAPI" _
  (pStats As MIB_UDPSTATS) As Long

Function c_state(s) As String
  Select Case s
  Case MIB_TCP_STATE_CLOSED: c_state = "CLOSED"
  Case MIB_TCP_STATE_LISTEN: c_state = "LISTENING"
  Case MIB_TCP_STATE_SYN_SENT: c_state = "SYN_SENT"
  Case MIB_TCP_STATE_SYN_RCVD: c_state = "SYN_RCVD"
  Case MIB_TCP_STATE_ESTAB: c_state = "ESTABLISHED"
  Case MIB_TCP_STATE_FIN_WAIT1: c_state = "FIN_WAIT1"
  Case MIB_TCP_STATE_FIN_WAIT2: c_state = "FIN_WAIT2"
  Case MIB_TCP_STATE_CLOSE_WAIT: c_state = "CLOSE_WAIT"
  Case MIB_TCP_STATE_CLOSING: c_state = "CLOSING"
  Case MIB_TCP_STATE_LAST_ACK: c_state = "LAST_ACK"
  Case MIB_TCP_STATE_TIME_WAIT: c_state = "TIME_WAIT"
  Case MIB_TCP_STATE_DELETE_TCB: c_state = "DELETE_TCB"
  Case Else: c_state = "UNKNOWN"
  End Select
End Function

Sub InitStates()
  IP_States(0) = "UNKNOWN"
  IP_States(1) = "CLOSED"
  IP_States(2) = "LISTENING"
  IP_States(3) = "SYN_SENT"
  IP_States(4) = "SYN_RCVD"
  IP_States(5) = "ESTABLISHED"
  IP_States(6) = "FIN_WAIT1"
  IP_States(7) = "FIN_WAIT2"
  IP_States(8) = "CLOSE_WAIT"
  IP_States(9) = "CLOSING"
  IP_States(10) = "LAST_ACK"
  IP_States(11) = "TIME_WAIT"
  IP_States(12) = "DELETE_TCB"
End Sub

Public Function GetAscIP(ByVal inn As Long) As String
  Dim nStr&
    Dim lpStr As Long
    Dim retString As String
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        CopyMemory ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "Unable to get IP"
    End If
End Function

Function c_port(s) As Long
On Error Resume Next
  c_port = Asc(Mid(s, 1, 1)) * 256 + Asc(Mid(s, 2, 1))
End Function




