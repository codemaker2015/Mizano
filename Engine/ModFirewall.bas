Attribute VB_Name = "ModFirewall"
Public cINIFile As New cINI
Public cINFO As New cINI
Public namePort As String
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
    End Type
Private pTablePtr As Long
    Private pDataRef As Long
    Public nRows As Long
    Public oRows As Long
    Private nCurrentRow As Long
    Private udtRow As MIB_TCPROW
    Private nState As Long
    Private nLocalAddr As Long
    Private nLocalPort As Long
    Private nRemoteAddr As Long
    Private nRemotePort As Long
    Private nProcId As Long
    Public nRet As Long
    Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long
Public Function TerminateThisConnection(rNum As Long) As Boolean
    
    On Error GoTo ErrorTrap
    
    udtRow.dwLocalAddr = Connection(rNum).LocalHost
    udtRow.dwLocalPort = Connection(rNum).LocalPort
    udtRow.dwRemoteAddr = Connection(rNum).LocalHost
    udtRow.dwRemotePort = Connection(rNum).LocalPort
    udtRow.dwState = 12
    SetTcpEntry udtRow
    
    TerminateThisConnection = True
    Exit Function
    
ErrorTrap:
    TerminateThisConnection = False
    
End Function

