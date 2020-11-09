VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------
'ProcessList.cls
'version 1.0
'-----------------------
'by: Martin Sykes
'when: 3rd July 2006
'site: www.martrinex.net
'-----------------------
'this code is a part of my two a week summer challenge so check martrinex.net for more codes like this!

'about:
'This code gets a list of processes currently running on your windows system
'a process is a set of commands in the systems memory, usually a whole program
'like calc.exe however a single progrem may have more then one process.

'Its primary use can be to close programs, but since we all get the apps id we can
'also patch programs / set programs as children and make things like, a new taskbar
'custom alt+tab style program!, task manager infact quite alot of things!
'
'I originaly used this code for my college project to list and close any unreckognised programs
'ie: games / viruses what may not be allowed to run on a business network.


'code for windows NT based systems only NOT win95 or win98!


'>> api used to close ANY process
Private Type LUID: lowpart As Long: highpart As Long: End Type
Private Type TOKEN_PRIVILEGES: PrivilegeCount As Long: LuidUDT As LUID: Attributes As Long: End Type

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'>> api used to get windows version
'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Private Type OSVERSIONINFO
'    OSVSize As Long
'    dwVerMajor As Long: dwVerMinor As Long
'    dwBuildNumber As Long           '
'    PlatformID As Long              '0 win 3.1 / 1 win 9x / 2 win NT
'    szCSDVersion As String * 128    'Service Pack number: 1,2,3
'End Type
'Private Const WIN_NT = 2

'>> api used to list all processes
'Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Const PROCESS_TERMINATE As Long = &H1
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
'Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal moduleName As String, ByVal nSize As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

'>> public variables used to store our processess
Private Type typProcess
    Handle As Long
    exeName As String
End Type
Private Process() As typProcess, intProcessCount As Integer
Event FoundProcess(Index As Integer, Handle As Long, exeName As String)

'gets how many processes are stored up in memory
Public Function processCount() As Integer
 processCount = intProcessCount
End Function

'gets a process exe name
Public Property Get ProcessName(Index) As String
On Error GoTo bailout
 ProcessName = Process(Index).exeName

bailout:
If Err.Number <> 0 Then Err.Raise 1, "ProcessName", "invalid index, index [" & Index & "] not found"
End Property

'gets a process system handle
Public Property Get ProcessHandle(Index) As Long
On Error GoTo bailout
 ProcessHandle = Process(Index).Handle
bailout:
If Err.Number <> 0 Then Err.Raise 1, "ProcessHandle", "invalid index, index [" & Index & "] not found"
End Property


Public Function CheckProcesses() As Integer
'Note that this sub assumes you are running windows NT or later
CheckWindowsVersion

'Unfortunatly windows doesnt tell us how many processes are running,
'so we have to gues and keep increasing our memory until it is not
'all used up, this is the reason for the loops and all the bytesneed
'lookily this is api much faster then vb code so the loops dont create
'too much of a din't in our processor.

'------------------
'>> declare variables needed for this process
'------------------
Dim cb As Long             '<< current writeable bytes in memory (length of array ProcessIDs)
Dim cbNeeded As Long       '<< current written bytes in memory (length of used array ProcessIDs)
Dim ProcessIDs() As Long   '<< stored all the process ID's of running applications
Dim lngResult As Long      '<< results from API calls
Dim moduleName As String   '<< get the process EXE name
Dim hProcess As Long       '<< our process used to read other processes (bypass security)
Dim i As Long              '<< to loop through all processes

'------------------
'>> clean up old memory
'------------------
clearProcesses

'------------------
'>> get a list of open programs (threads)
'------------------
cb = 32 * 4 '<< start with assuming running are: (16 > programs <= 32)
Do
 ReDim ProcessIDs(cb / 4)
 lngResult = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
 If lngResult <> 1 Then Exit Function
 If cb = cbNeeded Then cb = cb * 2 Else Exit Do
Loop

'------------------
'>> loop through all programs and get their filenames
'------------------
For i = 1 To cbNeeded / 4
   DoEvents
   '>> open a process what allows us to peak at information within another program
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
   '>> if it worked, get the programs EXE name and path
   If hProcess <> 0 And ProcessIDs(i) <> 0 Then

       '>> get the filenames lenth
       moduleName = Space(256)
       lngResult = GetModuleFileNameExA(hProcess, 0, moduleName, Len(moduleName))
       'Debug.Print moduleName
       '>> notify this program of the new program!
       If lngResult <> 0 Then
       
                moduleName = Left(moduleName, lngResult)
                Debug.Print moduleName
'         Call FoundProcess(ProcessIDs(i), moduleName)
         Call appendProcess(ProcessIDs(i), moduleName)
       End If
   End If
   '>> close our process
   lngResult = CloseHandle(hProcess)
Next
End Function

'>> clean up our memory
Public Sub clearProcesses()
 ReDim Process(0) As typProcess
 intProcessCount = 0
End Sub

'>> add a process on to the bottom of our list
Private Function appendProcess(Handle As Long, exeName As String) As Integer
 ReDim Preserve Process(intProcessCount)
 With Process(intProcessCount)

    .exeName = exeName
    .Handle = Handle
 End With
 appendProcess = intProcessCount
 intProcessCount = intProcessCount + 1
 
 RaiseEvent FoundProcess(appendProcess, Handle, exeName)
End Function

Private Sub CheckWindowsVersion()
 '   Dim uOSInfo As OSVERSIONINFO
 '   'set the variables size (making room to hold the information)
 '   uOSInfo.OSVSize = Len(uOSInfo)
 '   'use the windows API call to get the OS version
 '   Call GetVersionEx(uOSInfo)
 '   'check the information and format it to more english
 '   If uOSInfo.PlatformID <> WIN_NT Then
        'MsgBox "This application cannot be run on this operating system"
If GetVersion < 0 Then Err.Raise 2, "ProcessList", "This module can only be run on NT based operating systems"
 '   End If
End Sub

' Terminate any application and return an exit code to Windows
Public Function KillProcess(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) As Boolean
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
    
    ' Windows NT/2000 require a special treatment
    ' to ensure that the calling process has the
    ' privileges to shut down the system
    
    ' under NT the high-order bit (that is, the sign bit)
    ' of the value retured by GetVersion is cleared
    If GetVersion() >= 0 Then
        ' open the tokens for the current process
        ' exit if any error
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
            GoTo CleanUp
        End If
        
        ' retrieves the locally unique identifier (LUID) used
        ' to locally represent the specified privilege name
        ' (first argument = "" means the local system)
        ' Exit if any error
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            GoTo CleanUp
        End If
    
        ' complete the TOKEN_PRIVILEGES structure with the # of
        ' privileges and the desired attribute
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
    
        ' try to acquire debug privilege for this process
        ' exit if error
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
    End If
    
    ' now we can finally open the other process
    ' while having complete access on its attributes
    ' exit if any error
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        ' call was successful, so we can kill the application
        ' set return value for this function
        KillProcess2 = (TerminateProcess(hProcess, ExitCode) <> 0)
        ' close the process handle
        CloseHandle hProcess
    End If
    
    If GetVersion() >= 0 Then
        ' under NT restore original privileges
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
        
CleanUp:
        If hToken Then CloseHandle hToken
    End If
End Function

Private Sub Class_initialize()
 Me.CheckProcesses
End Sub

Private Sub Class_Terminate()
 clearProcesses
End Sub
'Kill process functions
Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    On Error Resume Next
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim res As Boolean
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    res = GetExitCodeProcess(hProcess, lExitCode)
    res = TerminateProcess(hProcess, lExitCode)
    CloseHandle (hProcess)
End Sub



