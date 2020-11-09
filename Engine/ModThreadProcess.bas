Attribute VB_Name = "ModThreadProcess"
Option Explicit

Public Thread() As THREADENTRY32

Public Const Delete As Long = &H10000
Public Const READ_CONTROL As Long = &H20000
Public Const WRITE_DAC As Long = &H40000
Public Const WRITE_OWNER As Long = &H80000
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000
Public Const GENERIC_ALL As Long = &H10000000


Public Const EXCEPTION_NONCONTINUABLE As Long = &H1
Public Const EXCEPTION_MAXIMUM_PARAMETERS As Long = 15


Public Const THREAD_TERMINATE As Long = &H1
Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const THREAD_GET_CONTEXT As Long = &H8
Public Const THREAD_SET_CONTEXT As Long = &H10
Public Const THREAD_SET_INFORMATION As Long = &H20
Public Const THREAD_QUERY_INFORMATION As Long = &H40
Public Const THREAD_SET_THREAD_TOKEN As Long = &H80
Public Const THREAD_IMPERSONATE As Long = &H100
Public Const THREAD_DIRECT_IMPERSONATION As Long = &H200
Public Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)

Public Const THREAD_BASE_PRIORITY_LOWRT As Long = 15
Public Const THREAD_BASE_PRIORITY_MAX As Long = 2
Public Const THREAD_BASE_PRIORITY_MIN As Long = -2
Public Const THREAD_BASE_PRIORITY_IDLE As Long = -15



Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function GetThreadTimes Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function lstrlen Lib "kernel32.dll" (ByVal lpString As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Boolean
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadID As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SetThreadIdealProcessor Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwIdealProcessor As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Boolean
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean

Public Const CREATE_NEW As Long = 1
Public Const CREATE_ALWAYS As Long = 2
Public Const OPEN_EXISTING As Long = 3
Public Const OPEN_ALWAYS As Long = 4
Public Const TRUNCATE_EXISTING As Long = 5

Public Const DEBUG_PROCESS As Long = &H1
Public Const DEBUG_ONLY_THIS_PROCESS As Long = &H2
Public Const CREATE_SUSPENDED As Long = &H4
Public Const DETACHED_PROCESS As Long = &H8
Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const IDLE_PRIORITY_CLASS As Long = &H40
Public Const HIGH_PRIORITY_CLASS As Long = &H80
Public Const REALTIME_PRIORITY_CLASS As Long = &H100
Public Const CREATE_NEW_PROCESS_GROUP As Long = &H200
Public Const CREATE_UNICODE_ENVIRONMENT As Long = &H400
Public Const CREATE_SEPARATE_WOW_VDM As Long = &H800
Public Const CREATE_SHARED_WOW_VDM As Long = &H1000
Public Const CREATE_FORCEDOS As Long = &H2000
Public Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000
Public Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000
Public Const CREATE_BREAKAWAY_FROM_JOB As Long = &H1000000




Public Const HW_PROFILE_GUIDLEN As Long = 39

Public Const MAX_PROFILE_LEN As Long = 80

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Public Const MAXLONG As Long = &H7FFFFFFF

Public Const SEM_FAILCRITICALERRORS As Long = &H1
Public Const SEM_NOGPFAULTERRORBOX As Long = &H2
Public Const SEM_NOALIGNMENTFAULTEXCEPT As Long = &H4
Public Const SEM_NOOPENFILEERRORBOX As Long = &H8000

Public Const THREAD_PRIORITY_LOWEST As Long = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_BELOW_NORMAL As Long = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_NORMAL As Long = 0
Public Const THREAD_PRIORITY_HIGHEST As Long = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_ABOVE_NORMAL As Long = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_ERROR_RETURN As Long = (MAXLONG)
Public Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT
Public Const THREAD_PRIORITY_IDLE As Long = THREAD_BASE_PRIORITY_IDLE

Public Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type HW_PROFILE_INFO
    dwDockInfo As Long
    szHwProfileGuid As String * HW_PROFILE_GUIDLEN
    szHwProfileName As String * MAX_PROFILE_LEN
End Type


'Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwflags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean


Public Const HF32_DEFAULT As Long = 1
Public Const HF32_SHARED As Long = 2

Public Const LF32_FIXED As Long = &H1
Public Const LF32_FREE As Long = &H2
Public Const LF32_MOVEABLE As Long = &H4

Public Const MAX_MODULE_NAME32 As Long = 255

Public Const TH32CS_SNAPHEAPLIST As Long = &H1
'Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const TH32CS_SNAPTHREAD As Long = &H4
Public Const TH32CS_SNAPMODULE As Long = &H8
'Public Const TH32CS_SNAPALL As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT As Long = &H80000000
    

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type



Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const sLocation As String = "mdlProcess"




Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32, ByVal lProcessID As Long) As Long
On Error GoTo VB_Error
    
    '// I'm tired, just ask me...
    
    ReDim Thread(0)
    
    Dim THREADENTRY32 As THREADENTRY32
    Dim hSnapShot As Long
    Dim lThread As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID) ': 'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot ::: INVALID_HANDLE_VALUE failed", sLocation, "Thread32_Enum")
    
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapShot, THREADENTRY32) = False Then
        Thread32_Enum = -1
         
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    
    Do
        If Thread32Next(hSnapShot, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
    Loop
    
    Thread32_Enum = lThread
    
Exit Function
VB_Error:
   
Resume Next
End Function


Public Sub SuspendThreads(P_ID As Long)
    
    Dim lCount As Long
    Dim i As Long
    
    lCount = Thread32_Enum(Thread(), P_ID)
    
  
        For i = 0 To lCount
            If Thread(i).th32OwnerProcessID = P_ID Then
                Thread_Suspend CLng(Thread(i).th32ThreadID)
            End If
        Next i
End Sub

Public Sub ResumeThreads(P_ID As Long)
    '// A little different
    
    Dim lCount As Long
    Dim i As Long
    
    lCount = Thread32_Enum(Thread(), P_ID)
    
        For i = 0 To lCount
            If Thread(i).th32OwnerProcessID = P_ID Then
                Thread_Resume CLng(Thread(i).th32ThreadID)
            End If
        Next i
   
End Sub


Public Sub Thread_Suspend(T_ID As Long)
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
        'If hThread = 0 Then Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"  'Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
        
        lSuspendCount = SuspendThread(hThread)
        
       
        
    Exit Sub
VB_Error:
   
    Resume Next
End Sub

Public Sub Thread_Resume(T_ID As Long)
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
    
        lSuspendCount = ResumeThread(hThread)
        
       
        
    Exit Sub
VB_Error:
    Resume Next
End Sub





