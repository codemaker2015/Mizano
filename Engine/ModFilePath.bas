Attribute VB_Name = "ModFilePath"
Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Public Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Public Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' --------------------------------------
Public Declare Function GetUSERNAME Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
' Returns version information about a specified file ------------------------------
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal X&, ByVal y&, ByVal FLAGS&) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
' Sets a file’s attributes -------------------------
' mdlMemory ---------------------------------
Public Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Type VERHEADER
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    Comments As String
    LegalTradeMarks As String
    PrivateBuild As String
    SpecialBuild As String
End Type

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 36
    szTypeName As String * 80
End Type


Public Const WTS_CURRENT_SERVER_HANDLE = 0&
Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Public Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
End Type
Public PCName As String
Public Enum IconRetrieve
    ricnLarge = 32
    ricnSmall = 16
End Enum


Private Function GetShortPath(strFileName As String) As String
    Dim lngRes&, strPath$: strPath = String$(64, 0)
    lngRes = GetShortPathName(strFileName, strPath, 64)
    GetShortPath = Left$(strPath, lngRes)
End Function
Public Function GetVerHeader(ByVal fPN$, ByRef oFP As VERHEADER)
On Error Resume Next
Dim lngBufferlen&, lngDummy&, lngRc&, lngVerPointer&, lngHexNumber&, i%
Dim bytBuffer() As Byte, bytBuff(255) As Byte, strBuffer$, strLangCharset$, strVersionInfo(11) As String, strTemp$
 If Dir(fPN$, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = "" Then
    oFP.CompanyName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.FileDescription = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.FileVersion = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.InternalName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.LegalCopyright = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.OrigionalFileName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.ProductName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.ProductVersion = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.Comments = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.LegalTradeMarks = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.PrivateBuild = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.SpecialBuild = "The file """ & GetShortPath(fPN) & """ N/A"
    Exit Function
 End If
   lngBufferlen = GetFileVersionInfoSize(fPN$, lngDummy)
    If lngBufferlen > 0 Then
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(fPN$, 0&, lngBufferlen, bytBuffer(0))
       If lngRc <> 0 Then
        lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
         If lngRc <> 0 Then
          MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
           lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
            strLangCharset = Hex(lngHexNumber)
             Do While Len(strLangCharset) < 8
              strLangCharset = "0" & strLangCharset
             Loop
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             strVersionInfo(8) = "Comments"
             strVersionInfo(9) = "LegalTrademarks"
             strVersionInfo(10) = "PrivateBuild"
             strVersionInfo(11) = "SpecialBuild"
            For i = 0 To 11
               strBuffer = String$(255, 0)
               strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(i)
               lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   lstrcpy strBuffer, lngVerPointer
                   strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                   strVersionInfo(i) = strBuffer
                Else
                  strVersionInfo(i) = "N/A"
                   End If
            Next i
          End If
       End If
    End If
     For i = 0 To 11
      If Trim(strVersionInfo(i)) = "" Then strVersionInfo(i) = "N/A"
     Next i
    oFP.CompanyName = strVersionInfo(0)
    oFP.FileDescription = strVersionInfo(1)
    oFP.FileVersion = strVersionInfo(2)
    oFP.InternalName = strVersionInfo(3)
    oFP.LegalCopyright = strVersionInfo(4)
    oFP.OrigionalFileName = strVersionInfo(5)
    oFP.ProductName = strVersionInfo(6)
    oFP.ProductVersion = strVersionInfo(7)
    oFP.Comments = strVersionInfo(8)
    oFP.LegalTradeMarks = strVersionInfo(9)
    oFP.PrivateBuild = strVersionInfo(10)
    oFP.SpecialBuild = strVersionInfo(11)
End Function


Public Function GetMemory(ProcessID As Long) As String
    
    On Error Resume Next
    Dim byteSize As Double, hProcess As Long, ProcMem As PROCESS_MEMORY_COUNTERS
    
    ProcMem.cb = LenB(ProcMem)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    
    If hProcess <= 0 Then GetMemory = "N/A": Exit Function
    
    GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
    
    byteSize = ProcMem.WorkingSetSize
    GetMemory = byteSize
    
    Call CloseHandle(hProcess)
    
End Function
Function GetAttribute(ByVal sFilePath As String) As String
        
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "R": Case 2: GetAttribute _
            = "H": Case 3: GetAttribute = "RH": Case 4: _
            GetAttribute = "S": Case 5: GetAttribute = _
            "RS": Case 6: GetAttribute = "HS": Case 7: _
            GetAttribute = "RHS"
        '-------------------------------------------------'
        Case 32: GetAttribute = "A": Case 33: GetAttribute _
            = "RA": Case 34: GetAttribute = "HA": Case 35: _
            GetAttribute = "RHA": Case 36: GetAttribute = _
            "SA": Case 37: GetAttribute = "RSA": Case 38: _
            GetAttribute = "HSA": Case 39: GetAttribute = _
            "RHSA"
        '-------------------------------------------------'
        Case 128: GetAttribute = "Normal"
        '-------------------------------------------------'
        Case Else: GetAttribute = "N/A"
    End Select

End Function
Public Function GetPriority(PID As Long)
Dim hwnd As Long, pri As Long
    hwnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, PID)
    pri = GetPriorityClass(hwnd)
    CloseHandle hwnd
    GetPriority = pri
End Function

