Attribute VB_Name = "ModProcess"
Public Declare Function EnumProcesses Lib "psapi.dll" _
   (ByRef lpidProcess As Long, ByVal cb As Long, _
      ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Const PROCESS_TERMINATE As Long = (&H1)
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long

Public Type jailedProc
    jailPID As Long
    ExeName As String
    attempts As Integer
    prevAction As String
    firstTime As String
    dateOf As String
    lastTime As String
    onNow As Boolean
    attemptTimes() As String
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    childWnd As Integer
    procName As String
End Type

Public Const PROCESS_QUERY_INFORMATION = &H400

Public procinfo() As PROCESSENTRY32
Public arrLen As Integer
Public noList As Boolean
Public tmrON As Boolean
Public runningProc As Integer
Public monitorOn As Boolean
Public jailInfo() As jailedProc
'Public colHead As ColumnHeader
'Public lstItem As ListItem
Public tempArr1() As String
Public tempArr2() As String
Public tempArr3() As String
Public tempArr4() As String
Public copyArr() As Integer
Public firstRun As Boolean
Public glbPID As Long
Public frmIndex As Integer
Public frm As Form
Public refProc As Boolean
Public skipProc As Integer
Public unloadOK As Boolean
Public logOn As Boolean
Public protectPass As String
Public protectOpt As Boolean
Public protectAccess As Boolean
Public protectLogs As Boolean
Public protectInfo As Boolean
Public prevIndex As Integer
Public prevCapt As String
Public showGo As Boolean
Public taskmgrFrozen As Boolean
Public hotkeyPrompt As Boolean
Public tempAccPass As Boolean
Public pkResult As Long
Public optString As String
Public logNew As Boolean
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" _
   (ByVal hProcess As Long, ByVal hModule As Long, _
      ByVal moduleName As String, ByVal nSize As Long) As Long

Public Declare Function EnumProcessModules Lib "psapi.dll" _
   (ByVal hProcess As Long, ByRef lphModule As Long, _
      ByVal cb As Long, ByRef cbNeeded As Long) As Long


Public Const PROCESS_VM_READ = 16
'Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Sub enumProc()
Dim found As Integer
Dim inList As Boolean
    inList = False
    arrLen = 0
    runningProc = 0
    skipProc = 0
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    r = Process32Next(hSnapShot, uProcess)
    Do While r
        runningProc = runningProc + 1
        
        ReDim Preserve tempArr1(runningProc)
        ProcessName = Left$(uProcess.szexeFile, IIf(InStr(1, uProcess.szexeFile, Chr$(0)) > 0, InStr(1, uProcess.szexeFile, Chr$(0)) - 1, 0))
        tempArr1(runningProc) = ProcessName
        If noList = False Then
            If refProc = True Then
                FrmTest.lvProcess.AddItem ProcessName ' & "=" & uProcess.th32ProcessID

            End If
        End If
        uProcess.procName = ProcessName
        For i = 0 To 150
            If procinfo(i).th32ProcessID = 0 Then
                arrLen = arrLen + 1
                procinfo(i) = uProcess
                procinfo(i).childWnd = 0
                Exit For
            Else
                If i = 150 Then
                    MsgBox "System Full"
                    Exit For
                End If
            End If
        Next i
        r = Process32Next(hSnapShot, uProcess)
    Loop
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    If firstRun = True Then
        ReDim tempArr2(UBound(tempArr1))
        tempArr2 = tempArr1
    Else
        If monitorOn = True Then
        '--------------------------------Check for added----------------------------------
                ReDim copyArr(UBound(tempArr1))
                ReDim tempArr3(UBound(tempArr2))
                tempArr3 = tempArr2
                For i = 1 To UBound(tempArr1)
                    For Z = 1 To UBound(tempArr3)
                        If UCase(tempArr1(i)) = UCase(tempArr3(Z)) Then
                            tempArr3(Z) = ""
                            copyArr(i) = 1
                            Exit For
                        End If
                    Next Z
                Next i
                Call newProcesses
        '----------------------------Check for deleted--------------------------------------
               ReDim copyArr(UBound(tempArr2))
               ReDim tempArr4(UBound(tempArr2))
                For i = 1 To UBound(tempArr2)
                    For Z = 1 To UBound(tempArr1)
                        If UCase(tempArr2(i)) = UCase(tempArr1(Z)) Then
                            tempArr4(Z) = ""
                            copyArr(i) = 1
                            Exit For
                        End If
                    Next Z
                Next i
               Call cleanupProcesses
        '------------------------------------------------------------------
            End If
        End If
    ReDim tempArr2(UBound(tempArr1))
    tempArr2 = tempArr1
    CloseHandle hSnapShot
    FrmTest.lblProcRun.Caption = runningProc
End Sub

Public Sub newProcesses()
Dim newProc As String
    For i = 1 To UBound(copyArr)
        If copyArr(i) = 0 Then
            newProc = tempArr1(i)
            If InStr(1, newProc, "svchost.exe") > 0 Then
            'MsgBox ""
            Else
                
                refProc = True
                For Z = 0 To UBound(jailInfo)
                    If UCase(newProc) = UCase(jailInfo(Z).ExeName) Then
                        jailInfo(Z).lastTime = Time
                        jailInfo(Z).onNow = True
                      
                        Exit For
                    End If
                Next Z
                skipProc = checkForDouble(newProc)
                frmIndex = findFile(newProc)
                If newProc = "taskmgr.exe" Then
                    taskmgrFrozen = True
                End If
                glbPID = procinfo(frmIndex).th32ProcessID
             
                SuspendThreads (procinfo(frmIndex).th32ProcessID)
                DoEvents
                EnumWindows AddressOf EnumWindowsProc, ByVal 0&
                DoEvents
                
                        Call checkthis(procinfo(frmIndex).procName, procinfo(frmIndex).th32ProcessID)
                                               
                    Exit Sub
              
            End If
        End If
    Next i
End Sub
Sub checkthis(A As String, hNumrer1 As Long)
'frmTest.lstBox.Clear
FillProcessListNT
For i = 1 To FrmTest.lstBox.ListCount - 1
d = InStr(1, Trim$(FrmTest.lstBox.List(i)), Trim$(A), vbTextCompare)
    If d <> 0 Then
       ' MsgBox frmTest.lstBox.List(i)
        If FileorFolderExists(FrmTest.lstBox.List(i)) = False Then
            ResumeThreads (hNumrer1)
            Exit Sub
        End If
        
      If ResidentShield("c:", FrmTest.lstBox.List(i), hNumrer1) = False Then
             ResumeThreads (hNumrer1)
            Exit Sub
        End If
    End If
Next i
End Sub
Public Sub cleanupProcesses()
Dim delProc As String
    For i = 1 To UBound(copyArr)
        If copyArr(i) = 0 Then
            delProc = tempArr2(i)
            If InStr(1, delProc, "svchost.exe") > 0 Then
            Else
                FillProcessListNT
                refProc = True
                For Z = 0 To UBound(jailInfo)
                    If UCase(delProc) = UCase(jailInfo(Z).ExeName) Then
                        jailInfo(Z).onNow = False
                        Exit For
                    End If
                Next Z
                
            End If
        End If
    Next i
End Sub

Public Function findFile(fName As String) As Integer
Dim counter As Integer
    counter = 0
    For i = 1 To UBound(procinfo)
        If fName = procinfo(i).procName Then
            If counter = skipProc Then
                findFile = i
                Exit For
            Else
                counter = counter + 1
            End If
        End If
    Next i
End Function



Public Function checkForDouble(prName As String) As Integer
Dim doubles As Integer
    doubles = 0
    For i = 0 To UBound(procinfo)
        If UCase(prName) = UCase(procinfo(i).procName) Then
            doubles = doubles + 1
        End If
    Next i
    checkForDouble = doubles - 1
End Function

Public Function FillProcessListNT() As Long
'
' Clears the listbox and fill it with the
' processes and the modules used by each process.
'
Dim cb                As Long
Dim cbNeeded          As Long
Dim NumElements       As Long
Dim ProcessIDs()      As Long
Dim cbNeeded2         As Long
Dim NumElements2      As Long
Dim ModuleS(1 To 200) As Long
Dim lRet              As Long
Dim moduleName        As String
Dim nSize             As Long
Dim hProcess          As Long
Dim i                 As Long
Dim sModName          As String
Dim sChildModName     As String
Dim iModDlls          As Long
Dim iProcesses        As Integer
    
FrmTest.lstBox.Clear
'
' Get the array containing the process id's for each process object.
'
cb = 8
cbNeeded = 96
'
' There is no way to find out how big the passed in array
' must be. EnumProcesses() will never return a value in
' cbNeeded that is larger than the size of array value
' that you passed in the cb parameter.
'
' If cbNeeded == cb upon return, allocate a larger array
' and try again until cbNeeded is smaller than cb.
'
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)

Loop
'
' Calculate how many process IDs were returned.
'
NumElements = cbNeeded / 4
    
For i = 1 To NumElements
    '
    ' Get a handle to the Process.
    '
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
            Or PROCESS_VM_READ, 0, ProcessIDs(i))
    '
    ' Iterate through each process with an ID that <> 0.
    '
    If hProcess Then
        '
        ' Get an array of the module handles for the specified process.
        '
        lRet = EnumProcessModules(hProcess, ModuleS(1), 200, cbNeeded2)
      'frmTest.lstProcid1.AddItem lRet
        '
        ' If the Module Array is retrieved, Get the ModuleFileName.
        '
        If lRet <> 0 Then
            '
            ' Fill the ModuleName buffer with spaces.
            '
            moduleName = Space(MAX_PATH)
            '
            ' Preset buffer size.
            '
            nSize = 500
            '
            ' Get the module file name.
            '
            lRet = GetModuleFileNameExA(hProcess, ModuleS(1), moduleName, nSize)
            
            '
            ' Get the module file name out of the buffer, lRet is how
            ' many characters the string is, the rest of the buffer is spaces.
            '
            sModName = Left$(moduleName, lRet)
            '
            ' Add the process to the listbox.
            '
            FrmTest.lstBox.AddItem sModName
          
            '
            ' Increment the count of processes we've added.
            '
            iProcesses = iProcesses + 1
                
            iModDlls = 1
            Do
                iModDlls = iModDlls + 1
                '
                ' Fill the ModuleName buffer with spaces.
                '
                moduleName = Space(MAX_PATH)
                '
                ' Preset buffer size.
                '
                nSize = 500
                '
                ' Get the module file name out of the buffer, lRet is how
                ' many characters the string is, the rest of the buffer is spaces.
                '
                lRet = GetModuleFileNameExA(hProcess, ModuleS(iModDlls), moduleName, nSize)
                sChildModName = Left$(moduleName, lRet)
                    'frmTest.lstProcid1.AddItem lRet
                If sChildModName = sModName Then Exit Do
                If Trim(sChildModName) <> "" Then
                If ChekLstExist(Trim(sChildModName)) = False Then
                           
                        FrmTest.lstBox.AddItem "    " & sChildModName
                    
                End If
                'frmTest.lstProcid1.AddItem hProcess
                End If
            Loop
        End If
    Else
        '
        ' Return the number of Processes found.
        '
        FillProcessListNT = 0
    End If
    '
    ' Close the handle to the process.
    '
    lRet = CloseHandle(hProcess)
Next
'
' Return the number of Processes found.
'
'MsgBox iModDlls
FillProcessListNT = iProcesses
End Function
Function ChekLstExist(q1 As String) As Boolean
For i = 1 To FrmTest.lstBox.ListCount - 1
If Trim$(FrmTest.lstBox.List(i)) = q1 Then
    ChekLstExist = True
    Exit Function
End If
        ChekLstExist = False
Next i
End Function








