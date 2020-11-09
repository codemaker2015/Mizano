Attribute VB_Name = "ModPublic"
Public Hasil, dataX, addres As String
Public CRC2 As String
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetUSERNAME Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
' Registry access
' mdlRegistry -----------------------------------
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public PopupReason As String
Public ShowPopup As Boolean
Public nRows As Long
Public UserCom As String
Public ComName As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const SPI_SETNONCLIENTMETRICS As Long = &H2A
Public Const SMTO_ABORTIFHUNG As Long = &H2
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)
Public Type SECURITY_ATTRIBUTES
   nLength                 As Long
   lpSecurityDescriptor    As Long
   bInheritHandle          As Long
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
    
Public Const ERROR_SUCCESS = 0&

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000

Public Enum TipeReg
    REG_SZ = 1                         ' Unicode nul terminated string
    reg_binarY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
End Enum
Public Const vAppVersion = "ATV Guard V.2"
Public Sub LogFile(strLog As String)
On Error Resume Next

Dim ff As Integer
ff = FreeFile

MkDir App.path & "\Log"
Open App.path & "\Log\" & "ATV_log" & ".txt" For Append As #ff
Print #ff, Date & " " & " " & " " & " " & " " & " " & Time & _
                " " & " " & " " & " " & " " & " " & strLog
Close #ff

End Sub


Public Function FileParsePath(sPathname As String, bRetFile As Boolean, bExtension As Boolean) As String
    Dim sEditArray() As String
    sEditArray = Split(sPathname, "\", -1)
    If bRetFile = True Then
        Dim sFilename As String
        sFilename = sEditArray(UBound(sEditArray))
        If bExtension = True Then
            FileParsePath = sFilename
        Else
            sEditArray = Split(sFilename, ".vir", -1)
            FileParsePath = sEditArray(LBound(sEditArray))
        End If
    Else
        Dim sPathnameA As String
        Dim i As Integer
        For i = 0 To UBound(sEditArray) - 1
            sPathnameA = sPathnameA & sEditArray(i) & "\"
        Next
        FileParsePath = sPathnameA
    End If
    On Error GoTo 0
End Function
Public Function GetEntryCount() As Long
    GetEntryCount = nRows - 2 '// The last entry is always an EOF of sorts
End Function
Public Function NameOfTheComputer(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function

Public Function GetUserCom() As String

    GetUserCom = Environ$("username")
    ComName = NameOfTheComputer(PCName)

    StatusRegister.Label3.Caption = "User Name : " + GetUserCom
    StatusRegister.Label4.Caption = "PC Name   : " + PCName
    
End Function
Public Function HitDatabase()
    Dim i As Integer
    Dim vCount As Integer
    
    frmDatabase.lstVirus.ListItems.Clear
    vCount = 0
    For i = 0 To UBound(VSig)
        vCount = vCount + 1
        frmUpdate.lstVirus.ListItems.Add , , VSig(i).Name, , 1
    Next i
    frmUpdate.lblVirusCount.Caption = VSInfo.VirusCount
    frmUpdate.lblLastUpdate.Caption = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
    frmScanner.lblFile.Caption = VSInfo.VirusCount
    'frmUpdate.Label48.Caption = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
End Function
Function ReplacePathSystem(np As String) As String
On Error Resume Next
Dim buff As String
buff = Replace(np, "\??\", "", , , vbTextCompare)
buff = Replace(buff, "\\?\", "", , , vbTextCompare)
buff = Replace(buff, "\SystemRoot\", MyWindowDir, , , vbTextCompare)
buff = Replace(buff, "%systemroot%", MyWindowDir, , , vbTextCompare)
buff = Replace(buff, "\\", "\", , , vbTextCompare)
ReplacePathSystem = buff
End Function
Public Function Dire(str_File As String) As Boolean
    On Error GoTo err
    Dire = Not (Dir(str_File) = "" And Dir(str_File, vbHidden) = "" And Dir(str_File, vbSystem) = "" And Dir(str_File, vbNormal) = "")
Exit Function
err:
    Dire = False
    err.Clear
End Function
Sub GetWTSProcesses(coll As Collection)
On Error Resume Next
Dim Retval As Long
Dim count As Long
Dim i As Integer
Dim lpBuffer As Long
Dim P As Long
Dim udtProcessInfo As WTS_PROCESS_INFO

If IsWinNT Then
    Retval = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, count)
    If Retval Then
       P = lpBuffer
         For i = 1 To count
             CopyMemory udtProcessInfo, ByVal P, LenB(udtProcessInfo)
             coll.Add GetUserNameA(udtProcessInfo.pUserSid), "#" & udtProcessInfo.ProcessID
             P = P + LenB(udtProcessInfo)
         Next i
         WTSFreeMemory lpBuffer   'Free your memory buffer
     End If
End If
End Sub

Public Sub AlwaysOnTop(hwnd As Long, SetOnTop As Boolean)
    If SetOnTop Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPFLAGS
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPFLAGS
    End If
End Sub
Function GetUserNameA(sID As Long) As String
If IsWinNT Then

    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    GetUserNameA = Left$(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left$(retname, InStr(retname, vbNullChar) - 1)
End If
End Function
