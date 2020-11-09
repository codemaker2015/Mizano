Attribute VB_Name = "ModQuarantine"
' Sets a file’s attributes -------------------------
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As FILE_ATTRIBUTE) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Enum FILE_ATTRIBUTE
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_ATTRIBUTE_COMPRESSED = &H800
End Enum
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function DeleteIt(whereit As String)

SetFileAttributes whereit, FILE_ATTRIBUTE_NORMAL
TerminateExeName GetFileName(whereit)
Kill (whereit)
If IsFileExist(whereit) = True Then
    Call MsgBox("File can't be deleted!", vbCritical + vbOKOnly, "Error Detected !")
End If
End Function
Private Function IsFileExist(ByVal sPath As String) As Boolean
    
If PathFileExists(sPath) = 1 And PathIsDirectory(sPath) = 0 Then
    IsFileExist = True
Else
    IsFileExist = False
End If
    
End Function
Function TerminateExeName(ExeName As String) ' As Long
On Error GoTo ErrHandle
    
    Dim uProcess As PROCESSENTRY32
    Dim lProc As Long, hProcSnap As Long
    Dim ExePath As String
    Dim hPID As Long, hExit As Long
    Dim i As Integer

    uProcess.dwSize = Len(uProcess)
    hProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    lProc = Process32First(hProcSnap, uProcess)
    Do While lProc
        i = InStr(1, uProcess.szexeFile, Chr$(0))
        ExePath = UCase$(Left$(uProcess.szexeFile, i - 1))
        If UCase$(GetFileName(ExePath)) = UCase$(ExeName) Then
            hPID = OpenProcess(1&, -1&, uProcess.th32ProcessID)
            hExit = TerminateProcess(hPID, 0&)
            Call CloseHandle(hPID)
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
    Exit Function
    
ErrHandle:
End Function

Public Function GetFileName(sFilename As String) As String

Dim buffer As String

buffer = String(255, 0)
GetFileTitle sFilename, buffer, Len(buffer)
buffer = StripNulls(buffer)
GetFileName = buffer
    
End Function
