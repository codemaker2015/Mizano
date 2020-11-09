Attribute VB_Name = "ModScanning"
Option Explicit
Dim Seal As New ClsHuffman
Dim Total_size As Double
Public jumlah_file, JumDir As Single
Public monitoring  As Boolean
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Const pr = "ATV Guard Shield"
Private Declare Function FindFirstFile Lib "kernel32" Alias _
"FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias _
"FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As _
WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" _
Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" _
(ByVal hFindFile As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" _
(ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" _
Alias "GetWindowsDirectoryA" _
(ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
'-----------------------------------
Dim VirusDetect As String
Public TipeHeuristic As String
Public VirusDetected As Long
Public VirusCleaned As Long
Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Public Function isProperFile(sPath As String, sExt As String) As Boolean ' Checking Available File / Path
On Error Resume Next

If InStr(1, UCase(sExt), UCase(Right(sPath, 3))) > 0 Then
   isProperFile = True
Else
   isProperFile = False
End If

End Function
Function Scan(Path As String)
    Dim FileName As String
    Dim X As ListItem
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    Dim Algorithm As clsCRC
    Set Algorithm = New clsCRC
    Dim CeksemuaEkstensi As String
    On Error Resume Next
    If frmScanner.Command2.Caption = "Scan" Then Exit Function
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        If (DirName <> ".") And (DirName <> "..") Then
            If GetFileAttributes(Path & DirName) And _
            FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                nDir = nDir + 1
                JumDir = JumDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        DoEvents
        Loop
        Cont = FindClose(hSearch)
    End If
    hSearch = FindFirstFile(Path & "*.*", WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont And frmScanner.Command2.Caption = "Stop"
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                Scan = Scan + (WFD.nFileSizeHigh * MAXDWORD) + _
                WFD.nFileSizeLow
                jumlah_file = jumlah_file + 1
                frmScanner.lblScan.Caption = Path & FileName
                addres = Path & FileName
                 If isProperFile(CeksemuaEkstensi, "SYS LNK VBE HTM HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI") = True Then
                    sCRC = GetChecksum(addres)  ' Checksum .. Action !!
                    cek_virus '------------------> Check With Database
                Else
                End If
                Total_size = Total_size + FileLen(Path & FileName)
                frmScanner.Label59.Caption = jumlah_file & " [ " & JumDir _
                & " ]"
                ' ATV ... --- Action
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
            DoEvents
        Wend
        Cont = FindClose(hSearch)
    End If
    If nDir > 0 Then
        For i = 0 To nDir - 1
            Scan = Scan + Scan(Path & dirNames(i) & "\")
            DoEvents
        Next i
    End If
End Function

Public Function IsFile(Where As String) As Boolean
On Error GoTo FixE
    If FileLen(Where) > 0 Then IsFile = True Else IsFile = False
Exit Function

FixE:
IsFile = False
End Function

Public Function ResidentShield(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long) As Boolean
    On Error Resume Next
    ResidentShield = False
    FrmTest.lblCountVir.Caption = Int(FrmTest.lblCountVir.Caption) + 1
    Dim nama, Exten As String
    Dim strFile As String, strName As String
     Dim fso As New FileSystemObject
    'main folder
    Dim mFolder As folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As folder
    Dim sFile As file
    'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
         'scan virus
        Dim sCRC As String
        sCRC = GetChecksum(mi_file)
        If Left(mi_file, 1) = "/" Then Exit Function
        'MsgBox mi_file
        Dim i As Long
           Debug.Print sCRC
        'compare with database
        For i = 0 To UBound(VSig)
        If monitoring = False Then Exit Function
       DoEvents
       FrmTest.lblPathText.Caption = mi_file
                If sCRC = VSig(i).value Then   'start cleaning
                    FrmTest.lblFoundVir.Caption = Int(FrmTest.lblFoundVir.Caption) + 1
                      FrmTest.lblZabl.Caption = Int(FrmTest.lblZabl.Caption) + 1
                            FrmTest.TxVirus.Text = "Value" + vbCrLf + VSig(i).Name
                            Shield.lblFoundVir.Caption = Int(Shield.lblFoundVir.Caption) + 1
                            Shield.lblPathText.Caption = mi_file
                            Shield.Label37.Caption = " CLEANED VIRUSES + QUARANTINE FILE "
                            Shield.TxVirus.Text = "" + vbCrLf + VSig(i).Name
                            Call Message
                            Process_Kill (proccAidi)
                           Dim fn1 As Long
                           nama = GetFileName(mi_file)
                           Exten = Right$(mi_file, 3)
                           SetFileAttributes nama, FILE_ATTRIBUTE_NORMAL
                           DoEvents
                           TerminateExeName mi_file
                If Seal.EncodeFile(mi_file, App.Path & "\Quarantine\" & nama & "." & Exten & ".atv") = False Then
                    MsgBox "Virus seal infalid !", vbCritical, "ATV Guard"
                End If
                Open (mi_file) For Output As #1
                Close (1)
                Kill (mi_file)
                Call killDelVir(mi_file, VSig(i).Name)
                mi_file = ""
                proccAidi = 0
                ResidentShield = True
                Exit Function
            End If
            Next i
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    FrmTest.TxVirus.Text = ""
End Function
Sub killDelVir(s13 As String, virn2 As String)
On Error Resume Next
  Dim fso As New FileSystemObject, txtfile, fil1, fil2
Set fil1 = fso.GetFile(s13)
    fil1.Delete True
If Dir$(s13, vbNormal) = "" Then
        FrmTest.lblDeleteVir.Caption = Int(FrmTest.lblDeleteVir.Caption) + 1
        FrmTest.TxVirus.Text = "Agung <" + virn2 + ">  Agung"
        Call Message
    Else
            FrmTest.TxVirus.Text = "Program" + vbCrLf + virn2
        Call Message
    End If
Set fso = Nothing
    Set fil1 = Nothing
End Sub
Function WinDir() As String
    Dim sSave As String, ret As Long
    sSave = Space(255)
    ret = GetWindowsDirectory(sSave, 255)
    WinDir = Left$(sSave, ret)
End Function

Public Sub cek_virus()
Static Num As Integer
Static X As ListItem
Static V_name As String
On Error Resume Next
Dim A As Long
For A = 0 To UBound(VSig)
If GetInputState() <> 0 Then DoEvents
If sCRC = VSig(A).value Then   'start cleaning
Beep 800, 80: Beep 500, 80
Set X = frmScanner.ListView1.ListItems.Add(, , VSig(A).Name, 3, 3)
X.SubItems(1) = addres
X.SubItems(2) = "Detected"
X.SubItems(3) = Format(Date, "ddd, dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS")
X.SubItems(4) = Int((FileLen(addres) / 1024) * 100 + 0.5) / 100 & " KB" 'FileLen(strFile) & " Bytes"
End If
Next A
End Sub

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

Sub LogQ(sMessage6 As String)
Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
ffile = App.Path + "\Quarantine\removed.log"
If Dir$(App.Path + "\Quarantine", vbDirectory) = "" Then
  MkDir App.Path + "\Quarantine"
End If
If Dir$(ffile, vbNormal) <> "" Then
    If FileLen(ffile) >= 3145728 Then
        Kill ffile
    End If
End If
End Sub

