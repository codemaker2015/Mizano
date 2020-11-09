VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Статистика"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   5430
   HelpContextID   =   11
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5430
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1740
      ScaleHeight     =   375
      ScaleWidth      =   2085
      TabIndex        =   25
      Top             =   630
      Width           =   2145
   End
   Begin VB.ListBox lstBox 
      Height          =   255
      Left            =   1770
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Command3"
      Height          =   975
      Left            =   4830
      TabIndex        =   24
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   210
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lstRunning 
      Height          =   450
      Left            =   180
      TabIndex        =   22
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Timer tmrProcRef 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4170
      Top             =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   4695
      Left            =   150
      TabIndex        =   1
      Top             =   1260
      Width           =   5115
      Begin VB.TextBox lblPathText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   420
         Width           =   4785
      End
      Begin VB.Label lblProcRun 
         BackColor       =   &H80000009&
         Caption         =   "Label8"
         Height          =   405
         Left            =   3030
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "Количество активных процессов"
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   1818
         Width           =   2595
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         BorderWidth     =   2
         X1              =   120
         X2              =   4890
         Y1              =   3870
         Y2              =   3870
      End
      Begin VB.Label lblFound 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   3030
         TabIndex        =   17
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Caption         =   "Количество инфицированных:"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   2256
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Количество заблокированных"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   2694
         Width           =   2355
      End
      Begin VB.Label lblCleaned 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   3030
         TabIndex        =   14
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label lblZabl 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   3030
         TabIndex        =   13
         Top             =   2700
         Width           =   825
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Количество перемещенных"
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   3132
         Width           =   2355
      End
      Begin VB.Label lblVirusCount 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   2610
         TabIndex        =   11
         Top             =   3930
         Width           =   1755
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "Количество вирусов в базе"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   3900
         Width           =   2355
      End
      Begin VB.Label lblLastUpdate 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   2610
         TabIndex        =   9
         Top             =   4260
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "Дата последнего обновления"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   4290
         Width           =   2355
      End
      Begin VB.Label lblCount 
         BackColor       =   &H80000009&
         Caption         =   "Label4"
         Height          =   345
         Left            =   3030
         TabIndex        =   7
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Количество проверенных"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   1380
         Width           =   2355
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Количество удаленных"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   3570
         Width           =   1965
      End
      Begin VB.Label lblDelete 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   345
         Left            =   3030
         TabIndex        =   4
         Top             =   3540
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Последний проверенный"
         Height          =   315
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   2295
      End
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   420
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.ListBox lvProcess 
      Height          =   1230
      Left            =   720
      TabIndex        =   19
      Top             =   1110
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMonme 
      BackColor       =   &H80000009&
      Caption         =   "Монитор выключен"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1530
      TabIndex        =   3
      Top             =   120
      Width           =   2625
   End
   Begin VB.Menu bbb 
      Caption         =   "Файл"
      Begin VB.Menu f 
         Caption         =   "Статистика"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu l3 
         Caption         =   "Включить"
      End
      Begin VB.Menu hm1 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuNast 
         Caption         =   "Настройка"
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuex1 
         Caption         =   "Выгрузить"
      End
   End
   Begin VB.Menu h1 
      Caption         =   "Помощь"
      Begin VB.Menu l2 
         Caption         =   "Справка"
      End
      Begin VB.Menu h51 
         Caption         =   "-"
      End
      Begin VB.Menu h2 
         Caption         =   "Сайт"
      End
      Begin VB.Menu h3 
         Caption         =   "Проверить обновление"
      End
      Begin VB.Menu h4 
         Caption         =   "FAQ"
      End
      Begin VB.Menu h8 
         Caption         =   "-"
      End
      Begin VB.Menu h5 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private programs As New ProcessList
Private Type SYSTEM_INFO
dwOemID As Long
dwPageSize As Long
lpMinimumApplicationAddress As Long
lpMaximumApplicationAddress As Long
dwActiveProcessorMask As Long
dwNumberOrfProcessors As Long
dwProcessorType As Long
dwAllocationGranularity As Long
dwReserved As Long
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Private Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const APP_SYSTRAY_ID = 999 'unique identifier
Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205
Private Const NOTIFYICON_VERSION = &H3
' Used as the ID of the call back message
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
' Constants used to detect clicking on the icon
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
   
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)



'Public WithEvents cTray As mdlMain
'-----------------------
'ProcessList.cls
'version 1.0
'-----------------------
'by: Martin Sykes
'when: 3rd July 2006
'site: www.martrinex.net
'-----------------------
'this code is a part of my two a week summer challenge so check martrinex.net for more codes like this!


'quick example of this module in action, to close notepad and calculator






Public Function CloseEXE(exeName As String)
 '>> our module to list and close windows processes

 
 '>> loop through all the processes what it found
 For i = 0 To processes.processCount - 1
  Debug.Print processes.ProcessName(i)
  '>> see if our exe is running with a quick and crude check
  If InStr(LCase(processes.ProcessName(i)), exeName) <> 0 Then
   '>> close the exe (note it will stay in the list since it is not live)
   processes.KillProcess processes.ProcessHandle(i)
  End If
  '>> find the next exe
 Next
 
 'processes.CheckProcesses '<< we could get a new list of processes after closing everything but we dont need to
 
 '>> we are done shutdown the class module
 Set processes = Nothing
End Function







Private Sub Command1_Click()
    l3_Click
End Sub



 Sub Check_activitiproc()
'при включении сканим все процессы
On Error GoTo 100
frmTest.lblCount.Caption = 0
'>> get new list of whats running
Frame1.Refresh
programs.CheckProcesses
'>> clear our listbox
lstRunning.Clear
List1.Clear
'>> fill our list box
For i = 0 To programs.processCount - 1
frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
If monitorOn = False Then Exit Sub
DoEvents
 lstRunning.AddItem programs.ProcessName(i)
  List1.AddItem programs.ProcessHandle(i)
  If Left(programs.ProcessName(i), 1) <> "\" Then
                '####
                If FileorFolderExists(programs.ProcessName(i)) = False Then
                    'ResumeThreads (hNumrer1)
                    GoTo 2
                End If
      If modScanVirus.ScanFilepR("c:", programs.ProcessName(i), programs.ProcessHandle(i)) = False Then
      Debug.Print "процессы" + programs.ProcessName(i) + "==" + CStr(programs.ProcessHandle(i))
             'ResumeThreads (hNumrer1)
            'Exit Sub
       End If
     
      
                
   End If
2:
Next
lstRunning.Clear
List1.Clear
Exit Sub
100:
MsgBox "" + Error$
End Sub



Private Sub Command3_Click()
lstBox.Clear
FillProcessListNT
'>> clear our listbox
'MsgBox lstBox.ListCount


'>> fill our list box
For i = 1 To lstBox.ListCount - 1
frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
DoEvents

  If Left(lstBox.List(i), 1) <> "\" Then
                '####
                If FileorFolderExists(Trim$(lstBox.List(i))) = False Then
                    'ResumeThreads (hNumrer1)
                    GoTo 2
                End If
      If modScanVirus.ScanDLL("c:", Trim$(lstBox.List(i))) = False Then
      Debug.Print "процессы" + lstBox.List(i)
             'ResumeThreads (hNumrer1)
            'Exit Sub
       End If
     
      
                
   End If
2:
Next i
MsgBox "done"
Exit Sub
100:
MsgBox "" + Error$
End Sub

Private Sub f_Click()
f.Enabled = False
   Me.WindowState = vbNormal
    Me.Show
 
    'ShellTrayRemove
    
End Sub




Sub checkReestr12()
        Dim FirstStart As String
        FirstStart = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPathM")
        'check if it is ever started
        If FirstStart = "" Then
           CreateStringValue &H80000001, "Software\BGAntivirus", 1, "AppPathM", App.Path & "\" & App.exeName
            'default app setting    'default setting cmd
           CreateDwordValue &H80000001, "Software\BGAntivirus", "LogMon", "1"
            CreateDwordValue &H80000001, "Software\BGAntivirus", "LogAppend", Val(True)
            'CreateDwordValue &H80000001, "Software\BGAntivirus", "LogAppend", "1"
            CreateDwordValue &H80000001, "Software\BGAntivirus", "MonQv", "1"
            CreateDwordValue &H80000001, "Software\BGAntivirus", "StartFon", "1"
            CreateDwordValue &H80000001, "Software\BGAntivirus", "Monitoring", "0"
                              
            CreateDwordValue &H80000001, "Software\BGAntivirus", "LogSizeMon", "1"
                            
                            'Me.Text2.Text = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon"))
                            
            'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", 10
            'CreateStringValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
        Else
            'check if previous open is the same as the current
            If FirstStart <> App.Path & "\" & App.exeName Then
                'change key if different
                CreateStringValue &H80000001, "Software\BGAntivirus", 1, "AppPathM", App.Path & "\" & App.exeName
            End If
        End If
End Sub
Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Одная копия программы уже запушена.Если в трее не появится значек, то перезагрузите операционную систему Виндовс", vbCritical, pr
    End
End If
SystemInformation 'проверяем XP это или нет
Call checkReestr12
Dim fn As Long
fn = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartFon"))
If fn = 1 Then
Me.Hide
Else
Me.Show
f.Enabled = False
End If

ShellTrayAdd

'updateRunning_Timer
      frmTest.lblFound.Caption = 0
      frmTest.lblZabl.Caption = 0
      frmTest.lblCleaned.Caption = 0
            frmTest.lblCount.Caption = 0
            frmTest.lblDelete.Caption = 0
     ' Me.lblVirusCount.Caption = VSInfo.VirusCount
    'Me.lblLastUpdate.Caption = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
'   iccex.uCallbackMessage = WM_MOUSEMOVE
'читаем предыдущее состояние мониторинга
   noList = False
    tmrON = False
    keyOn = 0
    monitorOn = False
    firstRun = True
    refProc = True
    unloadOK = False
    logOn = True
    logNew = False
    protectOpt = False
    protectAccess = False
    showGo = False
    hotkeyPrompt = False
    taskmgrFrozen = False
    tempAccPass = False
    protectPass = ""
    prevIndex = 1
       
   Call ReadSig
RefreshDefList

    ReDim procinfo(150) As PROCESSENTRY32
    ReDim jailInfo(1) As jailedProc
    Call enumProc
    firstRun = False
    
Dim mok1 As Boolean
mok1 = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Monitoring"))
If mok1 = True Then
    'включаем мониторинг
    l3_Click
End If

End Sub

Sub RefreshDefList()
    
    Dim i As Long
   ' Me.lvVirusList.ListItems.Clear
    For i = 0 To UBound(VSig)
        'Me.lvVirusList.ListItems.Add , , VSig(i).Name, 4, 4
    Next i
    Me.lblVirusCount.Caption = virCount
    Me.lblLastUpdate.Caption = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
    'Me.SB.Panels(6).Text = VSInfo.VirusCount
    'Me.SB.Panels(5).Text = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Static message As Long
Static RR As Boolean
    
    'x is the current mouse location along the x-axis
    message = x / Screen.TwipsPerPixelX
    
    If RR = False Then
        RR = True
        Select Case message
            ' Left double click (This should bring up a dialog box)
            Case WM_LBUTTONDBLCLK
                Me.Show
            ' Right button up (This should bring up a menu)
            Case WM_RBUTTONUP
                Me.PopupMenu bbb
        End Select
        RR = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   ShellTrayRemove
   End
End Sub

Private Sub h2_Click()
Call ShellExecute(0, "Open", "http://mrbelyash.narod.ru/antivirus/belyashAV.htm", "", "", 1)
End Sub

Private Sub h3_Click()
h3.Enabled = False
Call updateme
h3.Enabled = True
End Sub

Private Sub h5_Click()
frmAbout.Show vbModal
End Sub

Private Sub hm1_Click()
frmAbout.Show vbModal
End Sub

Private Sub l2_Click()
ShowTopicID 1, 11
End Sub

Private Sub l3_Click()
If l3.Caption = "Scanning" Then
    LogPrint "Scanning"
    frmTest.Text1.Text = ""
           Command1.Caption = "Scan"
                   lblMonme.ForeColor = &HFF0000
           'Command1.Refresh
        l3.Caption = "Scanning"
    monitoring = True
        lblMonme.Caption = "Монитор включен"
             monitorOn = True
             Check_activitiproc
             frmTest.Text1.Text = ""
        tmrProcRef.Enabled = True
        Call uc7Interval_Clicked
        tmrON = True

Else
    LogPrint "Scanning"
     l3.Caption = "Scanning"
          lblMonme.ForeColor = &HFF&
          monitoring = False
             lblMonme.Caption = "Монитор выключен"
        Command1.Caption = "Scanning"
        'Command1.Refresh
        tmrProcRef.Enabled = False
        tmrON = False
        frmTest.Text1.Text = ""
End If

End Sub


Sub updateme()
    
End Sub
Private Sub List1_Click()
For i = 1 To List1.ListCount - 1
If List1.Selected(i) Then
    Process_Kill (List1.List(i))
End If
Next
End Sub

Private Sub lstRunning_Click()
Dim i As Integer
Dim z As String
For i = 1 To lstRunning.ListCount
    If lstRunning.Selected(i) Then
    
       ' If MsgBox("Do you want to close this program?" & lstRunning + "-----" + CStr(List1.List(i)), vbYesNo) = vbYes Then
            'programs.KillProcess programs.ProcessHandle(id)
                modScanVirus.ScanFileProcSplash "c:", lstRunning.List(i), List1.List(i)
                        'ResumeThreads (z)
                        Exit Sub
       ' End If
    End If
Next
End Sub

Private Sub mnuex1_Click()
'запоминаем состояние мониторинга для следующего запуска
If monitoring = True Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Monitoring", Val(True)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Monitoring", Val(False)
End If
End
End Sub

Private Sub mnuNast_Click()
mnuNast.Visible = False
nastr.Show vbModal
End Sub

Private Sub updateRunning_Timer()
frmTest.lblCount.Caption = 0
'>> get new list of whats running
Frame1.Refresh
programs.CheckProcesses
'>> clear our listbox
lstRunning.Clear
List1.Clear
'>> fill our list box
For i = 0 To programs.processCount - 1
frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
DoEvents
 lstRunning.AddItem programs.ProcessName(i)
  List1.AddItem programs.ProcessHandle(i)
  If Left(programs.ProcessName(i), 1) <> "\" Then
                '####
                modScanVirus.ScanFileProcSplash "c:", programs.ProcessName(i), programs.ProcessHandle(i) ',list
   End If
Next
End Sub
Private Sub ShellTrayAdd()
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
  'set up the type members
   With nid
   
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = Me.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = frmTest.Icon
      .uCallbackMessage = WM_MOUSEMOVE
      'szTip is the tooltip shown when the
      'mouse hovers over the systray icon.
      'Terminate it since the strings are
      'fixed-length in NOTIFYICONDATA.
      .szTip = "Belyash Shield 2008b" & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      
   End With
   
  'add the icon ...
   Call Shell_NotifyIcon(NIM_ADD, nid)
   
  '... and inform the system of the
  'NOTIFYICON version in use
   Call Shell_NotifyIcon(NIM_SETVERSION, nid)
       
End Sub


Private Sub ShellTrayRemove()

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
      
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = frmTest.hwnd
      .uID = APP_SYSTRAY_ID
   End With
   
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Private Sub ShellTrayModifyTip(nIconIndex As Long)

   Dim nid As NOTIFYICONDATA

   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = frmTest.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      
      'InfoTitle is the balloon tip title,
      'and szInfo is the message displayed.
      'Terminating both with vbNullChar prevents
      'the display of the unused padding in the
      'strings defined as fixed-length in NOTIFYICONDATA.
      .szInfoTitle = "Belyash Shield" & vbNullChar
      .szInfo = Text1.Text & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub


Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function


Private Function GetSelectedOptionIndex() As Long

  'returns the selected item index from
  'an option button array. Use in place
  'of multiple If...Then statements!
  'If your array contains more elements,
  'just append them to the test condition,
  'setting the multiplier to the button's
  'negative -index.
   GetSelectedOptionIndex = 1
End Function
'--end block--'




Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ShellTrayAdd
        Me.Hide
        mnuNast.Visible = True
        'хендл окна
      'Me.hWnd = hWnd
         f.Enabled = True
         hm1.Visible = True
        Exit Sub
            'ShellTrayAdd
      Else
            If Me.Width <> 5565 Then
            Me.Width = 5565
            End If
            If Me.Height <> 6915 Then
                Me.Height = 6915
            End If
     ' frmMain.StartUpPosition = 2
            Dim LeftPos As Integer
            Dim TopPos As Integer
                LeftPos = Int((Screen.Width - Me.Width) / 2)
            TopPos = 0
            Me.Top = TopPos
                Me.Left = LeftPos
                f.Enabled = False
                hm1.Visible = False
            Me.WindowState = vbNormal
            Me.Show
            ShellTrayRemove
    End If

  
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)


'Call ReleaseCapture
'Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'передаем данные в объект
'      cTray.CallEvent X, Y

End Sub

Public Sub message()
ShellTrayModifyTip GetSelectedOptionIndex()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    ShowTopicID 1, 11
End If
If KeyCode = vbKeyF11 Then
    Me.BackColor = &HE0E0E0
    Frame1.BackColor = &HE0E0E0
    
    Dim Ctl As Control
Dim T As Label
For Each Ctl In Me.Controls
  If TypeOf Ctl Is Label Then
    Set T = Ctl
    'Присваиваем значение
    T.BackColor = &HE0E0E0
  End If
Next

End If
End Sub
Sub SystemInformation()
Dim msg As String ' Status information.
Dim NewLine As String ' New-line.
Dim Ret As Integer ' OS Information
Dim ver_major As Integer ' OS Version
Dim ver_minor As Integer ' Minor Os Version
Dim Build As Long ' OS Build
NewLine = Chr(13) + Chr(10) ' New-line.
' Get operating system and version.
Dim verinfo As OSVERSIONINFO
verinfo.dwOSVersionInfoSize = Len(verinfo)
Ret = GetVersionEx(verinfo)
If Ret = 0 Then
MsgBox "Error Getting Version Information"
End
End If


Select Case verinfo.dwPlatformId
Case 2
GoTo 2
Case Else
MsgBox "Программа корректно работает только под Windows XP" + vbCrLf + CStr(verinfo.dwPlatformId), vbExclamation, "Внимание"
End
End Select

2:
End Sub

Private Sub setMonitor()

        If tmrON = True Then
        tmrProcRef.Enabled = False
'        lblMonitor.Caption = "OFF"
'        lblMonitor.ForeColor = &HC0&
'        uc7Monitor.Caption = "Monitor Processes"
        tmrON = False
    Else
        monitorOn = True
        tmrProcRef.Enabled = True
        Call uc7Interval_Clicked
'        lblMonitor.Caption = "ON"
'        lblMonitor.ForeColor = &H8000&
'        uc7Monitor.Caption = "Stop Monitoring"
        tmrON = True
    End If

End Sub

Private Sub uc7Interval_Clicked()
    If txtProcRef = "" Then
        tmrON = True
'        Call setMonitor
    ElseIf Trim(txtProcRef) = 0 Then
        'tmrProcRef.Interval = 250
    Else
        'tmrProcRef.Interval = txtProcRef.Text * 1000
    End If
End Sub


Private Sub tmrProcRef_Timer()
ReDim procinfo(150) As PROCESSENTRY32
    If refProc = True Then
        lvProcess.Clear
        Call enumProc
        refProc = False
    Else
        Call enumProc
    End If
End Sub
