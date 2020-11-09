VERSION 5.00
Begin VB.Form FrmTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Resident Shield"
   ClientHeight    =   6375
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Guard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   0
      Picture         =   "FrmTest.frx":0CCA
      ScaleHeight     =   6450
      ScaleWidth      =   8505
      TabIndex        =   1
      Top             =   0
      Width           =   8505
      Begin VB.Timer tmTgu 
         Interval        =   1500
         Left            =   7800
         Top             =   3240
      End
      Begin Mizano.Abutton Command4 
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   3840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BorderColor     =   -2147483627
         BorderColorPressed=   -2147483628
         BorderColorHover=   -2147483627
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   5595
         Begin VB.Line Line8 
            BorderColor     =   &H8000000B&
            BorderWidth     =   2
            X1              =   120
            X2              =   4890
            Y1              =   2790
            Y2              =   2790
         End
         Begin VB.Label lblDeleteVir 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3030
            TabIndex        =   25
            Top             =   2460
            Width           =   825
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Size [ Byte ]"
            Height          =   255
            Left            =   150
            TabIndex        =   24
            Top             =   2490
            Width           =   1965
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Curent Scanning Status"
            Height          =   255
            Left            =   150
            TabIndex        =   23
            Top             =   300
            Width           =   2355
         End
         Begin VB.Label lblCountVir 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3030
            TabIndex        =   22
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Update"
            Height          =   255
            Left            =   150
            TabIndex        =   21
            Top             =   3210
            Width           =   2355
         End
         Begin VB.Label Label32 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   2520
            TabIndex        =   20
            Top             =   3240
            Width           =   2325
         End
         Begin VB.Label Label33 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Signature Count"
            Height          =   255
            Left            =   150
            TabIndex        =   19
            Top             =   2820
            Width           =   2355
         End
         Begin VB.Label Label34 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   2490
            TabIndex        =   18
            Top             =   2850
            Width           =   2475
         End
         Begin VB.Label Label35 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quarantine"
            Height          =   255
            Left            =   150
            TabIndex        =   17
            Top             =   2055
            Width           =   2355
         End
         Begin VB.Label lblZabl 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3030
            TabIndex        =   16
            Top             =   1620
            Width           =   825
         End
         Begin VB.Label lblCleanedVir 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3030
            TabIndex        =   15
            Top             =   2040
            Width           =   825
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Move"
            Height          =   255
            Left            =   150
            TabIndex        =   14
            Top             =   1620
            Width           =   2355
         End
         Begin VB.Label Label37 
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            Height          =   255
            Left            =   150
            TabIndex        =   13
            Top             =   1170
            Width           =   2355
         End
         Begin VB.Label lblFoundVir 
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3030
            TabIndex        =   12
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label38 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Runing Status"
            Height          =   255
            Left            =   150
            TabIndex        =   11
            Top             =   735
            Width           =   2595
         End
         Begin VB.Label lblProcRun 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   3030
            TabIndex        =   10
            Top             =   720
            Width           =   975
         End
      End
      Begin Mizano.Abutton Abutton1 
         Height          =   375
         Left            =   7200
         TabIndex        =   30
         ToolTipText     =   "Exit Resident Shield..."
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BorderColor     =   -2147483627
         BorderColorPressed=   -2147483628
         BorderColorHover=   -2147483627
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lvProcess 
         Height          =   1230
         Left            =   1620
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer tmrProcRef 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3120
         Top             =   2040
      End
      Begin VB.ListBox lstRunning 
         Height          =   450
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ListBox List1 
         Height          =   450
         Left            =   1110
         TabIndex        =   6
         Top             =   2070
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000009&
         Caption         =   "Command3"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.ListBox lstBox 
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox TxVirus 
         Height          =   345
         Left            =   960
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3885
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "2.0.1.0"
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "FDD Version        "
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Version     :       V.2"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield The Component  is an antivirus Monitor Which Resides in the memory Of computer Scan Files in Real Time Protection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblPathText 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   240
         TabIndex        =   28
         Top             =   5400
         Width           =   8055
      End
      Begin VB.Label lblMonme 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Left            =   2520
         TabIndex        =   29
         Top             =   4800
         Width           =   2625
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Real Time Protection Status"
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Image Image11 
         Height          =   810
         Left            =   7200
         Picture         =   "FrmTest.frx":7BE2
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   26
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00E0E0E0&
      Caption         =   "UnLoad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MnStart 
         Caption         =   "Start"
      End
      Begin VB.Menu MnUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu MnAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MnEprogram 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu bbb 
      Caption         =   "Файл"
      Visible         =   0   'False
      Begin VB.Menu l3 
         Caption         =   "Включить"
      End
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim delCancel As Boolean
Dim msg As String
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
Dim REG As New cRegistry
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim Pesan As String
Dim sTitle As String

Private Sub Abutton1_Click()
 Dim i As Integer
 For i = 1500 To 2000 Step 100
        Beep i, 20
    Next i
Call SysTrayFunc
End Sub

Private Sub cmdRemove_Click()

    Call Shell_NotifyIcon(NIM_DELETE, ARV)
    
End Sub
Private Sub uc7Interval_Clicked()
tmrON = True
End Sub
Sub CopyPlugin()
    On Error Resume Next
    Dim h As String
    h = Dir(nPath(MyWindowSys) & "comctl32.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If h = "" Then
        FileCopy nPath(App.Path) & "\Data\comctl32.ocx", nPath(MyWindowSys) & "comctl32.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\comctl32.ocx", 0
    End If
    h = Dir(nPath(MyWindowSys) & "mscomct2.ocx", vbArchive + vbHidden + vbSystem + vbNormal + vbReadOnly)
    If h = "" Then
        FileCopy nPath(App.Path) & "\Data\mscomct2.ocx", nPath(MyWindowSys) & "mscomct2.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\mscomct2.ocx", 0
    End If
    h = Dir(nPath(MyWindowSys) & "comdlg32.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If h = "" Then
        FileCopy nPath(App.Path) & "\Data\comdlg32.ocx", nPath(MyWindowSys) & "comdlg32.ocx"
'        Shell "regsvr32 /s" & App.path & "\Data\comdlg32.ocx", 0
    End If
    h = Dir(nPath(MyWindowSys) & "mscomctl.ocx", vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
    If h = "" Then
        FileCopy nPath(App.Path) & "\Data\mscomctl.ocx", nPath(MyWindowSys) & "mscomctl.ocx"
'        Shell "regsvr32 /s" & nPath(MyWindowSys) & "\comdlg32.ocx", 0
    End If
End Sub


Private Sub Command4_Click()
l3_Click
End Sub

Private Sub Form_Load()
Call SysTrayFunc
Me.lblFoundVir.Caption = 0
      Me.lblZabl.Caption = 0
      Me.lblCleanedVir.Caption = 0
      Me.lblCountVir.Caption = 0
      Me.lblDeleteVir.Caption = 0
   noList = False
    tmrON = False
    'keyOn = 0
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
    Me.Show
FileSize = 524288
ReDim procinfo(150) As PROCESSENTRY32
    ReDim jailInfo(1) As jailedProc
    Call enumProc
    firstRun = False
    l3_Click
    AlwaysOnTop Me.hwnd, True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX

    Select Case lMsg
      Case WM_LBUTTONUP
      Case WM_RBUTTONUP
            FrmTest.PopupMenu FrmTest.Menu
      Case WM_MOUSEMOVE
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONDBLCLK
            frmHome.Show
      Case WM_RBUTTONDOWN
      Case WM_RBUTTONDBLCLK
      Case Else
    End Select
End Sub
Private Sub l3_Click()
If l3.Caption = "Включить" Then
    FrmTest.TxVirus.Text = ""
           Command4.Caption = "Stop RTP"
           lblMonme.ForeColor = &HFF0000
           'Command1.Refresh
        l3.Caption = "Выключить"
    monitoring = True
    
        lblMonme.Caption = "Shield Activity"
             monitorOn = True
             Check_activitiproc
             FrmTest.TxVirus.Text = ""
PesanARV.Image2.Visible = False
PesanARV.Show
PesanARV.Label1.Caption = "Shield Activity"
PesanARV.Picture2.Visible = True
PesanARV.Label2.Caption = "Resident Shield Activity. ATV Guard will protect your computer from viruses. "
        tmrProcRef.Enabled = True
        Call uc7Interval_Clicked
        tmrON = True

Else
    
     l3.Caption = "Включить"
          lblMonme.ForeColor = &HFF&
          monitoring = False
        lblMonme.Caption = "Shield not activity"
        Command4.Caption = "Start RTP"
PesanARV.Show
PesanARV.Label1.Caption = "Warning !!"
PesanARV.Label2.Caption = "Shield Not Activity. Your computer may be at risk! "
        tmrProcRef.Enabled = False
        tmrON = False
        FrmTest.TxVirus.Text = ""
End If

End Sub
 Sub Check_activitiproc()
Dim i As Long
On Error GoTo 100
Me.lblCountVir.Caption = 0
'>> get new list of whats running
Frame1.Refresh
programs.CheckProcesses
'>> clear our listbox
lstRunning.Clear
List1.Clear
'>> fill our list box
For i = 0 To programs.processCount - 1
Me.lblCountVir.Caption = Int(Me.lblCountVir.Caption) + 1
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
      If ModScanning.ResidentShield("c:", programs.ProcessName(i), programs.ProcessHandle(i)) = False Then
      Debug.Print "Agung" + programs.ProcessName(i) + "==" + CStr(programs.ProcessHandle(i))
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

Private Sub MnAbout_Click()
About.Show
End Sub

Private Sub MnEprogram_Click()
End
End Sub

Private Sub MnStart_Click()
frmHome.Show
End Sub

Private Sub MnUpdate_Click()
MsgBox "Visited Http://rexsonic-technologie.webs.com", vbInformation, "ATV Guard"
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

Private Function IsOutOfdate()
Dim Thn As Integer
Dim Bln As Byte
Thn = Year(Now)
Bln = Month(Now)
'If Thn = 2010 And Bln <= 11 Then ' If UpToDate
If VSInfo.LastUpdate <= Format(Date, "dd/mmmm/yyyy") Then
Else
   If Bln < 11 Then ' If Error Read From Virus Definition File
      MsgBox "Error to Read Date in Your System !", vbCritical
   Else  ' If Out Of date
   MsgBox "Your Database is Out of date !", vbCritical, "ATV Guard "
      PesanARV.Label1.Caption = "Out of date"
      PesanARV.Label3.Visible = True
   End If
End If
End Function
Private Sub tmTgu_Timer()
Static X As Byte
X = X + 1
If X = 2 Then
   ' Message Show
     IsOutOfdate
   ' ===================
End If
If X = 7 Then
   tmTgu.Enabled = False
End If
End Sub
