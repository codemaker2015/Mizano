VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Firewall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Monitoring Firewall"
   ClientHeight    =   9435
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   20400
   Icon            =   "frmFirewall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFirewall.frx":0CCA
   ScaleHeight     =   9435
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   3840
      TabIndex        =   0
      Top             =   4560
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Direction"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Host"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remort Port"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "File Path"
         Object.Width           =   4762
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   7560
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox pic32 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pic16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   120
         Top             =   240
      End
      Begin ComctlLib.ImageList iml32 
         Left            =   2640
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   2040
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList iml16 
         Left            =   1440
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   840
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin Mizano.Abutton Abutton4 
      Height          =   375
      Left            =   11160
      TabIndex        =   1
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
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
   Begin Mizano.Abutton Abutton3 
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   7920
      Width           =   2295
      _ExtentX        =   4048
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
   Begin Mizano.Abutton Abutton2 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Mizano.Abutton Abutton1 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Image imageHome 
      Height          =   735
      Left            =   120
      Top             =   600
      Width           =   615
   End
   Begin VB.Image imageMinimize 
      Height          =   615
      Left            =   18000
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imageMaximize 
      Height          =   615
      Left            =   18840
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imageClose 
      Height          =   615
      Left            =   19800
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   10080
      X2              =   10800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©  VSoft Technologies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   11160
      TabIndex        =   6
      Top             =   8400
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   3480
      X2              =   11160
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "URL Blocker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Refreshlist 
         Caption         =   "Refresh"
      End
      Begin VB.Menu ViewCon 
         Caption         =   "View Connections"
      End
      Begin VB.Menu ShowPop 
         Caption         =   "Show Popup"
      End
      Begin VB.Menu AutoFresh 
         Caption         =   "Automatic Refresh"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "mnuModify"
      Visible         =   0   'False
      Begin VB.Menu mnuModifyTrust 
         Caption         =   "Trust"
      End
      Begin VB.Menu mnuModifyAsk 
         Caption         =   "Ask"
      End
      Begin VB.Menu mnuModifyBlock 
         Caption         =   "Block"
      End
      Begin VB.Menu mnuModifyDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Firewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hdcDest&, _
    ByVal X&, ByVal y&, ByVal FLAGS&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private shInfo As SHFILEINFO
Public Tablenum As Long
Private pTcpTable As MIB_TCPTABLE

Public Function GetRule(value As String) As String
Select Case value
Case "0" 'ask
GetRule = "Ask"
Case "1" 'block
GetRule = "Block"
Case "2" 'trust
GetRule = "Trust"
End Select

End Function

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
'mdlFirewall.Execute True
End Sub

Private Sub lvApplications_BeforeLabelEdit(Cancel As Integer)

End Sub

Public Sub Scanner()
  Dim i As Integer, o As Integer
  Dim fileNum As String
  Dim Item As ListItem
  Dim arrport As String
  Dim Algoritma As String
  Dim A As Long
On Error Resume Next
  ListView1.ListItems.Clear
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    iml32.ListImages.Clear
    iml16.ListImages.Clear
    DoEvents
  LoadProcesses
  For i = 0 To StatsLen - 1
  If Connection(i).FileName <> "" Then Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))) Else Set Item = ListView1.ListItems.Add(, , "Unknown")
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(1) = "Incomming" Else Item.SubItems(1) = "Outgoing"
    Item.SubItems(2) = Connection(i).LocalPort
    Item.SubItems(3) = Connection(i).RemoteHost
    Item.SubItems(4) = Connection(i).RemotePort
    Item.SubItems(5) = Connection(i).State
    Item.SubItems(6) = Connection(i).FileName
    Algoritma = GetChecksum(Connection(i).FileName)
    If trustApl(Algoritma) = True Then
    If Algoritma <> "" Then
    ModThreadProcess.Thread_Suspend (Connection(i).ProcessID)
    FireDetect.Label4.Caption = Connection(i).ProcessID
    FireDetect.Label3.Caption = Connection(i).State
    FireDetect.Label5.Caption = Connection(i).LocalPort
    FireDetect.Show
    End If
    End If
    Next
   DoEvents
   ShowIcons
   DoEvents
   Me.MousePointer = vbNormal
   End Sub
Public Sub RefreshConnection()
  Dim i As Integer, o As Integer
  Dim fileNum As String
  Dim Item As ListItem
  Dim arrport As String
  Dim Algoritma As String
  Dim A As Long
On Error Resume Next
  ListView1.ListItems.Clear
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    iml32.ListImages.Clear
    iml16.ListImages.Clear
    DoEvents
  LoadProcesses
  For i = 0 To StatsLen - 1
  If Connection(i).FileName <> "" Then Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))) Else Set Item = ListView1.ListItems.Add(, , "Unknown")
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(1) = "Incomming" Else Item.SubItems(1) = "Outgoing"
    Item.SubItems(2) = Connection(i).LocalPort
    Item.SubItems(3) = Connection(i).RemoteHost
    Item.SubItems(4) = Connection(i).RemotePort
    Item.SubItems(5) = Connection(i).State
    Item.SubItems(6) = Connection(i).FileName
    Next
   DoEvents
   ShowIcons
   DoEvents
   Me.MousePointer = vbNormal
   End Sub
Function trustApl(CZX As String) As Boolean
'On Error GoTo 10
FireDetect.Data1.DatabaseName = App.Path + "\Datnet.mdb"
FireDetect.Data1.RecordSource = "PortTable3"
FireDetect.Data1.Refresh
trustApl = False
FireDetect.Data1.Recordset.FindFirst "fldPort = '" _
& Trim(CZX) & "'"
If FireDetect.Data1.Recordset.NoMatch Then
trustApl = False
Else
trustApl = True
Exit Function
trustApl = False
End If
Exit Function
10:
MsgBox "" + Error$ + CStr(err)
Debug.Print Error$
End Function

Private Sub Abutton1_Click()
Scanner
End Sub
Private Sub Abutton3_Click()
Dim i As Integer
TerminateThisConnection (Connection(i).ProcessID)
ModThreadProcess.Thread_Resume (Connection(i).ProcessID)
ModLoadProcess.KillProcessById (Connection(i).ProcessID)
End Sub

Private Sub AutoFresh_Click()
If AutoFresh.Checked = True Then
Timer1.Enabled = False
AutoFresh.Checked = False
Else
AutoFresh.Checked = True
Timer1.Enabled = True
End If
End Sub

Private Sub Button2_Click()
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add 1, , "File", 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders.Add 2, , "Direction", 1000, lvwColumnCenter
ListView1.ColumnHeaders.Add 3, , "Local Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 4, , "Remote Host", ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders.Add 5, , "Remote Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 6, , "Status", 1300, lvwColumnCenter
ListView1.ColumnHeaders.Add 7, , "File Path", ListView1.Width \ 2 + 1000

RefreshConnection
End Sub

Private Sub Button4_Click()

End Sub

Private Sub ExitProg_Click()
Unload Me
End Sub


Private Sub Abutton2_Click()
ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add 1, , "File", 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders.Add 2, , "Direction", 1000, lvwColumnCenter
ListView1.ColumnHeaders.Add 3, , "Local Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 4, , "Remote Host", ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders.Add 5, , "Remote Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 6, , "Status", 1300, lvwColumnCenter
ListView1.ColumnHeaders.Add 7, , "File Path", ListView1.Width \ 2 + 1000

RefreshConnection
End Sub

Private Sub Abutton4_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim strParse() As String
Dim i As Integer
Dim Item As ListItem
'Me.Visible = False
pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
ShowPop.Checked = False
DoEvents

AlwaysOnTop Me.hwnd, True
End Sub

Private Sub LoadProcessInfo()

LoadAllNTProcesses
DoEvents

'If OldnUmProc = ProgEntries Then Exit Sub

'ListView6.ListItems.Clear

'For i = 0 To ProgEntries

'If CurProcesses(i).ProcessID <> "0" Then

 '   Set Item = ListView6.ListItems.Add(, , CurProcesses(i).FileName)
  '  Item.SubItems(1) = CurProcesses(i).ProcessID
   ' Item.Tag = i

'End If

'Next i

'OldnUmProc = ProgEntries
End Sub

Private Sub imageHome_Click()
 Unload Me
    frmHome.Show
End Sub

Private Sub Refreshlist_Click()
RefreshConnection
End Sub

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub



Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, shInfo, Len(shInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, shInfo, Len(shInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, shInfo.iIcon, pic32.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, shInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub ShowPop_Click()
If ShowPop.Checked = True Then
ShowPopup = False
ShowPop.Checked = False
Else
ShowPopup = True
ShowPop.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
'RefreshView
 
Dim pdwSize As Long
Dim bOrder As Long
Dim nRet As Long
Dim TableLen As Long

nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)

TableLen = pTcpTable.dwNumEntries
If Tablenum <> TableLen Then RefreshConnection
Tablenum = TableLen

End Sub

Private Sub ViewCon_Click()

ListView1.ColumnHeaders.Clear

ListView1.ColumnHeaders.Add 1, , "File", 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders.Add 2, , "Direction", 1000, lvwColumnCenter
ListView1.ColumnHeaders.Add 3, , "Local Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 4, , "Remote Host", ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders.Add 5, , "Remote Port", 1100, lvwColumnCenter
ListView1.ColumnHeaders.Add 6, , "Status", 1300, lvwColumnCenter
ListView1.ColumnHeaders.Add 7, , "File Path", ListView1.Width \ 2 + 1000

RefreshConnection
End Sub





