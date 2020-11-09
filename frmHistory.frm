VERSION 5.00
Begin VB.Form frmHistory 
   Caption         =   "Mizano -History"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   9030
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   9855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back"
      Height          =   495
      Left            =   18240
      TabIndex        =   1
      Top             =   9840
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Code
Private Const CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
Private Const CSIDL_PROGRAMS = 2 '// Program Files
Private Const CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
Private Const CSIDL_PRINTERS = 4 '// Printers - virtual folder
Private Const CSIDL_DOCUMENTS = 5 '// My Documents
Private Const CSIDL_FAVORITES = 6 '// Favourites
Private Const CSIDL_STARTUP = 7 '// Startup Folder
Private Const CSIDL_RECENT = 8 '// Recent Documents
Private Const CSIDL_SENDTO = 9 '// Send To Folder
Private Const CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
Private Const CSIDL_STARTMENU = 11 '// Start Menu
Private Const CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
Private Const CSIDL_DRIVES = 17 '// My Computer - virtual folder
Private Const CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
Private Const CSIDL_NETHOOD = 19 '// NetHood Folder
Private Const CSIDL_FONTS = 20 '// Fonts folder
Private Const CSIDL_SHELLNEW = 21 '// ShellNew folder

Private Sub Form_Load()
    'MsgBox "Recent Folder " & fGetSpecialFolder(CSIDL_RECENT)
    File1.path = fGetSpecialFolder(CSIDL_RECENT)
    Dim i As Integer
    i = 0
    For i = 0 To File1.ListCount - 1
        'MsgBox File1.List(i)
        List1.AddItem (GetTargetPath(fGetSpecialFolder(CSIDL_RECENT) & "\" & File1.List(i)))
    Next i
End Sub

