VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form PowerRemoval 
   BorderStyle     =   0  'None
   Caption         =   "Power Removal"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   20490
   Icon            =   "PowerRemoval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "PowerRemoval.frx":0CCA
   ScaleHeight     =   8610
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   4560
      Picture         =   "PowerRemoval.frx":1AB37
      ScaleHeight     =   5445
      ScaleWidth      =   10635
      TabIndex        =   2
      Top             =   3600
      Width           =   10635
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Control Button"
         Height          =   1215
         Left            =   6840
         TabIndex        =   13
         Top             =   4080
         Width           =   3615
         Begin Mizano.Abutton Abutton3 
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
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
         Begin Mizano.Abutton Abutton6 
            Height          =   375
            Left            =   1800
            TabIndex        =   14
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
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
            Left            =   1800
            TabIndex        =   15
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Runing Process Name  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   6585
         Begin VB.PictureBox PicIconP32 
            Height          =   495
            Left            =   6600
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblPath 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   5595
         End
         Begin VB.Label lblFile 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   5070
         End
         Begin VB.Label lblDescription 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   6255
         End
         Begin VB.Label lblCompany 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   6255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lstView 
         Height          =   3480
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6138
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList3"
         SmallIcons      =   "ImageList3"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process Name"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Directory"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Discription"
            Object.Width           =   6880
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Size"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Process ID"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Base P"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Threads"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Attributes"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Priority"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Checksum"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Mem Usage"
            Object.Width           =   1766
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Power Removal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   990
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   75
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   375
         Top             =   225
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   4210752
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Image ImgIcon 
         Height          =   240
         Left            =   75
         Picture         =   "PowerRemoval.frx":21A4F
         Top             =   225
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Timer tmrProcessRefresh 
      Interval        =   5000
      Left            =   150
      Top             =   1170
   End
   Begin VB.Image imageHome 
      Height          =   855
      Left            =   240
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Image imageMinimize 
      Height          =   855
      Left            =   17280
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imageMaximize 
      Height          =   855
      Left            =   18480
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageClose 
      Height          =   855
      Left            =   19560
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PowerRemoval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageClose_Click()
    Unload Me
End Sub

Private Sub imageHome_Click()
 Unload Me
    frmHome.Show
End Sub

Private Sub imageMaximize_Click()
    frmParent.WindowState = vbMaximized
End Sub

Private Sub imageMinimize_Click()
    frmParent.WindowState = vbMinimized
End Sub

Private Sub Abutton1_Click()
    ViewProcess
End Sub

Private Sub Abutton2_Click()
Dim i As Integer
    Dim Pesan As String, strFile As String
    Dim fso As New FileSystemObject, FileName As file
    
    strFile = lstView.SelectedItem.SubItems(1)
    Set FileName = fso.GetFile(strFile)
    
    Pesan = "WARNING: Terminate a process can cause undesired" & vbCrLf & _
            "results including loss of data and system instability. The" & vbCrLf & _
            "process will not be given the chance to save its state or" & vbCrLf & _
            "data before it is terminated." & vbCrLf & vbCrLf & _
            "Are you sure you want to terminate process" & " " & FileName.ShortName
            If MsgBox(Pesan, vbYesNo + 48, APP_PROGRAM & " ATV Guard" & Chr(0)) = vbYes Then
               Dim h As Long
                   h = lstView.SelectedItem.Index
                    For i = 1 To lstView.ListItems.count
                      If lstView.ListItems(i).Selected Then
                        Call KillProcessById(CLng(lstView.ListItems(i).Tag))
                      End If
                    Next i
            End If
    ViewProcess
End Sub

Private Sub Abutton4_Click()

End Sub

Private Sub Abutton5_Click()

End Sub


Private Sub Abutton6_Click()
Me.Hide
End Sub

Private Sub Form_Load()
ViewProcess
AlwaysOnTop Me.hwnd, True
End Sub

Private Sub lstView_Click()
    Dim strFile As String, uProcess As PROCESSENTRY32
    Dim hVer As VERHEADER
    Dim fso As New FileSystemObject, FileInfo As file
    Dim strF As String
    
    PicIconP32.Cls
    strFile = lstView.SelectedItem.SubItems(1)
    
    If strF <> strFile Then
        On Error GoTo SalahProses
        Set FileInfo = fso.GetFile(strFile)
        GetVerHeader strFile, hVer
    
        Label8.Caption = "File"
        Label7.Caption = "Folder"
    
        lblDescription.Caption = hVer.FileDescription
        lblCompany.Caption = hVer.CompanyName
        lblFile.Caption = ": " & FileInfo.ShortName ' GetFileName(strFile)
        lblPath.Caption = ": " & FileInfo.ParentFolder ' GetFilePath(strFile)
        Exit Sub
    End If
    
SalahProses:
        MsgBox err.Description & " " & " " & _
        "or File Can not be delete.", vbExclamation, "Warning"
End Sub

Private Sub lstView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    lstView.Sorted = True
    
    lstView.SortKey = ColumnHeader.Index - 1
    If lstView.SortOrder = lvwDescending Then
       lstView.SortOrder = lvwAscending
    Else
       lstView.SortOrder = lvwDescending
    End If

End Sub



Private Sub MnPPS_Click()

End Sub
