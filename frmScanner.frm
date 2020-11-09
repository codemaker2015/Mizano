VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmScanner 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmScanner.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicScanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   5520
      Picture         =   "frmScanner.frx":19E6D
      ScaleHeight     =   6930
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   3600
      Width           =   8505
      Begin VB.PictureBox Scanner 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   240
         Picture         =   "frmScanner.frx":20D85
         ScaleHeight     =   5535
         ScaleWidth      =   8175
         TabIndex        =   1
         Top             =   1200
         Width           =   8175
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check All "
            Height          =   195
            Left            =   120
            Picture         =   "frmScanner.frx":27C9D
            TabIndex        =   11
            Top             =   0
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   6135
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            Picture         =   "frmScanner.frx":2EBB5
            ScaleHeight     =   285
            ScaleWidth      =   7905
            TabIndex        =   2
            Top             =   5160
            Width           =   7935
            Begin VB.Label lblCount 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   5130
               TabIndex        =   5
               ToolTipText     =   "Кол-во проверенных"
               Top             =   30
               Width           =   1005
            End
            Begin VB.Label lblFound 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6180
               TabIndex        =   4
               ToolTipText     =   "Кол-во найденых"
               Top             =   30
               Width           =   690
            End
            Begin VB.Label lblCleaned 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6930
               TabIndex        =   3
               ToolTipText     =   "Кол-во вылеченых"
               Top             =   30
               Width           =   870
            End
         End
         Begin Mizano.Abutton Abutton4 
            Height          =   375
            Left            =   5520
            TabIndex        =   6
            Top             =   4680
            Width           =   2535
            _ExtentX        =   4471
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
         Begin Mizano.Abutton C 
            Height          =   375
            Left            =   2760
            TabIndex        =   7
            Top             =   4680
            Width           =   2775
            _ExtentX        =   4895
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
         Begin Mizano.Abutton Command2 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   4680
            Width           =   2655
            _ExtentX        =   4683
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
         Begin Mizano.Abutton Command5 
            Height          =   375
            Left            =   6360
            TabIndex        =   9
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
         Begin MSComctlLib.ImageList ImlVir 
            Left            =   5640
            Top             =   1440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   18
            ImageHeight     =   17
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmScanner.frx":2EEF7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmScanner.frx":2F246
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmScanner.frx":2F730
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2295
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   4048
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImlVir"
            SmallIcons      =   "ImlVir"
            ForeColor       =   255
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Threat Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Location "
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Report"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date/Time"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Size [ Byte]"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblVir 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   23
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label lblFile 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Label lblScan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1200
            TabIndex        =   21
            Top             =   3840
            Width           =   6855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Threat"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "File(s)"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Object Scan"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label JumlahFile 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6360
            TabIndex        =   17
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label56 
            BackColor       =   &H00FFFFFF&
            Caption         =   "File Scanned"
            Height          =   255
            Left            =   5280
            TabIndex        =   16
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Label55 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dir Scanned"
            Height          =   255
            Left            =   5280
            TabIndex        =   15
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label JumlahDirectori 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6360
            TabIndex        =   14
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   8520
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Image imageAbout 
      Height          =   735
      Left            =   240
      Top             =   7560
      Width           =   2535
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
Attribute VB_Name = "frmScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageAbout_Click()
 About.Show
End Sub

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

