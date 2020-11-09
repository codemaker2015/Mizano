VERSION 5.00
Begin VB.Form frmQuarantine 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmQuarantine.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   5760
      Picture         =   "frmQuarantine.frx":19E6D
      ScaleHeight     =   6930
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   4080
      Width           =   8505
      Begin VB.FileListBox Qrtna 
         Height          =   2820
         Left            =   240
         Pattern         =   "*.atv"
         TabIndex        =   1
         Top             =   1560
         Width           =   7935
      End
      Begin Mizano.Abutton Abutton9 
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   4800
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Mizano.Abutton Abutton8 
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   4800
         Width           =   1815
         _ExtentX        =   3201
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
      Begin Mizano.Abutton Abutton7 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   4800
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   240
         TabIndex        =   5
         Top             =   4800
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Quarantine"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Select one or more files to delete or restore from prison [ Quarantine Files ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   4440
         Width           =   5415
      End
      Begin VB.Label lblviri 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Viruses in Quarantine"
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
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2040
      End
   End
   Begin VB.Image imageAbout 
      Height          =   735
      Left            =   240
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Image imagepannel 
      Height          =   855
      Left            =   240
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Image imageHome 
      Height          =   975
      Left            =   240
      Top             =   5160
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
Attribute VB_Name = "frmQuarantine"
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

