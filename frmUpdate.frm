VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   0  'None
   Caption         =   "Update - Mizano"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmUpdate.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   6240
      Picture         =   "frmUpdate.frx":19E6D
      ScaleHeight     =   6930
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   3480
      Width           =   8505
      Begin VB.PictureBox Signature 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5415
         Left            =   120
         Picture         =   "frmUpdate.frx":20D85
         ScaleHeight     =   5415
         ScaleWidth      =   8295
         TabIndex        =   1
         Top             =   1200
         Width           =   8295
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Type Detection From Signature :"
            Height          =   1695
            Left            =   3840
            TabIndex        =   4
            Top             =   1320
            Width           =   3975
            Begin VB.Label lbltype2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1440
               TabIndex        =   11
               Top             =   600
               Width           =   1845
            End
            Begin VB.Label lblname1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1440
               TabIndex        =   10
               Top             =   360
               Width           =   1845
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Author      :"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   6
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   120
               Left            =   240
               TabIndex        =   9
               Top             =   1320
               Width           =   1170
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Thread Type :"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   6
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   120
               Left            =   240
               TabIndex        =   8
               Top             =   600
               Width           =   1170
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FileType    :"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   6
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   120
               Left            =   240
               TabIndex        =   7
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Thread Name :"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   6
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   120
               Left            =   240
               TabIndex        =   6
               Top             =   360
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Risk        :"
               BeginProperty Font 
                  Name            =   "Terminal"
                  Size            =   6
                  Charset         =   255
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   120
               Left            =   240
               TabIndex        =   5
               Top             =   1080
               Width           =   1170
            End
         End
         Begin Mizano.Abutton Abutton2 
            Height          =   375
            Left            =   4920
            TabIndex        =   2
            Top             =   4800
            Width           =   2895
            _ExtentX        =   5106
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
            Left            =   4920
            TabIndex        =   3
            Top             =   4320
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSComctlLib.ListView lstVirus 
            Height          =   4095
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   7223
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList2"
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Threat Name"
               Object.Width           =   4886
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   480
            Top             =   1200
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmUpdate.frx":27C9D
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Signature :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   17
            Top             =   4560
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Update      :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   16
            Top             =   4920
            Width           =   1815
         End
         Begin VB.Label lblLastUpdate 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2280
            TabIndex        =   15
            Top             =   4920
            Width           =   2415
         End
         Begin VB.Label lblVirusCount 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Image Image13 
            Height          =   570
            Left            =   6840
            Picture         =   "frmUpdate.frx":28037
            Stretch         =   -1  'True
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Update"
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
            Left            =   5400
            TabIndex        =   13
            Top             =   0
            Width           =   1815
         End
      End
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
Attribute VB_Name = "frmUpdate"
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

