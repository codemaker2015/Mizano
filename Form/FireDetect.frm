VERSION 5.00
Begin VB.Form FireDetect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firewall"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "FireDetect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   0
      Picture         =   "FireDetect.frx":000C
      ScaleHeight     =   3210
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin Mizano.Abutton Abutton3 
         Height          =   375
         Left            =   7440
         TabIndex        =   1
         Top             =   2640
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
         Left            =   5040
         TabIndex        =   2
         Top             =   2640
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
      Begin Mizano.Abutton Abutton1 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   2640
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
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   120
         Picture         =   "FireDetect.frx":6F24
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4800
         TabIndex        =   11
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4800
         TabIndex        =   10
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Firewall Detected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7200
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet "
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield Security"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FireDetect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REG As cRegistry
Private Sub Abutton1_Click()
Me.Hide
End Sub

Private Sub Abutton2_Click()
TerminateThisConnection (Me.Label4.Caption)
ModThreadProcess.Thread_Resume (Me.Label4.Caption)
ModLoadProcess.KillProcessById (Me.Label4.Caption)
Unload Me
End Sub

Private Sub Abutton3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Beep 800, 80: Beep 500, 80
AlwaysOnTop Me.hwnd, True
End Sub
