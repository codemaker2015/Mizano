VERSION 5.00
Begin VB.Form Shield 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resident Shield"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "Shield.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Shield.frx":000C
   ScaleHeight     =   3435
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   0
      Picture         =   "Shield.frx":6F24
      ScaleHeight     =   3450
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin Mizano.Abutton Abutton1 
         Height          =   375
         Left            =   7560
         TabIndex        =   9
         Top             =   2880
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
      Begin VB.TextBox TxVirus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   3840
         TabIndex        =   1
         Top             =   2040
         Width           =   5925
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Threat Found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield"
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
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Real Time Protection"
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
         TabIndex        =   10
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   120
         Picture         =   "Shield.frx":DE3C
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Location "
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
         Left            =   2400
         TabIndex        =   8
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Threat name"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblPathText 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   2520
         Width           =   6015
      End
      Begin VB.Label lblProcRun 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   1080
         Width           =   5895
      End
      Begin VB.Label Label38 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Infected File"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblFoundVir 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Action"
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   1560
         Width           =   6075
      End
   End
End
Attribute VB_Name = "Shield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abutton1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Beep 800, 80: Beep 500, 80
AlwaysOnTop Me.hwnd, True
End Sub


