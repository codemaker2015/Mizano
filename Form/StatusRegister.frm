VERSION 5.00
Begin VB.Form StatusRegister 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   0
      Picture         =   "StatusRegister.frx":0000
      ScaleHeight     =   3210
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      Begin Mizano.Abutton Abutton1 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   2640
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "STATUS"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3015
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Registred To : VSoft Technologies"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   1200
            Width           =   2775
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "License to freeware"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   3720
         Picture         =   "StatusRegister.frx":6F18
         Top             =   1440
         Width           =   960
      End
   End
End
Attribute VB_Name = "StatusRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abutton1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
 strUserCom = GetUserCom
End Sub
