VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   0
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   3570
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   -120
      Width           =   7425
      Begin Mizano.ProgressBar P2 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         Color           =   12937777
         Color2          =   12937777
      End
      Begin VB.Timer TmrLoad 
         Interval        =   100
         Left            =   6360
         Top             =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   7440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Your PC is Protected by Mizano"
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
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   720
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Label lblLoad 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait Mizano is configuring environment."
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
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   5790
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Monitoring Intelligent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B9AF93&
         Height          =   405
         Index           =   4
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   4725
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Call SysTrayFunc
End Sub

Private Sub TmrLoad_Timer()
P2.value = P2.value + 5
        lblLoad.Caption = "Loading Mizano : " & P2.value & "% of 100%"
        
            If P2.value = 10 Then
          lblProcess.Caption = "Initializing Mizano System...."
            ElseIf P2.value = 50 Then
                lblProcess.Caption = " Accesing Windows Registry...."
            ElseIf P2.value = 85 Then
              lblProcess.Caption = "Completing Mizano Registry System..."
            End If
            
        If P2.value = 100 Then
            Unload Me
            'Call SysTrayFunc
            frmAntivirus.Show
        End If
        
End Sub

