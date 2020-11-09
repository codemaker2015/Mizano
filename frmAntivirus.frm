VERSION 5.00
Begin VB.Form frmAntivirus 
   BorderStyle     =   0  'None
   Caption         =   "Home - Mizano"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmAntivirus.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image imgAbout 
      Height          =   615
      Left            =   360
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Image imgPanel 
      Height          =   735
      Left            =   240
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Image imgHome 
      Height          =   735
      Left            =   240
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Image imgDatabase 
      Height          =   3255
      Left            =   16920
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Image imgUpdate 
      Height          =   3015
      Left            =   12480
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Image imgUserStatus 
      Height          =   3375
      Left            =   8160
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Image imgRegistryFixer 
      Height          =   3375
      Left            =   3840
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Image imgPowerSystem 
      Height          =   3015
      Left            =   16680
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Image imgQuarantine 
      Height          =   2895
      Left            =   12360
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Image imgShield 
      Height          =   3015
      Left            =   8040
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image imageClose 
      Height          =   855
      Left            =   19560
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imageMaximize 
      Height          =   855
      Left            =   18480
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageMinimize 
      Height          =   855
      Left            =   17280
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imageFullscan 
      Height          =   3135
      Left            =   3840
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "frmAntivirus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageFullscan_Click()
    frmAntivirus.Picture = LoadPicture(App.path + "\Images\antivirus2.jpg")
    frmScanner.Show
End Sub
Private Sub imageClose_Click()
    Unload Me
End Sub

Private Sub imageMaximize_Click()
    frmParent.WindowState = vbMaximized
End Sub

Private Sub imageMinimize_Click()
    frmParent.WindowState = vbMinimized
    Unload Me
End Sub

Private Sub imgAbout_Click()
    About.Show
End Sub

Private Sub imgDatabase_Click()
    frmDatabase.Show
End Sub

Private Sub imgHome_Click()
    Unload Me
    frmHome.Show
End Sub

Private Sub imgPanel_Click()
    frmAntivirus.Picture = LoadPicture(App.path + "\Images\virus.jpg")
End Sub

Private Sub imgPowerSystem_Click()
    PowerRemoval.Show
End Sub

Private Sub imgQuarantine_Click()
    frmQuarantine.Show
End Sub

Private Sub imgRegistryFixer_Click()
    frmRegistryFixer.Show
End Sub

Private Sub imgShield_Click()
    FrmTest.Show
End Sub

Private Sub imgUserStatus_Click()
    StatusRegister.Show
End Sub
