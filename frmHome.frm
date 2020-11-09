VERSION 5.00
Begin VB.Form frmHome 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   10215
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image imgAbout 
      Height          =   2415
      Left            =   3720
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Image imageBootable 
      Height          =   1935
      Index           =   1
      Left            =   720
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Image imageTransfer 
      Height          =   1935
      Index           =   1
      Left            =   720
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Image imageParentcontrol 
      Height          =   1815
      Index           =   1
      Left            =   600
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Image imageAntivirus 
      Appearance      =   0  'Flat
      Height          =   1935
      Index           =   1
      Left            =   600
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Image imageClose 
      Height          =   855
      Left            =   19560
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageMaximize 
      Height          =   855
      Left            =   18480
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageMinimize 
      Height          =   855
      Left            =   17400
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageBootable 
      Height          =   3735
      Index           =   0
      Left            =   16920
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Image imageTransfer 
      Height          =   3735
      Index           =   0
      Left            =   12360
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Image imageParentcontrol 
      Height          =   4095
      Index           =   0
      Left            =   7920
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Image imageAntivirus 
      Appearance      =   0  'Flat
      Height          =   3495
      Index           =   0
      Left            =   3720
      Top             =   6360
      Width           =   2655
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageAntivirus_Click(Index As Integer)
    Splash.Show
End Sub

Private Sub imageBootable_Click(Index As Integer)
    frmBootable.Show
    'MsgBox "Sorry, This facility is currently unavailable!!!", vbInformation, "Mizano"
End Sub

Private Sub imageClose_Click()
   Unload Me
End Sub

Private Sub imageParentcontrol_Click(Index As Integer)
    frmParent.Show
End Sub

Private Sub imageTransfer_Click(Index As Integer)
    frmCopy.Show

'    MsgBox "Sorry, This facility is currently unavailable!!!", vbInformation, "Mizano"
End Sub

Private Sub imgAbout_Click()
About.Show
End Sub
