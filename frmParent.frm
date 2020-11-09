VERSION 5.00
Begin VB.Form frmParent 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Parent Control"
   ClientHeight    =   11025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmParent.frx":0000
   ScaleHeight     =   11025
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image imageHome 
      Height          =   735
      Left            =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imageBootable 
      Height          =   1695
      Left            =   480
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Image imgTransfer 
      Height          =   1695
      Left            =   600
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Image imgParentControl 
      Height          =   1695
      Left            =   600
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image imgAntivirus 
      Height          =   1815
      Left            =   720
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image imageClose 
      Height          =   735
      Left            =   19800
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imageMaximize 
      Height          =   615
      Left            =   18840
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imageMinimize 
      Height          =   615
      Left            =   18000
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imageurlblocker 
      Height          =   3855
      Left            =   16080
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Image imageHistory 
      Height          =   3615
      Left            =   9600
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Image imageLocking 
      Height          =   3735
      Left            =   3600
      Top             =   5280
      Width           =   3015
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageClose_Click()
    Unload Me
End Sub

Private Sub imageHistory_Click()
    Unload Me
    frmHistory.Show
End Sub

Private Sub imageHome_Click()
    Unload Me
    frmHome.Show
End Sub

Private Sub imageLocking_Click()
    frmFileManagerDrive.Show
End Sub

Private Sub imageMaximize_Click()
    frmParent.WindowState = vbMaximized
End Sub

Private Sub imageMinimize_Click()
    frmParent.WindowState = vbMinimized
End Sub

Private Sub imageurlblocker_Click()
    'Firewall.Show
    frmURLBlocker.Show
End Sub

Private Sub imgAntivirus_Click()
    Unload Me
    frmAntivirus.Show
End Sub

Private Sub imgParentControl_Click()
    frmParent.Show
End Sub

Private Sub imgTransfer_Click()
    Unload Me
    frmCopy.Show
End Sub
