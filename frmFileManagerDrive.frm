VERSION 5.00
Begin VB.Form frmFileManagerDrive 
   BackColor       =   &H8000000E&
   Caption         =   "Steganos"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   ScaleHeight     =   9435
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   10
      Left            =   13200
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   9
      Left            =   10320
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   8
      Left            =   7680
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   7
      Left            =   4800
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   6
      Left            =   1920
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   15960
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   13200
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   10320
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   7680
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   6
      Left            =   2160
      Top             =   4680
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   7
      Left            =   5040
      Top             =   4680
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   8
      Left            =   7920
      Top             =   4680
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   9
      Left            =   10560
      Top             =   4680
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   10
      Left            =   13440
      Top             =   4680
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   5
      Left            =   16200
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   4
      Left            =   13440
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   3
      Left            =   10560
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   2
      Left            =   7920
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   1
      Left            =   5040
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image imgDrive 
      Height          =   960
      Index           =   0
      Left            =   2160
      Top             =   2160
      Width           =   960
   End
End
Attribute VB_Name = "frmFileManagerDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Integer
    i = 0
    For i = 0 To Drive1.ListCount - 1
        lblDriveName(i).Caption = Drive1.List(i)
        imgDrive(i).Picture = LoadPicture(geticon("drive"))
    Next i
End Sub

Private Sub imgDrive_Click(Index As Integer)
    drive = lblDriveName(Index).Caption
    frmFileManagerFolder.Dir1.path = Left$(drive, 1) & ":\"
    frmFileManagerFolder.Show
    Unload Me
End Sub
