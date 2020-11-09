VERSION 5.00
Begin VB.Form frmFileManagerFolder 
   BackColor       =   &H8000000E&
   Caption         =   "Steganos - File Manager"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgPrevious 
      Height          =   495
      Left            =   240
      Picture         =   "frmFileManagerFolder.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   26
      Left            =   17760
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   26
      Left            =   17280
      TabIndex        =   27
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   25
      Left            =   15720
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   25
      Left            =   15240
      TabIndex        =   26
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   24
      Left            =   13560
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   24
      Left            =   13080
      TabIndex        =   25
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   23
      Left            =   11520
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   23
      Left            =   11040
      TabIndex        =   24
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   22
      Left            =   9360
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   22
      Left            =   8880
      TabIndex        =   23
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   21
      Left            =   7080
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   21
      Left            =   6600
      TabIndex        =   22
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   20
      Left            =   4920
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   20
      Left            =   4440
      TabIndex        =   21
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   19
      Left            =   2880
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   19
      Left            =   2400
      TabIndex        =   20
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   18
      Left            =   720
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   18
      Left            =   240
      TabIndex        =   19
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   17
      Left            =   17760
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   17
      Left            =   17280
      TabIndex        =   18
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   16
      Left            =   15720
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   16
      Left            =   15240
      TabIndex        =   17
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   15
      Left            =   13560
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   15
      Left            =   13080
      TabIndex        =   16
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   14
      Left            =   11520
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   14
      Left            =   11040
      TabIndex        =   15
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   13
      Left            =   9360
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   13
      Left            =   8880
      TabIndex        =   14
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   12
      Left            =   7080
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   12
      Left            =   6600
      TabIndex        =   13
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   11
      Left            =   4920
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   11
      Left            =   4440
      TabIndex        =   12
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   10
      Left            =   2880
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   10
      Left            =   2400
      TabIndex        =   11
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   9
      Left            =   720
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   8
      Left            =   17760
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   8
      Left            =   17280
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   7
      Left            =   15720
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   7
      Left            =   15240
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   6
      Left            =   13560
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   6
      Left            =   13080
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   5
      Left            =   11520
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   11040
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   4
      Left            =   9360
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   8880
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   3
      Left            =   7080
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   6600
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   2
      Left            =   4920
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   2
      Left            =   4440
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   1
      Left            =   2880
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image imgFolder 
      Height          =   1320
      Index           =   0
      Left            =   720
      Top             =   3000
      Width           =   915
   End
End
Attribute VB_Name = "frmFileManagerFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Integer
    i = 0
    Dir1.path = drive
    
    For i = 0 To Dir1.ListCount
           lblCaption(i).Caption = getfilename2(Dir1.List(i))
           imgFolder(i).Picture = LoadPicture(GetIcon("folder"))
        If i = 20 Then
            Exit Sub
        End If
    Next i
End Sub

Private Sub imgFolder_Click(Index As Integer)
    folder = Dir1.List(Index)
    file = Dir1.List(Index)
    frmFileManagerFile.File1.path = folder
    frmFileManagerFile.Show
    Unload Me
End Sub

Private Sub imgPrevious_Click()
    frmFileManagerDrive.Show
    Unload Me
End Sub
