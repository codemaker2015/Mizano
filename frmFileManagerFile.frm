VERSION 5.00
Begin VB.Form frmFileManagerFile 
   BackColor       =   &H8000000E&
   Caption         =   "Steganos - File Manager"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2070
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   43
      Left            =   17310
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   43
      Left            =   17325
      TabIndex        =   44
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   42
      Left            =   15630
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   42
      Left            =   15645
      TabIndex        =   43
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   41
      Left            =   13950
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   41
      Left            =   13965
      TabIndex        =   42
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   40
      Left            =   12270
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   40
      Left            =   12285
      TabIndex        =   41
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   39
      Left            =   10590
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   39
      Left            =   10605
      TabIndex        =   40
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   38
      Left            =   8910
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   38
      Left            =   8925
      TabIndex        =   39
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   37
      Left            =   7230
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   37
      Left            =   7245
      TabIndex        =   38
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   36
      Left            =   5550
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   36
      Left            =   5565
      TabIndex        =   37
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   35
      Left            =   3870
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   35
      Left            =   3885
      TabIndex        =   36
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   34
      Left            =   2190
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   34
      Left            =   2205
      TabIndex        =   35
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   33
      Left            =   510
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   33
      Left            =   525
      TabIndex        =   34
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   32
      Left            =   17310
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   32
      Left            =   17325
      TabIndex        =   33
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   31
      Left            =   15630
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   31
      Left            =   15645
      TabIndex        =   32
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   30
      Left            =   13950
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   30
      Left            =   13965
      TabIndex        =   31
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   29
      Left            =   12270
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   29
      Left            =   12285
      TabIndex        =   30
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   28
      Left            =   10590
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   28
      Left            =   10605
      TabIndex        =   29
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   27
      Left            =   8910
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   27
      Left            =   8925
      TabIndex        =   28
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   26
      Left            =   7230
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   26
      Left            =   7245
      TabIndex        =   27
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   25
      Left            =   5550
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   25
      Left            =   5565
      TabIndex        =   26
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   24
      Left            =   3870
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   24
      Left            =   3885
      TabIndex        =   25
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   23
      Left            =   2190
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   23
      Left            =   2205
      TabIndex        =   24
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   22
      Left            =   510
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   22
      Left            =   525
      TabIndex        =   23
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   21
      Left            =   17310
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   21
      Left            =   17325
      TabIndex        =   22
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   20
      Left            =   15630
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   20
      Left            =   15645
      TabIndex        =   21
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   19
      Left            =   13950
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   19
      Left            =   13965
      TabIndex        =   20
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   18
      Left            =   12270
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   18
      Left            =   12285
      TabIndex        =   19
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   17
      Left            =   10590
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   17
      Left            =   10605
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   16
      Left            =   8910
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   16
      Left            =   8925
      TabIndex        =   17
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   15
      Left            =   7230
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   15
      Left            =   7245
      TabIndex        =   16
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   13
      Left            =   3870
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   14
      Left            =   5550
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   14
      Left            =   5550
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   13
      Left            =   3885
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   11
      Left            =   510
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   12
      Left            =   2190
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   12
      Left            =   2190
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   11
      Left            =   525
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   0
      Left            =   495
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   1
      Left            =   2175
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   2
      Left            =   3855
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   3
      Left            =   5535
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   4
      Left            =   7215
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   5
      Left            =   8895
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   6
      Left            =   10575
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   7
      Left            =   12255
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   8
      Left            =   13935
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   9
      Left            =   15615
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image imgFile 
      Height          =   1500
      Index           =   10
      Left            =   17295
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   2175
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   3855
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   5535
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   7215
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   8895
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   10575
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   12255
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   13935
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   15615
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   10
      Left            =   17295
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   510
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image imgPrevious 
      Height          =   495
      Left            =   250
      Picture         =   "frmFileManagerFile.frx":0000
      Stretch         =   -1  'True
      Top             =   250
      Width           =   495
   End
End
Attribute VB_Name = "frmFileManagerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_DIR = &H18D
Private Const LB_RESETCONTENT = &H184

Private Const DDL_ARCHIVE = &H20
Private Const DDL_DIRECTORY = &H10
Private Const DDL_DRIVES = &H4000
Private Const DDL_EXCLUSIVE = &H8000
Private Const DDL_HIDDEN = &H2&
Private Const DDL_READONLY = &H1
Private Const DDL_READWRITE = &H0
Private Const DDL_SYSTEM = &H4

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
    Dim myname As String
    File1.path = folder
    Dim i As Integer
    i = 0

    Dim aPattern$()
 
    'Clear content
     SendMessage File1.hwnd, LB_RESETCONTENT, 0, 0
 
    'Add files specified in the pattern property
     aPattern = Split(File1.Pattern, ";")
     For i = 0 To UBound(aPattern)
         SendMessage File1.hwnd, LB_DIR, DDL_ARCHIVE Or DDL_HIDDEN, ByVal Replace$(File1.path & "\" & Trim$(aPattern(i)), "\\", "\")
     Next i
    
     For i = 0 To File1.ListCount - 1
           lblFile(i).Caption = File1.List(i)
           imgFile(i).Picture = LoadPicture(GetIcon(getextension(lblFile(i).Caption)))
        If i = 20 Then
            Exit Sub
        End If
     Next i
End Sub

Private Sub imgFile_Click(Index As Integer)
    Dim extn As String, Name As String
    If getextension(file) <> "" Then
        Name = getfilename(file)
        file = Mid(file, 1, Len(file) - Len(Name) - 1)
    End If
    
    file = file & "\" & lblFile(Index).Caption
    
    'MsgBox file
    
    If GetAttr(file) = 2 And vbHidden Then
        frmOptions.lblHide.Caption = "Show"
    Else
        frmOptions.lblHide.Caption = "Hide"
    End If
    
    frmOptions.Show
    frmOptions.Left = imgFile(Index).Left + 1000
    frmOptions.Top = imgFile(Index).Top + 1000
End Sub

Private Sub imgPrevious_Click()
    Unload Me
    frmFileManagerFolder.Show
End Sub
