VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCopy 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   1185
   ClientTop       =   1545
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmCopy.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   7500
   Begin VB.TextBox txtDestFileName 
      Height          =   435
      HideSelection   =   0   'False
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDestFolder 
      Height          =   435
      HideSelection   =   0   'False
      Left            =   1440
      TabIndex        =   1
      Text            =   "Select Destination Folder..."
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox txtSource 
      Height          =   435
      HideSelection   =   0   'False
      Left            =   1440
      TabIndex        =   0
      Text            =   "Select Source File..."
      Top             =   840
      Width           =   4815
   End
   Begin VB.Image btnCopyFile 
      Height          =   975
      Left            =   3120
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image btnSelectFolder 
      Height          =   615
      Left            =   6360
      Top             =   1800
      Width           =   615
   End
   Begin VB.Image btnSelectFile 
      Height          =   735
      Left            =   6360
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SetCurrentDirectory Lib "kernel32" _
    Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Declare Function GetCurrentDirectory Lib "kernel32" _
    Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

'Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
        
Private Declare Function lstrcat Lib "kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Sub btnCopyFile_Click()
    If Not Dir(Trim(txtSource.Text)) = "" Then
        If Not Dir(Trim(txtDestFolder.Text), vbDirectory) = "" Then
            If Not Right(Trim(txtDestFolder.Text), 1) = "\" Then
                txtDestFolder.Text = Trim(txtDestFolder.Text) & "\"
            End If
            Dim destFile As String
            
            destFile = txtDestFolder
            If Not Dir(txtDestFolder & Trim(txtDestFileName.Text)) = "" Then
                Dim msg As String
                msg = "Destination folder already contains file with the same name." & vbNewLine
                msg = msg & "Select YES if you wish to overwrite existing file." & vbNewLine
                msg = msg & "Otherwise select NO and change destination file name."
                
                If MsgBox(msg, vbInformation + vbYesNo, "File Exists") = vbYes Then
                    Kill destFile
                Else
                    txtDestFileName.SelStart = 0
                    txtDestFileName.SelLength = Len(txtDestFileName.Text)
                    txtDestFileName.SetFocus
                    Exit Sub
                End If
            End If
            Shell "xcopy /S " & Chr$(34) & Trim(txtSource.Text) & Chr$(34) & " " & Chr$(34) & destFile & Chr$(34), vbHide
            MsgBox "File's done."
        Else
            MsgBox "Please select destination folder.", vbExclamation, "Missing Destination Folder"
        End If
    Else
        MsgBox "Please select source file.", vbExclamation, "Missing Source File"
    End If
End Sub

Private Sub btnSelectFile_Click()

On Error GoTo ErrHandler

    With CommonDialog1
        .CancelError = True
        .FLAGS = cdlOFNExplorer
        .ShowOpen
        If Not .FileName = "" Then
            txtSource.Text = .FileName
            txtDestFileName.Text = Mid(Trim(txtSource.Text), InStrRev(Trim(txtSource.Text), "\") + 1)
        Else
            txtSource.Text = "Select Source File..."
            txtDestFileName.Text = ""
        End If
    End With
    
    Exit Sub

ErrHandler:
    err.Clear
    txtSource.Text = "Select Source File..."
    txtDestFileName.Text = ""

End Sub

Private Sub btnSelectFolder_Click()
'===================================
Dim lRet As Long
Dim sBuffer As String
Dim sTitle As String
Dim tBrowseInfo As BrowseInfo
Dim sCurDir As String
Dim lPidl As Long

    sTitle = "Select Destination Folder"
    
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
                   BIF_EDITBOX Or BIF_VALIDATE Or BIF_NEWDIALOGSTYLE
    End With
    
    lRet = SHBrowseForFolder(tBrowseInfo)
    
    If lRet > 0 Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lRet, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtDestFolder.Text = sBuffer
    End If

End Sub

Private Sub Form_Load()
    txtSource.SelStart = 0
    txtSource.SelLength = Len(txtSource.Text)
    txtDestFolder.SelStart = 0
    txtDestFolder.SelLength = Len(txtDestFolder.Text)
End Sub


