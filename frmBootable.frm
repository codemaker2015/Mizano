VERSION 5.00
Begin VB.Form frmBootable 
   Caption         =   "Mizano - Bootable"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbBootable 
      Caption         =   "Bootable"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbDrive 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "--Select--"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtResults 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lstDrives 
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drive:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   645
      Width           =   1095
   End
   Begin VB.Label lblDriveName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   650
      Width           =   1575
   End
End
Attribute VB_Name = "frmBootable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_SETTABSTOPS = &HCB

Private Sub cmbBootable_Click()
    Dim oShell As Object
    Set oShell = CreateObject("WSCript.shell")
    If cmbDrive.Text <> "--Select--" Then
        Dim command As String
        command = App.path & "\boot\booter.bat"
        oShell.Run "cmd /C " & Chr$(34) & command & Chr$(34), 0, True
        MsgBox "Success!!!" & vbCrLf & "Please copy windows ISO data to this drive"
    Else
        MsgBox "Noremovable media found", vbCritical, "Mizano bootable"
    End If
End Sub

Private Sub Form_Load()
Dim drive_names() As String
Dim i As Integer
Dim tabs(1 To 1) As Long

    ' Load the drive names.
    drive_names = GetDriveNames()
    lstDrives.Clear
    For i = LBound(drive_names) To UBound(drive_names)
        lstDrives.AddItem drive_names(i)
    Next i

    ' Set tabs in txtResults.
    tabs(1) = 120

    ' Set the tabs.
    SendMessage txtResults.hwnd, EM_SETTABSTOPS, 1, tabs(1)
    
    Dim drive_info As New DriveInfo

    ' Clear previous results. (Loading data for
    ' an empty drive can take a little while.)
    txtResults.Text = ""
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Get the drive's information.
    Dim isAvailable As Boolean
    
    For i = 0 To lstDrives.ListCount - 1
        drive_info.Initialize lstDrives.List(i)
        If drive_info.getType() = "Removable" Then
            cmbDrive.AddItem (lstDrives.List(i))
            isAvailable = True
        End If
    Next i
    If isAvailable = True Then
        cmbDrive.Text = cmbDrive.List(0)
        lblDriveName.Caption = drive_info.VolumeName
    Else
        MsgBox "No removable media found.", vbCritical, "Mizano - Bootable"
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Resize()
Dim wid As Single

    lstDrives.Height = ScaleHeight

    wid = ScaleWidth - txtResults.Left
    If wid < 120 Then wid = 120
    txtResults.Move txtResults.Left, 0, wid, ScaleHeight
End Sub


' Display information about this drive.
Private Sub lstDrives_Click()
Dim drive_info As New DriveInfo

    ' Clear previous results. (Loading data for
    ' an empty drive can take a little while.)
    txtResults.Text = ""
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Get the drive's information.
    drive_info.Initialize lstDrives.Text

    ' Display the drive's information.
    txtResults.Text = drive_info.ToString()
    MsgBox txtResults.Text
    Screen.MousePointer = vbDefault
End Sub

