Attribute VB_Name = "ModThread"
Private Declare Function DeleteFile Lib _
    "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
Option Explicit
Private Declare Function SetFileAttributes Lib _
    "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileName As String, _
    ByVal dwFileAttributes As Long) As Long
Private Const MaxLen = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const vbStar = "*"
Private Const vbAllFiles = "*.*"
Private Const vbBackslash = "\"
Private Const vbKeyDot = 46

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngCC As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Function DelVir(ByVal sFilePath As String) _
    As Long

    On Error Resume Next
    
    SetFileAttributes sFilePath, FILE_ATTRIBUTE_NORMAL
    DeleteFile sFilePath

End Function

Public Sub DeleteNow(ListView1 As ListView, _
    VirItem As Integer)
        
    On Error Resume Next
    
    Dim VirPath As String
    Dim VirCnt As Long, lVir As Long

    For VirCnt = 1 To ListView1.ListItems.count
        If ListView1.ListItems.Item(VirCnt).Selected = _
            True Then
            VirPath = ListView1.SelectedItem.SubItems _
                (VirItem)
            DelVir VirPath
            ListView1.ListItems.Remove VirCnt
        End If
        Exit Sub
    Next VirCnt
    
End Sub

Sub Main()
Dim iccex As tagInitCommonControlsEx
With iccex
                .lngSize = LenB(iccex)
                .lngCC = ICC_USEREX_CLASSES
            End With
            InitCommonControlsEx iccex
            On Error GoTo 0
Call ReadSig
Splash.Show
End Sub


