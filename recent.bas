Attribute VB_Name = "recent"
'Module Code
Option Explicit
Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Const MAX_PATH As Integer = 260

Public Function GetTargetPath(ByVal FileName As String)
 
Dim Obj As Object
Set Obj = CreateObject("WScript.Shell")
 
Dim Shortcut As Object
Set Shortcut = Obj.CreateShortcut(FileName)
GetTargetPath = Shortcut.TargetPath
Shortcut.Save
 
End Function


Public Function fGetSpecialFolder(CSIDL As Long) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(frmHistory.hwnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
                                
                            End If
End Function

