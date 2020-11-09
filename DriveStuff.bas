Attribute VB_Name = "DriveStuff"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
' Return an array of drive names.
Public Function GetDriveNames() As String()
Dim all_drives As String
Dim len_drives As Long
Dim drive_array() As String

    ' See how long the drive buffer needs to be.
    len_drives = GetLogicalDriveStrings(0, all_drives)

    ' Allocate space for the buffer.
    all_drives = Space$(len_drives)

    ' Get the drive information.
    If GetLogicalDriveStrings(len_drives, all_drives) <> 0 Then
        drive_array = Split(all_drives, Chr$(0))

        ' Remove the last two blank entries.
        ReDim Preserve drive_array(0 To UBound(drive_array) - 2)
        GetDriveNames = drive_array
    End If
End Function


