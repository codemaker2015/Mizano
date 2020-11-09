VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' API declarations.
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const FILE_CASE_PRESERVED_NAMES = &H2
Private Const FILE_CASE_SENSITIVE_SEARCH = &H1
Private Const FILE_UNICODE_ON_DISK = &H4
Private Const FILE_PERSISTENT_ACLS = &H8
Private Const FILE_FILE_COMPRESSION = &H10
Private Const FILE_VOLUME_IS_COMPRESSED = &H8000
Private Const FILE_SUPPORTS_ENCRYPTION = &H20000
Private Const FILE_SUPPORTS_OBJECT_IDS = &H10000
Private Const FILE_SUPPORTS_REPARSE_POINTS = &H80
Private Const FILE_SUPPORTS_SPARSE_FILES = &H40
Private Const FILE_VOLUME_QUOTAS = &H20

' Public information.
Public Enum DriveTypes
    drivetype_UNKNOWN = 0
    drivetype_NO_ROOT_DIR = 1
    drivetype_REMOVABLE = 2
    drivetype_FIXED = 3
    drivetype_REMOTE = 4
    drivetype_CDROM = 5
    drivetype_RAMDISK = 6
End Enum

Public DriveType As DriveTypes
Public VolumeName As String
Public FileSystemName As String
Public SerialNumber As Long
Public MaxComponentLength As Long
Public SupportsLongFileNames As Boolean
Public FileSystemFlags As Long
Public PreservesNames As Boolean
Public CaseSensitiveSearch As Boolean
Public SupportsUnicodeOnDisk As Boolean
Public SupportsPersistentAcls As Boolean
Public SupportsFileCompression As Boolean
Public VolumeIsCompressed As Boolean
Public SupportsEncryption As Boolean
Public SupportsObjectIds As Boolean
Public SupportsReparsePoints As Boolean
Public SupportsSparseFiles As Boolean
Public SupportsVolumeQuotas As Boolean
Public IsEmpty As Boolean
' Load the information for this drive.
' Parameter drive_name should be a single letter.
Public Sub Initialize(ByVal drive_name As String)
Dim volume_name As String
Dim file_system_name As String
Dim info_status As Long
Dim file_system_flags As Long

    ' Format the drive name as in A:\.
    drive_name = UCase$(Left$(drive_name, 1) & ":\")

    ' Get the drive type.
    Me.DriveType = GetDriveType(drive_name)

    ' Initialize name buffers.
    volume_name = Space$(256)
    file_system_name = Space$(256)

    ' Get the volume information.
    info_status = GetVolumeInformation(drive_name, _
        volume_name, Len(volume_name), Me.SerialNumber, _
        Me.MaxComponentLength, file_system_flags, _
        file_system_name, Len(file_system_name))

    ' Set the return values.
    Me.VolumeName = CleanString(volume_name)
    Me.FileSystemName = CleanString(file_system_name)
    Me.SupportsLongFileNames = (Me.MaxComponentLength = 255)
    Me.FileSystemFlags = file_system_flags
    Me.PreservesNames = (file_system_flags And FILE_CASE_PRESERVED_NAMES)
    Me.CaseSensitiveSearch = (file_system_flags And FILE_CASE_SENSITIVE_SEARCH)
    Me.SupportsUnicodeOnDisk = (file_system_flags And FILE_UNICODE_ON_DISK)
    Me.SupportsPersistentAcls = (file_system_flags And FILE_PERSISTENT_ACLS)
    Me.SupportsFileCompression = (file_system_flags And FILE_FILE_COMPRESSION)
    Me.VolumeIsCompressed = (file_system_flags And FILE_VOLUME_IS_COMPRESSED)
    Me.SupportsEncryption = (file_system_flags And FILE_SUPPORTS_ENCRYPTION)
    Me.SupportsObjectIds = (file_system_flags And FILE_SUPPORTS_OBJECT_IDS)
    Me.SupportsReparsePoints = (file_system_flags And FILE_SUPPORTS_REPARSE_POINTS)
    Me.SupportsSparseFiles = (file_system_flags And FILE_SUPPORTS_SPARSE_FILES)
    Me.SupportsVolumeQuotas = (file_system_flags And FILE_VOLUME_QUOTAS)

    If (Me.DriveType = drivetype_REMOVABLE) Or _
       (Me.DriveType = drivetype_CDROM) _
    Then
        Me.IsEmpty = (info_status = 0)
    Else
        Me.IsEmpty = False
    End If
End Sub
' Truncate the string at a NULL character if it
' contains one and remove leading and trailing spaces.
Private Function CleanString(ByVal txt As String) As String
Dim pos As Integer

    pos = InStr(txt, vbNullChar)
    If pos > 0 Then txt = Left$(txt, pos - 1)
    CleanString = Trim$(txt)
End Function

' Return a string representing the drive information.
Public Function ToString() As String
Dim txt As String

    txt = "DriveType:" & vbTab
    Select Case Me.DriveType
        Case drivetype_UNKNOWN
            txt = txt & "Unknown"
        Case drivetype_NO_ROOT_DIR
            txt = txt & "No root directory"
        Case drivetype_REMOVABLE
            txt = txt & "Removable"
            
        Case drivetype_FIXED
            txt = txt & "Fixed"
        Case drivetype_REMOTE
            txt = txt & "Remote"
        Case drivetype_CDROM
            txt = txt & "CD ROM"
        Case drivetype_RAMDISK
            txt = txt & "Ramdisk"
        Case Else
            txt = txt & "Unknown"
    End Select
    txt = txt & vbCrLf
    
    txt = txt & "VolumeName: " & vbTab & Me.VolumeName & vbCrLf
    txt = txt & "FileSystemName: " & vbTab & Me.FileSystemName & vbCrLf
    txt = txt & "SerialNumber: " & vbTab & Me.SerialNumber & vbCrLf
    txt = txt & "MaxComponentLength: " & vbTab & Me.MaxComponentLength & vbCrLf
    txt = txt & "SupportsLongFileNames: " & vbTab & Me.SupportsLongFileNames & vbCrLf
    txt = txt & "FileSystemFlags: " & vbTab & Hex$(Me.FileSystemFlags) & vbCrLf
    txt = txt & "PreservesNames: " & vbTab & Me.PreservesNames & vbCrLf
    txt = txt & "CaseSensitiveSearch: " & vbTab & Me.CaseSensitiveSearch & vbCrLf
    txt = txt & "SupportsUnicodeOnDisk: " & vbTab & Me.SupportsUnicodeOnDisk & vbCrLf
    txt = txt & "SupportsPersistentAcls: " & vbTab & Me.SupportsPersistentAcls & vbCrLf
    txt = txt & "SupportsFileCompression: " & vbTab & Me.SupportsFileCompression & vbCrLf
    txt = txt & "VolumeIsCompressed: " & vbTab & Me.VolumeIsCompressed & vbCrLf
    txt = txt & "SupportsEncryption: " & vbTab & Me.SupportsEncryption & vbCrLf
    txt = txt & "SupportsObjectIds: " & vbTab & Me.SupportsObjectIds & vbCrLf
    txt = txt & "SupportsReparsePoints: " & vbTab & Me.SupportsReparsePoints & vbCrLf
    txt = txt & "SupportsSparseFiles: " & vbTab & Me.SupportsSparseFiles & vbCrLf
    txt = txt & "SupportsVolumeQuotas: " & vbTab & Me.SupportsVolumeQuotas & vbCrLf
    txt = txt & "IsEmpty: " & vbTab & Me.IsEmpty & vbCrLf
    ToString = txt
End Function

Public Function getType() As String
    Dim txt As String
    
    Select Case Me.DriveType
        Case drivetype_UNKNOWN
            txt = "Unknown"
        Case drivetype_NO_ROOT_DIR
            txt = "No root directory"
        Case drivetype_REMOVABLE
            txt = "Removable"
        Case drivetype_FIXED
            txt = "Fixed"
        Case drivetype_REMOTE
            txt = "Remote"
        Case drivetype_CDROM
            txt = "CD ROM"
        Case drivetype_RAMDISK
            txt = "Ramdisk"
        Case Else
            txt = "Unknown"
    End Select
    getType = txt
End Function


