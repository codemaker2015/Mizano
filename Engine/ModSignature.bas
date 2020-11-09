Attribute VB_Name = "ModSignature"
Option Explicit
Public jmlProcess As Integer
Public sCRC As String
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'Check if a path or file exists
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Checks if a folder or file exists
Public Function FileorFolderExists(FolderOrFilename As String) As Boolean
    If PathFileExists(FolderOrFilename) = 1 Then
        FileorFolderExists = True
    ElseIf PathFileExists(FolderOrFilename) = 0 Then
        FileorFolderExists = False
    End If
End Function

Public Sub ReadSig()

    Dim F As Long
    On Error GoTo Trap_Error
    F = FreeFile
    Open App.Path & "\Signature.ATVGuard" For Binary Access Read As #F
        Get #F, , VSInfo
        ReDim VSig(VSInfo.VirusCount - 1) As VirusSig
        Dim i As Integer
        For i = 0 To VSInfo.VirusCount - 1
            Get #F, , VSig(i)
        Next
    Close #F

   On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetData of Form Database"
End Sub

Public Sub WriteSig(ByRef vs As VirusSig)
    
    Dim F As Long
    On Error GoTo Trap_Error
    F = FreeFile
    
    Dim i As Long
    
    'add 1 item into array
    ReDim Preserve VSig(UBound(VSig) + 1) As VirusSig
    VSig(UBound(VSig)).Name = vs.Name
    VSig(UBound(VSig)).Type = vs.Type
    VSig(UBound(VSig)).value = vs.value
    
    'add 1 for count
    VSInfo.VirusCount = UBound(VSig) + 1
    VSInfo.LastUpdate = Format(Date, "dd/mmmm/yyyy")
    Open App.Path & "\signature.ATVGuard" For Binary Access Write As #F
        Put #F, , VSInfo
        For i = 0 To UBound(VSig)
            Put #F, , VSig(i)
        Next
    Close #F

   On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure PutData of Form Database"
End Sub




