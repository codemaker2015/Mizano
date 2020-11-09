Attribute VB_Name = "ModSignature"
Option Explicit

'file size to be scanned virus
Public FileSize As Long
'declare Virus Def & info
Public VSig() As VirusSig
Public VSInfo As VS_Info
'declare variable for scan reg extensions
Public intSettingRegOption As Integer
Public strScanRegExt As String
'for faster DoEvents
'new DataType for Virus Signature
Public Type VirusSig

    Name As String
    Type As String
    value As String
    Action As String
    ActtionVal As String
    
End Type

'new DataType for Virus Signature Info
Public Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetData Form Database"
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

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure PutData Form Database"
End Sub





