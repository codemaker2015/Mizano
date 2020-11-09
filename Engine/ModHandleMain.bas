Attribute VB_Name = "ModHandleMain"
Option Explicit

'CRC32 variable

'file size to be scanned virus
Public FileSize As Long
'declare Virus Def & info
Public VSig() As VirusSig
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public VSInfo As VS_Info
'declare variable for scan reg extensions
Public intSettingRegOption As Integer
Public strScanRegExt As String
'for faster DoEvents
Declare Function GetInputState Lib "user32.dll" () As Long
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

Function nPath(mypath As String) As String
    If Right(mypath, 1) = "\" Then
       nPath = mypath
    Else
       nPath = mypath & "\"
    End If
End Function
Function MyWindowSys() As String
    Dim buff As String
    buff = String(255, 0)
    GetSystemDirectory buff, 255
    MyWindowSys = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

