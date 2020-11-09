Attribute VB_Name = "Module1"
Public drive As String
Public folder As String
Public file As String

Public Function geticon(extension As String) As String
    Select Case extension
        Case "jpeg", "png", "gif", "jpg": geticon = ""
        Case "mp4", "avi", "3gp": geticon = ""
        Case "mp3": geticon = ""
        Case "docx", "doc": geticon = ""
        Case "xlsx", "xls": geticon = ""
        Case "drive": geticon = ""
        Case "folder": geticon = ""
    End Select
End Function

Public Function getextension(FileName As String) As String
    Dim c As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        c = Mid(FileName, i, 1)
        If c = "." Then
            pos = i + 1
        End If
    Next i
    getextension = Mid(FileName, pos, (Len(FileName) + 1 - pos))
End Function
