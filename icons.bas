Attribute VB_Name = "general"
Public drive As String
Public folder As String
Public file As String
Public folderpath(20) As String

Public Sub Message()
    Shield.Show
End Sub

Public Function GetIcon(extension As String) As String
    Select Case extension
        Case "jpeg", "png", "gif", "jpg": GetIcon = App.Path & "\images\img.jpg"
        Case "mp4", "avi", "3gp": GetIcon = App.Path & "\images\video.jpg"
        Case "mp3": GetIcon = App.Path & "\images\music.jpg"
        Case "docx", "doc": GetIcon = App.Path & "\images\word.jpg"
        Case "xlsx", "xls": GetIcon = App.Path & "\images\excel.jpg"
        Case "drive": GetIcon = App.Path & "\images\drive.jpg"
        Case "folder": GetIcon = App.Path & "\images\folder.jpg"
        Case Else: GetIcon = App.Path & "\images\other.jpg"
    End Select
End Function

Public Function getextension(FileName As String) As String
    Dim C As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        C = Mid(FileName, i, 1)
        If C = "." Then
            pos = i + 1
        End If
    Next i
    If pos = 0 Then
        getextension = ""
    Else
        getextension = Mid(FileName, pos, (Len(FileName) + 1 - pos))
    End If
End Function

Public Function getfilename2(FileName As String) As String
    Dim C As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        C = Mid(FileName, i, 1)
        If C = "\" Then
            pos = i + 1
            Exit For
        End If
    Next i
    If pos = 0 Then
        getfilename2 = ""
    Else
        getfilename2 = Mid(FileName, pos, (Len(FileName) + 1 - pos))
    End If
End Function

Public Function getfilename3(FileName As String) As String
    Dim C As String
    Dim pos As Integer
    For i = Len(FileName) To 2 Step -1
        C = Mid(FileName, i, 1)
        If C = "." Then
            pos = i + 1
        End If
    Next i
    If pos = 0 Then
        getfilename3 = ""
    Else
        getfilename3 = Mid(FileName, 1, pos - 2)
    End If
End Function
