Attribute VB_Name = "ModRefresh"
Public Function HitDatabase()
    Dim i As Integer
    Dim vCount As Integer
    
    ATVGenerator.lstVirus.ListItems.Clear
    vCount = 0
    For i = 0 To UBound(VSig)
        vCount = vCount + 1
        ATVGenerator.lstVirus.ListItems.Add , , VSig(i).Name, , 1
    Next i
    ATVGenerator.lblVirusCount.Caption = VSInfo.VirusCount
End Function

Public Function Dire(str_File As String) As Boolean
    On Error GoTo err
    Dire = Not (Dir(str_File) = "" And Dir(str_File, vbHidden) = "" And Dir(str_File, vbSystem) = "" And Dir(str_File, vbNormal) = "")
Exit Function
err:
    Dire = False
    err.Clear
End Function
