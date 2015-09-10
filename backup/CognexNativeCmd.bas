Attribute VB_Name = "CognexNativeCmd"
Public Function SetOnline(online As Boolean) As String
    SetOnline = "SO"
    If online = True Then
        SetOnline = SetOnline & "1"
    Else
        SetOnline = SetOnline & "0"
    End If
    SetOnline = SetOnline & vbCrLf
End Function

Public Function LoadJob(jobName As String) As String
    LoadJob = "LF" & jobName & vbCrLf
End Function

Public Function Trigger() As String
    Trigger = "AC" & vbCrLf ' sw8 se8
End Function

Public Function GetCellValue(Col As String, Row As Integer) As String
    GetCellValue = "gv" & Col & Format(yourNumber, "000") & vbCrLf
End Function


