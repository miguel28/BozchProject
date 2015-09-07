Attribute VB_Name = "Globals"
Option Explicit
'==========================
'Global Variables
'==========================
Public IOPortCom As IOPort
Public UseEmulator As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function AddAlarmMessage(msg As String)
    frmAlarms.txtAlarms.text = frmAlarms.txtAlarms.text & msg & vbCrLf
End Function

Public Function ReadFromScanner() As String
    If UseEmulator = True Then
        
    Else
        ReadFromScanner = frmMain.comScanner.Input
    End If

End Function

