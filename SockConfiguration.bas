Attribute VB_Name = "SockConfiguration"
Option Explicit
Private FileConfig As String
Private RemotePort As Integer
Private RemoteHost As String

Public Sub ConfigureSocket(Port As Winsock, configFile As String)
    RemotePort = 80
    RemoteHost = "127.0.0.1"
    
    ReadSockConfig (configFile) ' Lee el archivo de configuracion como texto
    ParseSockConfig             ' Convierte Los Datos Leidos en Configuracion del puerto
    
    Port.RemotePort = RemotePort
    Port.RemoteHost = RemoteHost
End Sub

Private Function ReadSockConfig(configFile As String)
    Dim intFile As Integer
    Dim ExecutablePath As String
    Dim ConfigPath As String
    intFile = FreeFile
    
    ExecutablePath = App.path & "\"
    ConfigPath = ExecutablePath & configFile
    
    Open ConfigPath For Input As #intFile
    FileConfig = StrConv(InputB(LOF(intFile), intFile), vbUnicode)
    Close #intFile
End Function

Private Function ParseSockConfig()
    Dim configs() As String
    Dim arrayLen As Integer
    Dim i As Integer
    Dim options() As String
    configs = Split(FileConfig, vbCrLf)
    
    arrayLen = UBound(configs) - LBound(configs) + 1

    For i = 0 To arrayLen - 1
        If InStr(configs(i), "RemotePort=") > 0 Then
            options = Split(configs(i), "=")
            RemotePort = Int(options(1))
        End If
        If InStr(configs(i), "RemoteHost=") > 0 Then
            options = Split(configs(i), "=")
            RemoteHost = options(1)
        End If
    Next i
End Function


