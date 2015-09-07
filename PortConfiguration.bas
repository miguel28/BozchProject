Attribute VB_Name = "PortConfiguration"
Option Explicit
Private FileConfig As String
Private PortNumber As Integer
Private Settings As String
Private InBuffer As Integer
Private OutBuffer As Integer

Public Sub ConfigurePort(Port As MSCommLib.MSComm, configFile As String)
    ReadConfig (configFile) ' Lee el archivo de configuracion como texto
    ParseConfig             ' Convierte Los Datos Leidos en Configuracion del puerto
    
    PortNumber = 1
    Settings = "9600,n,8,1"
    InBuffer = 1024
    OutBuffer = 1024
    
    Port.CommPort = PortNumber
    Port.Settings = Settings
    Port.InBufferSize = InBuffer
    Port.OutBufferSize = OutBuffer
End Sub

Private Function ReadConfig(configFile As String)
    Dim intFile As Integer
    Dim ExecutablePath As String
    Dim ConfigPath As String
    intFile = FreeFile
    
    ExecutablePath = App.Path & "\"
    ConfigPath = ExecutablePath & configFile
    
    Open ConfigPath For Input As #intFile
    FileConfig = StrConv(InputB(LOF(intFile), intFile), vbUnicode)
    Close #intFile
End Function

Private Function ParseConfig()
    Dim configs() As String
    Dim arrayLen As Integer
    Dim i As Integer
    Dim options() As String
    configs = Split(FileConfig, vbCrLf)
    
    arrayLen = UBound(configs) - LBound(configs) + 1
    
    For i = 0 To arrayLen - 1
        If InStr(configs(i), "PortNumber=") > 0 Then
            options = Split(configs(i), "=")
            PortNumber = Int(options(1))
        End If
        If InStr(configs(i), "Settings=") > 0 Then
            options = Split(configs(i), "=")
            Settings = options(1)
        End If
        
        If InStr(configs(i), "InBufferSize=") > 0 Then
            options = Split(configs(i), "=")
            InBuffer = Int(options(1))
        End If
        If InStr(configs(i), "OutBufferSize=") > 0 Then
            options = Split(configs(i), "=")
            OutBuffer = options(1)
        End If
    Next i
End Function

