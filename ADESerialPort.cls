VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADESerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public WithEvents Port As MSComm
Attribute Port.VB_VarHelpID = -1

Private FileConfig As String
Private PortNumber As Integer
Private Settings As String
Private InBuffer As Integer
Private OutBuffer As Integer

Public Event OnComm()

Private Sub Class_Initialize()
    Set Port = New MSComm
End Sub

Public Function GetAvailableBytes() As Integer
    GetAvailableBytes = Port.InBufferCount
End Function

Public Function ConfigurePort(configFile As String)
    ReadConfig (configFile) ' Lee el archivo de configuracion como texto
    ParseConfig             ' Convierte Los Datos Leidos en Configuracion del puerto
    
    Port.CommPort = PortNumber
    Port.Settings = Settings
    Port.InBufferSize = InBuffer
    Port.OutBufferSize = OutBuffer
End Function

Private Function ReadConfig(configFile As String)
    On Error GoTo Error

    Dim intFile As Integer
    Dim ExecutablePath As String
    Dim ConfigPath As String
    intFile = FreeFile
    
    ExecutablePath = App.path & "\"
    ConfigPath = ExecutablePath & configFile
    
    Open ConfigPath For Input As #intFile
    FileConfig = StrConv(InputB(LOF(intFile), intFile), vbUnicode)
    Close #intFile
    Exit Function
Error:
    MsgBox "Error No se Pudo Archivo de configuracion para puerto serial: " & configFile
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

Private Sub Port_OnComm()
    RaiseEvent OnComm
End Sub
