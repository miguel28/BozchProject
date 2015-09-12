VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MachineClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================
'MES Telegram Headers
'==========================
Public lineNo As String
Public statNo As String
Public statIdx As String
Public fuNo As String
Public workPos As String
Public toolPos As String
Public processNo As String
Public processName As String
Public application As String

'==========================
'Application Variables
'==========================
Public ScannerAvailable As Boolean
Public SocketAvailable As Boolean
Public SocketData As String

Public SerialNumber As String
Public TypeNumber As String
Public typeVar As String
Public GoodParts As Integer
Public BadParts As Integer

'==========================
'Application Constants
'==========================
Private Const SerialNumberLength As Integer = 10 'This is the lenght spected to read from scanner

'==========================
'Definitions Of Ports
'==========================
Public WithEvents comZebra As ADESerialPort
Attribute comZebra.VB_VarHelpID = -1
Public WithEvents comScanner As ADESerialPort
Attribute comScanner.VB_VarHelpID = -1

'==========================
'Class Constructor
'==========================
Private Sub Class_Initialize()
    'Initialize Variables
    SerialNumber = ""
    TypeNumber = ""
    typeVar = ""
    GoodParts = 0
    BadParts = 0
    
    'Create Ports
    Set comZebra = New ADESerialPort
    Set comScanner = New ADESerialPort

    'Load Serial COM Configuration of config files
    comZebra.ConfigurePort "config\ZebraPort.ini"
    comScanner.ConfigurePort "config\ScannerPort.ini"
End Sub


Public Function ReadFromScanner() As Boolean
    Dim data As String
    Dim i As Integer
    
    ScannerAvailable = False
    i = 0
    
    If UseEmulator = True Then
        data = frmPortEmulator.txtScanner.text ' Read Emulated Data
    Else
        Do While machine.comScanner.Port.InBufferCount = SerialNumberLength And i < 15
            Sleep 10
            i = i + 1
        Loop
        data = machine.comScanner.Port.Input        ' Read From real Scanner Serial Port
    End If

    ' This if check if the for got the correct number od bytes from the scanner
    If Len(data) <> SerialNumberLength Then ' If the Received Data is diferent of spected
        MsgBox "Error al leer los datos del escaner se esperaban " & _
                Str(SerialNumberLength) & " pero se recibieron " & Str(Len(data)) & _
                ". Por favor vuelva a leer el archivo.", _
                vbCritical + vbOKOnly, "Error de lectura de archivo"
                
        ReadFromScanner = False             ' Return error
        SerialNumber = ""
    Else
        ReadFromScanner = True
        SerialNumber = data
    End If
    
End Function

Public Function PrintZebra(Datos As String)
    Dim maker As ZPLMaker
    Set maker = New ZPLMaker
    
    maker.Begin
    maker.SetOrigin 50, 50
    maker.SetFontSize 30, 7
    maker.BarCodeConfig 80, "Y", "Y", "N"
    maker.PutText Datos
    maker.Terminate

    If UseEmulator = True Then
        frmPortEmulator.txtZPL.text = frmPortEmulator.txtZPL.text & maker.Code & vbCrLf
    Else
        machine.comZebra.Port.Output = maker.Code
    End If
End Function

'==========================
'Port Receiver Events
'==========================

Private Sub comZebra_OnComm()
'
End Sub

Private Sub comScanner_OnComm()
    ScannerAvailable = False
End Sub

