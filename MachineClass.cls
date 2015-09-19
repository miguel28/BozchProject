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

Public ccsDutLabelPara1 As String
Public ccsDutLabelPara2 As String
Public ccsDutLabelPara3  As String
Public ccsDutLabelPara4  As String
Public ccsDutLabelPara5 As String
Public ccsFazitstring  As String

Public Field17_DMC As String

Public LabelType As String

Public DMC As String
'==========================
'Application Variables
'==========================
Public ScannerAvailable As Boolean
Public SocketAvailable As Boolean
Public SocketData As String
Public SocketConnected As Boolean

Public SerialNumber As String
Public TypeNumber As String
Public typevar As String
Public GoodParts As Integer
Public BadParts As Integer

'==========================
'Application Constants
'==========================
Public SerialNumberLength As Integer 'This is the lenght spected to read from scanner
Public DMCNumberLength As Integer    'This is the lenght spected to read from scanner

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
    SerialNumberLength = 9
    DMCNumberLength = 9
    
    SerialNumber = ""
    TypeNumber = ""
    typevar = ""
    GoodParts = 0
    BadParts = 0
    
    'Create Ports
    Set comZebra = New ADESerialPort
    Set comScanner = New ADESerialPort

    'Load Serial COM Configuration of config files
    comScanner.ConfigurePort "config\ScannerPort.ini"
    comZebra.ConfigurePort "config\ZebraPort.ini"
    
    If UseEmulator = False Then
        comZebra.Port.PortOpen = True
        comScanner.Port.PortOpen = True
    End If
    
    
    
End Sub

Public Function ReadFromScanner() As Boolean
    Dim data As String
    Dim i As Integer
    
    ScannerAvailable = False
    i = 0
    
    If UseEmulator = True Then
        data = frmPortEmulator.txtScanner.text ' Read Emulated Data
    Else
        Do While machine.comScanner.Port.InBufferCount = SerialNumberLength And i < 25
            Sleep 10
            i = i + 1
        Loop
        data = machine.comScanner.Port.Input        ' Read From real Scanner Serial Port
    End If
    
    data = Replace(data, vbCrLf, "")
    
    ' This if check if the for got the correct number od bytes from the scanner
    If Len(data) < SerialNumberLength Then ' If the Received Data is diferent of spected
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

Public Function ReadDMC() As Boolean
    Dim data As String
    Dim i As Integer
    
    ScannerAvailable = False
    i = 0
    
    If UseEmulator = True Then
        data = frmPortEmulator.txtScanner.text ' Read Emulated Data
    Else
        Do While machine.comScanner.Port.InBufferCount = DMCNumberLength And i < 25
            Sleep 10
            i = i + 1
        Loop
        data = machine.comScanner.Port.Input        ' Read From real Scanner Serial Port
    End If
    
    data = Replace(data, vbCrLf, "")
    
    ' This if check if the for got the correct number od bytes from the scanner
    If Len(data) < SerialNumberLength Then ' If the Received Data is diferent of spected
        MsgBox "Error al leer los datos del escaner se esperaban " & _
                Str(DMCNumberLength) & " pero se recibieron " & Str(Len(data)) & _
                ". Por favor vuelva a leer el archivo.", _
                vbCritical + vbOKOnly, "Error de lectura de archivo"
                
        ReadDMC = False             ' Return error
        DMC = ""
    Else
        ReadDMC = True
        DMC = data
    End If
    
End Function

Public Function PrintZebra()
    Dim maker As ZPLMaker
    Set maker = New ZPLMaker
    
    maker.Load "zpls\testimprimr.txt"
    maker.SetAttribute "numero", SerialNumber
    
    If UseEmulator = True Then
        frmPortEmulator.txtZPL.text = frmPortEmulator.txtZPL.text & maker.Code & vbCrLf
    Else
        machine.comZebra.Port.Output = maker.Code
    End If
End Function

Public Function PrintTestZebra()
    Dim maker As ZPLMaker
    Set maker = New ZPLMaker
    
    maker.Load "zpls\testimprimr.txt"
    maker.SetAttribute "numero", "PRUEBA"

    If UseEmulator = True Then
        frmPortEmulator.txtZPL.text = frmPortEmulator.txtZPL.text & maker.Code & vbCrLf
    Else
        machine.comZebra.Port.Output = maker.Code
    End If
End Function
