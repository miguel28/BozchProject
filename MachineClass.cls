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
Public WithEvents comZebra As ADESerialPort
Attribute comZebra.VB_VarHelpID = -1
Public WithEvents comScanner As ADESerialPort
Attribute comScanner.VB_VarHelpID = -1
Public WithEvents comCognex As ADESerialPort
Attribute comCognex.VB_VarHelpID = -1
Public WithEvents comHandScanner As ADESerialPort
Attribute comHandScanner.VB_VarHelpID = -1

Private Sub Class_Initialize()
    'Create Ports
    Set comZebra = New ADESerialPort
    Set comScanner = New ADESerialPort
    Set comCognex = New ADESerialPort
    Set comHandScanner = New ADESerialPort

    'Load Serial COM Configuration of config files
    comZebra.ConfigurePort "ZebraPort.ini"
    comScanner.ConfigurePort "ScannerPort.ini"
    comCognex.ConfigurePort "CognexPort.ini"
    comHandScanner.ConfigurePort "HandScannerPort.ini"
End Sub

'==========================
'Port Receiver Events
'==========================

Private Sub comZebra_OnComm()
'
End Sub
