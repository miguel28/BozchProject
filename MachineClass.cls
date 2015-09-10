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
    'Create Ports
    Set comZebra = New ADESerialPort
    Set comScanner = New ADESerialPort

    'Load Serial COM Configuration of config files
    comZebra.ConfigurePort "config\ZebraPort.ini"
    comScanner.ConfigurePort "config\ScannerPort.ini"
End Sub

'==========================
'Port Receiver Events
'==========================

Private Sub comZebra_OnComm()
'
End Sub

Private Sub comScanner_OnComm()
'
End Sub

