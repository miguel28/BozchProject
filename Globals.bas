Attribute VB_Name = "Globals"
Option Explicit
'==========================
'Global Variables
'==========================
Public machine As MachineClass
Public IOPortCom As IOPort
Public UseEmulator As Boolean
Public ScannerAvailable As Boolean

'==========================
'Application Variables
'==========================
Public PartNumber As String

'==========================
'Global Functions
'==========================
Public Function AddAlarmMessage(msg As String)
    frmAlarms.txtAlarms.text = frmAlarms.txtAlarms.text & msg & vbCrLf
End Function

Public Function InitializeProgram()
    Set machine = New MachineClass
    
    'Config IO Port
    Set IOPortCom = New IOPort
    UseEmulator = True
    
    If UseEmulator = True Then frmPortEmulator.Show
End Function


Public Function ReadFromScanner() As String
    If UseEmulator = True Then
        ReadFromScanner = frmPortEmulator.txtScanner.text
    Else
        'ReadFromScanner = machine.comScanner.Input
    End If
    ScannerAvailable = False
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
        'machine.comZebra.Output = maker.Code
    End If
End Function
