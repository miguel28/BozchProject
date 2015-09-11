Attribute VB_Name = "Globals"
Option Explicit
'==========================
'Global Variables
'==========================
Public machine As MachineClass
Public IOPortCom As IOPort
Public UseEmulator As Boolean


'==========================
'Application Variables
'==========================
Public PartNumber As String

'==========================
'Global Functions
'==========================
Public Function InitializeProgram()
    'Create a Instance of the machine
    Set machine = New MachineClass
    
    'Config IO Port
    Set IOPortCom = New IOPort
    
    UseEmulator = True
    If UseEmulator = True Then frmPortEmulator.Show
End Function

Public Function ReadTextFile(file As String)
    On Error GoTo Error

    Dim intFile As Integer
    Dim ExecutablePath As String
    Dim path As String
    Dim content As String
    
    intFile = FreeFile
    
    ExecutablePath = App.path & "\"
    path = ExecutablePath & file
    
    Open path For Input As #intFile
    content = StrConv(InputB(LOF(intFile), intFile), vbUnicode)
    Close #intFile

    ReadTextFile = content
    Exit Function
Error:
    MsgBox "Error No se Pudo encontrar el archivo: " & file

End Function

Public Function LoadPartNumbers(cbox As ComboBox)
    Dim Numbers As String
    Numbers = ReadTextFile("config\PartNumbers.ini")
    
    Dim lines() As String
    Dim arrayLen As Integer
    
    lines = Split(Numbers, vbCrLf)
    arrayLen = UBound(lines) - LBound(lines) + 1
    
    Dim i As Integer
    
    For i = 0 To arrayLen - 1
        cbox.AddItem lines(i)
    Next i
End Function
