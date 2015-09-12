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
    
    LoadTelegramHeader
    
    'Config IO Port
    Set IOPortCom = New IOPort
    
    UseEmulator = False
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
    
    Dim Lines() As String
    Dim arrayLen As Integer
    
    Lines = Split(Numbers, vbCrLf)
    arrayLen = UBound(Lines) - LBound(Lines) + 1
    
    Dim i As Integer
    
    For i = 0 To arrayLen - 1
        cbox.AddItem Lines(i)
    Next i
End Function

Public Function LoadTelegramHeader()
    Dim Lines As String
    Lines = ReadTextFile("config\TelegramHeader.ini")
    Dim configs() As String
    Dim arrayLen As Integer
    Dim i As Integer
    Dim options() As String
    configs = Split(Lines, vbCrLf)
    
    arrayLen = UBound(configs) - LBound(configs) + 1
    
    For i = 0 To arrayLen - 1
        If InStr(configs(i), "lineNo=") > 0 Then
            options = Split(configs(i), "=")
            machine.lineNo = options(1)
        End If
        
        If InStr(configs(i), "statNo=") > 0 Then
            options = Split(configs(i), "=")
            machine.statNo = options(1)
        End If
        
        If InStr(configs(i), "statIdx=") > 0 Then
            options = Split(configs(i), "=")
            machine.statIdx = options(1)
        End If
        
        If InStr(configs(i), "fuNo=") > 0 Then
            options = Split(configs(i), "=")
            machine.fuNo = options(1)
        End If
        
        If InStr(configs(i), "workPos=") > 0 Then
            options = Split(configs(i), "=")
            machine.workPos = options(1)
        End If
        
        If InStr(configs(i), "toolPos=") > 0 Then
            options = Split(configs(i), "=")
            machine.toolPos = options(1)
        End If
        
        If InStr(configs(i), "processNo=") > 0 Then
            options = Split(configs(i), "=")
            machine.processNo = options(1)
        End If
        
        If InStr(configs(i), "processName=") > 0 Then
            options = Split(configs(i), "=")
            machine.processName = options(1)
        End If
        
        If InStr(configs(i), "application=") > 0 Then
            options = Split(configs(i), "=")
            machine.application = options(1)
        End If
    Next i
End Function

Public Function CreateRandomEventNumber() As String
    Randomize
    Dim result As String
    result = Str(Int((Rnd * 1000) + 1))
    CreateRandomEventNumber = Replace(result, " ", "")
End Function
