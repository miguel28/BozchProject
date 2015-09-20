Attribute VB_Name = "Globals"
'========================================
' Force explicit variable declaration.
'========================================
Option Explicit

'==========================
'Global Variables
'==========================
Public machine As MachineClass
Public IOPortCom As IOPort
Public UseEmulator As Boolean

'==========================
'Global Functions
'==========================
Public Function InitializeProgram()
    UseEmulator = True
    
    'Create a Instance of the machine
    Set machine = New MachineClass
    machine.OpenSerialPorts
    LoadTelegramHeader
    
    'Config IO Port
    Set IOPortCom = New IOPort
  
    If UseEmulator = True Then frmPortEmulator.Show
End Function

'====================================================
' Reads a Simple File as Text
'====================================================
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

'====================================================
' Loads the telegram headers
'====================================================
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

'====================================================
' Creates random numbers for the telegrams events ID
'====================================================
Public Function CreateRandomEventNumber() As String
    Randomize
    Dim result As String
    result = Str(Int((Rnd * 1000) + 1))
    CreateRandomEventNumber = Replace(result, " ", "")
End Function

'====================================================
' Manages the Logs
'====================================================
Public Function AppendLog(msg As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile

    Dim filename As String
    filename = "Log_" & Format$(Now, "yyyy-mm-dd") & ".txt"
    Dim path As String
    path = App.path & "\logs\" & filename
    
    Dim result As String
    result = Dir$(path)
    
    If Dir$(path) = "" Then              ' Creates the file if not exist
        Open path For Output As #iFileNo
        Print #iFileNo, msg & " , Date: " & Format$(Now, "yyyy-mm-dd")
    Else
        Open path For Append As #iFileNo ' If the file exists then append the log
        Print #iFileNo, msg & " , Date: " & Format$(Now, "yyyy-mm-dd")
    End If
    Close #iFileNo
End Function
