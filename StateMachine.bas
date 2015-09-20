Attribute VB_Name = "StateMachine"
'========================================
' Force explicit variable declaration.
'========================================
Option Explicit

'==========================
' Private Variables
'==========================
Private TimerHandler As Long
Private StepNumber As Integer
Private TimeoutCounter As Integer
Private TimedOutResponse As Boolean
Private Const TimeoutLimit As Integer = 10000
'===================================================================
' State Machine Step Definitions
' Step 0 = Machine Inizialitation (Reinizialition)
'          Check the connection of MES System
' Step 1 = Vefify if the Operator has select the Typenumber and TypeVar
' Step 2 = Read Scaneer (Operator)
' Step 3 = Send PartReceived Telegram to MES
' Step 4 = Get from MES the Label Information (NOTE: Go to Step 7)
' Step 5 = (REMOVED) 'couse PartReceoved Start Procesing Teletram deletion
' Step 6 = (REMOVED) 'couse PartReceoved Start Procesing Teletram deletion
' Step 7 = Print Label In Zebra Printer and waits 3 Seconds
' Step 8 = Reads the Data Matrix Code
' Step 9 = Send partProcessed Telegram to MES
' Step 10= Read response of the MES if the part was process correctly
'          Finaly returns to Step 0
' Step 11= MES TIMEOUT
'===================================================================

'==========================
' State Machine Control Start/Stop
'==========================
Public Function StartStateMachine()
    'Reset the State Machine to Step 0
    StepNumber = 0
    
    'To start the timer:
    'TimerHandler = SetTimer(0, 0, 200, AddressOf TimerProc)
    frmMain.tmrUpdateStateMachine.Enabled = True
End Function

Public Function StopStateMachine()
    'To stop the timer:
    'KillTimer 0, TimerHandler
    frmMain.tmrUpdateStateMachine.Enabled = False
End Function

'==========================
' State Machine Procedures
'==========================

Private Sub TimerProc(ByVal hwnd As Long, _
                      ByVal lMsg As Long, _
                      ByVal lTimerID As Long, _
                      ByVal lTimer As Long)
    UpdateStateMachine
End Sub

Public Function UpdateGUI()
    frmMain.txtSN.text = machine.SerialNumber
    frmMain.txtcounter.text = Str(machine.GoodParts)
    frmMain.txtbadunits.text = Str(machine.BadParts)
    
    frmMain.txtModel.text = machine.TypeNumber
    frmMain.txtTypeVar.text = machine.typevar
    
    If machine.SocketConnected = True Then
        frmMain.lblAMS.Visible = False
    Else
        frmMain.lblAMS.Visible = True
    End If
    
    If StepNumber < 3 Then
        frmMain.btnUtils.Enabled = True
        frmMain.btnChangeModel.Enabled = True
    Else
        frmMain.btnUtils.Enabled = False
        frmMain.btnChangeModel.Enabled = False
    End If
    
End Function

Private Function UpdateMsg(msg As String)
    frmMain.txtOperador.text = msg
End Function

Private Function WatchdogTimeOut()
    If StepNumber = 4 Or StepNumber = 10 Then
        TimeoutCounter = TimeoutCounter + 200
        If TimeoutCounter >= TimeoutLimit Then
            TimedOutResponse = True
            StepNumber = 11
        End If
    Else
        TimeoutCounter = 0
        TimedOutResponse = False
    End If

End Function

Public Function UpdateStateMachine()
    UpdateGUI
    WatchdogTimeOut
    
    Select Case StepNumber
    Case 0
        ' Checks that the Operator has selected the correct Information
        UpdateMsg "Se necesita estar Conectado a sistemas MES"
        If machine.SocketConnected Then
            StepNumber = 1
        End If
    Case 1
        ' Checks that the Operator has selected the correct Information
        If Len(machine.TypeNumber) > 0 Then
            StepNumber = 2
        Else
            UpdateMsg "Seleccione Numero de Parte"
            machine.SerialNumber = ""
        End If
    Case 2
        UpdateMsg "Escanee Numero de Serie"
        If machine.comScanner.GetAvailableBytes >= machine.SerialNumberLength Then
            Dim result As Boolean
            ' Reads the data in the scanner and stores them
            ' in SerialNumber
            ' This function will return true in the data has
            ' has been correct
                
            result = machine.ReadFromScanner ' Read The scanner Data
            If result = True Then
                AppendLog ("Read Part From Part: " & machine.SerialNumber)
                StepNumber = 3
                UpdateGUI
            Else
                StepNumber = 0
            End If
                
        End If
    Case 3
        UpdateMsg "Enviando PartReceived a Sistema MES"
        SendPartReceive
        StepNumber = 4
            
    Case 4
        UpdateMsg "Esperando Respuesta PartReceived de MES"
        If machine.SocketAvailable Then
            If ReadPartReceive = True Then
                StepNumber = 7
            Else
                UpdateMsg "Error: esta pieza no debe ser procesda en esta estacion"
                DoEvents
                Sleep 3000
                StepNumber = 0
            End If
        End If
    Case 7
        UpdateMsg "Imprimiendo Etiqueta"
        machine.PrintZebra
        StepNumber = 8
        DoEvents
        Sleep 3000
    Case 8
        UpdateMsg "Pegue etiqueta en el producto y Escanee el codigo de barras"
        If machine.comScanner.GetAvailableBytes >= machine.DMCNumberLength Then
            Dim result2 As Boolean
                
            result2 = machine.ReadDMC ' Read The scanner Data
            If result2 = True Then
                '====================================
                ' If InStr(1, machine.DMC, machine.Field17_DMC) >= 0 Then
                '    StepNumber = 9
                'Else
                '    frmMain.txtOperador.text = "Error: Los DMC no corresponden"
                '    Sleep 3000
                '    StepNumber = 0
                'End If
                '====================================
            Dim Index As Integer
                Index = InStr(1, machine.DMC, machine.SerialNumber)
                If Index >= 1 Then
                    AppendLog ("Valid DMC Verification")
                    StepNumber = 9
                Else
                    AppendLog ("Invalid DMC Verification")
                    UpdateMsg "Error: Los DMC no corresponden"
                    machine.BadParts = machine.BadParts + 1
                    UpdateGUI
                    DoEvents
                    Sleep 3000
                    StepNumber = 0
                End If
            Else
                AppendLog ("Invalid DMC Verification")
                UpdateMsg "Error: Al Leer DMC"
                Sleep 3000
                StepNumber = 0
            End If
                
        End If
    Case 9
        UpdateMsg "Envaindo PartProcessed a MES"
        SendPartProcessed
        StepNumber = 10
    Case 10
        UpdateMsg "Esperando Respuesta PartProcessed de MES"
        If machine.SocketAvailable Then
            If ReadPartProcessed = True Then
                AppendLog ("Good Part")
                UpdateMsg "Pieza Processada Correctamente"
                machine.GoodParts = machine.GoodParts + 1
                UpdateGUI
                DoEvents
                Sleep 3000
                    
                StepNumber = 0 ' Termina ciclo empieza e nuevo
            Else
                AppendLog ("Bad Part")
                UpdateMsg "Error: Los DMC no corresponden"
                machine.BadParts = machine.BadParts + 1
                UpdateGUI
                DoEvents
                Sleep 3000
                StepNumber = 0
            End If
        End If
    Case 11
        AppendLog ("MES TIMEOUT Error: the response exceded 10 Seconds")
        UpdateMsg "Error: Error Connexion al MES"
        machine.BadParts = machine.BadParts + 1
        UpdateGUI
        DoEvents
        Sleep 3000
        StepNumber = 0
    End Select
End Function
