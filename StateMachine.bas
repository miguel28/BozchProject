Attribute VB_Name = "StateMachine"
Option Explicit
'==========================
' Private Variables
'==========================
Private TimerHandler As Long
Private StepNumber As Integer

'===================================================================
' State Machine Step Definitions
' Step 0 = Machine Inizialitation (Reinizialition)
' Step 1 = Vefify if the Operator has select the Typenumber and TypeVar
' Step 2 = Read Scaneer (Operator)
' Step 3 = Send PartReceived Telegram to MES
' Step 4 = Verifiying information of the Received Telegram MES
' Step 5 = Send to MES the telegram PartProcessing Start
' Step 6 = Refifiying information of the Received Telegram MES
' Step 7 = Print Label In Zebra Printer
'===================================================================

'==========================
' State Machine Control Start/Stop
'==========================
Public Function StartStateMachine()
    'Reset the State Machine to Step 0
    StepNumber = 1
    
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
    machine.TypeNumber = frmMain.cboxParts.SelText
End Function

Public Function UpdateStateMachine()
    Select Case StepNumber
        Case 1
            ' Checks that the Operator has selected the correct Information
            If frmMain.cboxParts.ListIndex < 0 Or Len(frmMain.txtTypeVar.text) <> 4 Then
                frmMain.txtOperador.text = "Seleccione Numero de Parte"
            Else
                machine.TypeNumber = frmMain.cboxParts.SelText
                machine.typeVar = frmMain.txtTypeVar.text
                StepNumber = 2
            End If
        Case 2
            frmMain.txtOperador.text = "Escanee Numero de Serie"
            If machine.comScanner.GetAvailableBytes > 8 Then
                Dim result As Boolean
                ' Reads the data in the scanner and stores them
                ' in SerialNumber
                ' This function will return true in the data has
                ' has been correct
                
                result = machine.ReadFromScanner ' Read The scanner Data
                If result = True Then
                    StepNumber = 3
                    UpdateGUI
                Else
                    StepNumber = 1
                End If
                
            End If
        Case 3
            frmMain.txtOperador.text = "Enviando Respuesta a Sistema MES"
            SendPartReceive
            StepNumber = 4
            
        Case 4
            frmMain.txtOperador.text = "Esperando Respuesta de MES"
            If machine.SocketAvailable Then
                If ReadPartReceive = True Then
                    StepNumber = 5
                Else
                    frmMain.txtOperador.text = "Error: esta pieza no debe ser procesda en esta estacion"
                    DoEvents
                    Sleep 3000
                    StepNumber = 1
                End If
            End If
            
        Case 5
            frmMain.txtOperador.text = "Requiriendo informacion de etiqueta de MES"
            SendPartProcessingStart
            StepNumber = 6
        Case 6
            frmMain.txtOperador.text = "Esperando Respuesta de MES"
            If machine.SocketAvailable Then
                If ReadPartProcessingStart = True Then
                    StepNumber = 7
                Else
                    frmMain.txtOperador.text = "Error: esta pieza no debe ser procesda en esta estacion"
                    Sleep 3000
                    StepNumber = 1
                End If
            End If
        Case 7
            frmMain.txtOperador.text = "Imprimiendo Etiqueta"
            machine.PrintZebra
            StepNumber = 8
            DoEvents
            Sleep 3000
        Case 8
            frmMain.txtOperador.text = "Pegue etiqueta en el producto y Escanee el codigo de barras"
    End Select
End Function
