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
    frmMain.txtSN.text = machine.SeriaNumber
    frmMain.txtcounter.text = Str(machine.GoodParts)
    frmMain.txtbadunits.text = Str(machine.BadParts)
End Function

Public Function UpdateStateMachine()
    Select Case StepNumber
        Case 1
            ' Checks that the Operator has selected the correct Information
            If frmMain.cboxParts.ListIndex < 0 Or Len(frmMain.txtTypeVar.text) <> 4 Then
                frmMain.txtOperador.text = "Seleccione Numero de Parte"
            Else
                machine.TypeNumber = frmMain.cboxParts.SelText
                machine.TypeVar = frmMain.txtTypeVar.text
                StepNumber = 2
            End If
        Case 2
            frmMain.txtOperador.text = "Escanee Numero de Serie"
            If machine.ScannerAvailable = True Then
                Dim result As Boolean
                ' Reads the data in the scanner and stores them
                ' in SerialNumber
                ' This function will return true in the data has
                ' has been correct
                
                result = machine.ReadFromScanner ' Read The scanner Data
                If result = True Then
                    StepNumber = 2
                    UpdateGUI
                Else
                    StepNumber = 1
                End If
                
            End If
        Case 2
            frmMain.txtOperador.text = "Enviando Respuesta a Sistema MES"

        Case 3
            frmMain.txtOperador.text = "Parte Aceptada"
            
        Case 4

    End Select
End Function
