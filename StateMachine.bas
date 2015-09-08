Attribute VB_Name = "StateMachine"
Private TimerHandler As Long
Private StepNumber As Integer

Public Function StartStateMachine()
    'Reset the State Machine to Step 0
    StepNumber = 0
    
    'To start the timer:
    TimerHandler = SetTimer(0, 0, 20, AddressOf TimerProc)
End Function

Public Function StopStateMachine()
    'To stop the timer:
    KillTimer 0, TimerHandler
End Function

Private Sub TimerProc(ByVal hwnd As Long, _
                      ByVal lMsg As Long, _
                      ByVal lTimerID As Long, _
                      ByVal lTimer As Long)
    UpdateStateMachine
End Sub


Private Function UpdateStateMachine()
    Select Case StepNumber
        Case 0
            frmMain.lblOperatorMsg.Caption = "Inicializando"
            IOPortCom.SetOutput 0, True
            IOPortCom.SetOutput 1, False
            IOPortCom.SetOutput 2, True
            IOPortCom.SetOutput 3, False

            If IOPortCom.GetInput(1) = True And IOPortCom.GetInput(3) = True Then
                StepNumber = 1
            End If
            
        Case 1
            frmMain.lblOperatorMsg.Caption = "Ponga Parte en el Sensor"
            If IOPortCom.GetInput(0) = True Then
                IOPortCom.SetOutput 0, False
                IOPortCom.SetOutput 1, True
                IOPortCom.SetOutput 2, False
                IOPortCom.SetOutput 3, True
            
                StepNumber = 2
            End If
        Case 2
            frmMain.lblOperatorMsg.Caption = "Cerrando"
            If IOPortCom.GetInput(2) = True And IOPortCom.GetInput(4) = True Then
                StepNumber = 3
            End If
            
        Case 3
            frmMain.lblOperatorMsg.Caption = "Escanee the numero de parte"
            If ScannerAvailable = True Then
                Sleep 200
                PartNumber = ReadFromScanner
                frmMain.lblPartNumber.Caption = "Numero de Parte " & PartNumber
                StepNumber = 4
            End If
        Case 4
            frmMain.lblOperatorMsg.Caption = "Imprimiendo Etiqueta"
            PrintZebra PartNumber
            StepNumber = 5
        Case 5
            Sleep 3000
            frmMain.lblOperatorMsg.Caption = "Etiqueta Impresa"
            StepNumber = 6
    End Select
End Function
