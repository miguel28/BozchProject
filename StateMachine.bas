Attribute VB_Name = "StateMachine"
Public StepNumber As Integer
'Step 1: Espera Parte En el sensor
'Step 2: Espera Escaner Parte

Public Function UpdateStateMachine()
    Select Case StepNumber
        Case 1
            frmMain.lblOperatorMsg.Caption = "Ponga Parte en el Sensor"
            IOPortCom.SetOutput 0, False
            IOPortCom.SetOutput 1, True
            IOPortCom.SetOutput 2, False
            IOPortCom.SetOutput 3, True
            
            If IOPortCom.GetInput(0) = True Then
                StepNumber = 2
            End If
        Case 2
            frmMain.lblOperatorMsg.Caption = "Escanee the numero de parte"
            IOPortCom.SetOutput 0, True
            IOPortCom.SetOutput 1, False
            IOPortCom.SetOutput 2, True
            IOPortCom.SetOutput 3, False
            
        Case 3

    End Select
End Function
