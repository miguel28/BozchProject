Attribute VB_Name = "StateMachine"
Public StepNumber As Integer
'Step 1: Espera Parte En el sensor
'Step 2: Espera Escaner Parte

Public Function UpdateStateMachine()
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
                Sleep 100
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
