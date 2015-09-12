Attribute VB_Name = "TelegramHandler"
Public Function SetHeader(parser As XMLParser)
    parser.SetAttribute "lineNo", machine.lineNo
    parser.SetAttribute "statNo", machine.statNo
    parser.SetAttribute "statIdx", machine.statIdx
    parser.SetAttribute "fuNo", machine.fuNo
    parser.SetAttribute "workPos", machine.workPos
    parser.SetAttribute "toolPos", machine.toolPos
    parser.SetAttribute "processNo", machine.processNo
    parser.SetAttribute "processName", machine.processName
    parser.SetAttribute "application", machine.application
    
    ' Random Event
    parser.SetAttribute "eventId", CreateRandomEventNumber
    
End Function


Public Function SendPartReceive()
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\partReceived_request.xml"
    
    SetHeader parser
    
    parser.SetAttribute "identifier", machine.SerialNumber
    parser.SetAttribute "typeNo", machine.TypeNumber
    parser.SetAttribute "typeVar", machine.typeVar
    
    frmMain.sockMES.SendData parser.Code
End Function

Public Function ReadPartReceive() As Boolean
    Dim parser As XMLParser
    Set parser = New XMLParser
    
    parser.Code = machine.SocketData
    machine.SocketAvailable = False
    
    Dim returnCode As Integer
    Dim retcode As String
    
    retcode = parser.GetAttribute("returnCode")
    returnCode = Int(retcode)
    
    If returnCode = 0 Then
        ReadPartReceive = True
    Else
        ReadPartReceive = False
    End If
    
End Function

Public Function SendPartProcessingStart()
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\partProcessingStarted_request.xml"
    
    SetHeader parser
    
    parser.SetAttribute "identifierType", machine.SerialNumber
    parser.SetAttribute "typeNo", machine.TypeNumber
    parser.SetAttribute "typeVar", machine.typeVar
    
    frmMain.sockMES.SendData parser.Code
End Function

Public Function ReadPartProcessingStart() As Boolean
    Dim parser As XMLParser
    Set parser = New XMLParser
    
    parser.Code = machine.SocketData
    machine.SocketAvailable = False
    
    Dim returnCode As Integer
    Dim retcode As String
    
    retcode = parser.GetAttribute("returnCode")
    returnCode = Int(retcode)
    
    If returnCode = 0 Then
        ReadPartProcessingStart = True
        
        machine.Field6_ChangeStatus = parser.GetAttribute("Field6_ChangeStatus"" value")
        machine.Field7_ProductionNumber = parser.GetAttribute("Field7_ProductionNumber"" value")
        machine.Field10_AudiPartNo = parser.GetAttribute("Field10_AudiPartNo"" value")
        machine.Field11_SW = parser.GetAttribute("Field11_SW"" value")
        machine.Field14_HW = parser.GetAttribute("Field14_HW"" value")
        machine.Field15_Qspec = parser.GetAttribute("Field15_Qspec"" value")
        machine.Field17_DMC = parser.GetAttribute("Field17_DMC"" value")
        machine.Field18_RefNo = parser.GetAttribute("Field18_RefNo"" value")
        machine.ccsFazitstring = parser.GetAttribute("ccsFazitstring"" value")
    Else
        ReadPartProcessingStart = False
    End If
    
End Function
