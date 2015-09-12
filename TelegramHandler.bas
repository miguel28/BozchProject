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
    
    parser.SetAttribute "identifier", machine.SerialNumber
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


Public Function SendPartProcessed()
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\partProcessed_request.xml"
    
    SetHeader parser
    
    parser.SetAttribute "identifier", machine.SerialNumber
    parser.SetAttribute "typeNo", machine.TypeNumber
    parser.SetAttribute "typeVar", machine.typeVar
    
    parser.SetAttribute "Field6_ChangeStatus", machine.Field6_ChangeStatus
    parser.SetAttribute "Field7_ProductionNumber", machine.Field7_ProductionNumber
    parser.SetAttribute "Field10_AudiPartNo", machine.Field10_AudiPartNo
    parser.SetAttribute "Field11_SW", machine.Field11_SW
    parser.SetAttribute "Field14_HW", machine.Field14_HW
    parser.SetAttribute "Field15_Qspec", machine.Field15_Qspec
    parser.SetAttribute "Field17_DMC", machine.Field17_DMC  ''' LEEER DEL SCANNER
    parser.SetAttribute "Field18_ReferenceNo", machine.Field18_RefNo
    parser.SetAttribute "ccsFazitstring", machine.ccsFazitstring

    frmMain.sockMES.SendData parser.Code
End Function

Public Function ReadPartProcessed() As Boolean
    Dim parser As XMLParser
    Set parser = New XMLParser
    
    parser.Code = machine.SocketData
    machine.SocketAvailable = False
    
    Dim returnCode As Integer
    Dim retcode As String
    
    retcode = parser.GetAttribute("returnCode")
    returnCode = Int(retcode)
    
    If returnCode = 0 Then
        ReadPartProcessed = True
    Else
        ReadPartProcessed = False
    End If
    
End Function

Public Function SendPLCChangeOver(model As String)
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\plcChangeOverStarted_request.xml"
    
    SetHeader parser

    parser.SetAttribute "typeNo", model
    parser.SetAttribute "typeVar", machine.typeVar
    
    frmMain.sockMES.SendData parser.Code
End Function

Public Function ReadPLCChangeOver() As Boolean
    Dim parser As XMLParser
    Set parser = New XMLParser
    
    parser.Code = machine.SocketData
    machine.SocketAvailable = False
    
    Dim returnCode As Integer
    Dim retcode As String
    
    retcode = parser.GetAttribute("returnCode")
    returnCode = Int(retcode)
    
    If returnCode = 0 Then
        ReadPLCChangeOver = True
    Else
        ReadPLCChangeOver = False
    End If
    
End Function
