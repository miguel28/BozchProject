Attribute VB_Name = "TelegramHandler"
'========================================
' Force explicit variable declaration.
'========================================
Option Explicit

Public Function SetHeader(parser As XMLParser)
    ' Header Of the Telegram is common to all telegrams
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
    parser.SetAttribute "typeVar", machine.typevar
    
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
    
    machine.ccsDutLabelPara1 = parser.GetAttribute("name=""ccsDutLabelPara1"" value")
    machine.ccsDutLabelPara2 = parser.GetAttribute("name=""ccsDutLabelPara2"" value")
    machine.ccsDutLabelPara3 = parser.GetAttribute("name=""ccsDutLabelPara3"" value")
    machine.ccsDutLabelPara4 = parser.GetAttribute("name=""ccsDutLabelPara4"" value")
    machine.ccsDutLabelPara5 = parser.GetAttribute("name=""ccsDutLabelPara5"" value")
    machine.ccsFazitstring = parser.GetAttribute("name=""ccsFazitstring"" value")
    
    If returnCode = 0 Then
        ReadPartReceive = True
    Else
        ReadPartReceive = False
    End If
    
End Function


Public Function SendPartProcessed()
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\partProcessed_request.xml"
    
    SetHeader parser
    
    parser.SetAttribute "identifier", machine.SerialNumber
    parser.SetAttribute "typeNo", machine.TypeNumber
    parser.SetAttribute "typeVar", machine.typevar
    
    parser.SetAttribute "ccsDutLabelPara1", machine.ccsDutLabelPara1
    parser.SetAttribute "ccsDutLabelPara2", machine.ccsDutLabelPara2
    parser.SetAttribute "ccsDutLabelPara3", machine.ccsDutLabelPara3
    parser.SetAttribute "ccsDutLabelPara4", machine.ccsDutLabelPara4
    parser.SetAttribute "ccsDutLabelPara5", machine.ccsDutLabelPara5
    parser.SetAttribute "ccsFazitstring", machine.ccsFazitstring
    parser.SetAttribute "Field17_DMC", machine.Field17_DMC  ''' LEEER DEL SCANNER

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

Public Function SendPLCChangeOver(model As String, typevar As String)
    Dim parser As XMLParser
    Set parser = New XMLParser
    parser.Load "xmls\plcChangeOverStarted_request.xml"
    
    SetHeader parser

    parser.SetAttribute "typeNo", model
    parser.SetAttribute "typeVar", typevar
    
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
    
    machine.LabelType = parser.GetAttribute("name=""LabelType"" value")
    
    machine.ManufacturerNo = parser.GetAttribute("name=""ManufacturerNo"" value")
    machine.PlantNo = parser.GetAttribute("name=""PlantNo"" value")
    machine.ManufacturerCode = parser.GetAttribute("name=""ManufacturerCode"" value")
    machine.DMCversion = parser.GetAttribute("name=""DMCversion"" value")
    machine.NumberPCB = parser.GetAttribute("name=""NumberPCB"" value")
    machine.DMCfixedUnitNo = parser.GetAttribute("name=""DMCfixedUnitNo"" value")
    
    If returnCode = 0 Then
        ReadPLCChangeOver = True
    Else
        ReadPLCChangeOver = False
    End If
    
End Function
