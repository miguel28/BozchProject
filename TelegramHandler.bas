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
