VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Code As String

Private Sub Class_Initialize()
    Code = ""
End Sub

Public Function Load(file As String)
    On Error GoTo Error

    Dim intFile As Integer
    Dim ExecutablePath As String
    Dim path As String
    Dim content As String
    
    intFile = FreeFile
    
    ExecutablePath = App.path & "\"
    path = ExecutablePath & file
    
    Open path For Input As #intFile
    content = StrConv(InputB(LOF(intFile), intFile), vbUnicode)
    Close #intFile

    Code = Replace(content, "﻿", "")
    Exit Function
Error:
    MsgBox "Error No se Pudo encontrar el telegrama prototipo: " & file, vbCritical _
            + vbOKOnly, "Error de lectura de archivo"

End Function

Public Function SetAttribute(attrName As String, value As String)
    Code = Replace(Code, "%" & attrName & "%", value)
End Function

Public Function GetAttribute(attrName As String) As String
    Dim SearchTerm As String
    Dim Index1 As Long
    Dim Index2 As Long
    Dim subString As String
    Dim result As String
    
    SearchTerm = attrName & "="""
    Index1 = InStr(1, Code, SearchTerm, vbBinaryCompare)
    Index1 = Index1 + Len(SearchTerm)
    
    subString = Mid(Code, Index1, Len(Code) - Index1)
    Index2 = InStr(1, subString, """", vbBinaryCompare)
    
    result = Mid(Code, Index1, Index2 - 1)
    GetAttribute = result
End Function

Public Function GetValue(valName As String) As String
    Dim SearchTerm As String
    Dim SearchTerm2 As String
    Dim Index1 As Long
    Dim Index2 As Long
    Dim result As String
    
    SearchTerm = "<" & valName & ">"
    SearchTerm2 = "</" & valName & ">"
    Index1 = InStr(1, Code, SearchTerm, vbBinaryCompare)
    Index2 = InStr(1, Code, SearchTerm2, vbBinaryCompare)
    
    result = Mid(Code, Index1, Index2 - Index1)
    GetValue = result
End Function
