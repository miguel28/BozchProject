VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPLMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Code As String


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

    Code = content
    Exit Function
Error:
    MsgBox "Error No se Pudo encontrar el archibo ZPL: " & file, vbCritical _
            + vbOKOnly, "Error de lectura de archivo"

End Function

Public Function SetAttribute(attrName As String, value As String)
    Code = Replace(Code, "%" & attrName & "%", value)
End Function

Public Function Begin()
    Code = "^XA"
End Function

Public Function Terminate()
    Code = Code & "^XZ"
End Function

Public Function SetOrigin(X As Integer, Y As Integer)
    Code = Code & "^FO" & Str(X) & "," & Str(Y)
End Function

Public Function SetFontSize(sX As Integer, sY As Integer)
    Code = Code & "^ADN," & Str(sX) & "," & Str(sY)
End Function

Public Function PutText(text As String)
    Code = Code & "^FD" & text & "^FS"
End Function

'80,Y,Y,N
Public Function BarCodeConfig(height As Integer, interp As String, textPos As String, orientation As String)
    Code = Code & "^BCN," & Str(height) & "," & interp & "," & textPos & "," & orientation
End Function

