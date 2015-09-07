VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'IO Port Emulator
Public Outputs As Integer
Public Inputs As Integer

Public Function GetOutput(Address As Integer) As Boolean
    Dim value As Integer
    value = 0
    value = (RShiftLong(Outputs, Address) And &H1)
    If value > 0 Then
        GetOutput = True
    Else
        GetOutput = False
    End If
End Function

Public Function SetOutput(Address As Integer, value As Boolean)
    Dim val As Integer
    If value = True Then
        val = Outputs Or (LShiftLong(1, Address))
    Else
        val = Outputs And (Not LShiftLong(1, Address))
    End If

    Outputs = val
End Function

Public Function GetInput(Address As Integer) As Boolean
    Dim value As Integer
    value = 0
    value = (RShiftLong(Inputs, Address) And &H1)
    If value > 0 Then
        GetInput = True
    Else
        GetInput = False
    End If
End Function

Public Function SetInput(Address As Integer, value As Boolean)
    Dim val As Integer
    If UseEmulator = True Then ' Use en emulator Only
    
        If value = True Then
            val = Inputs Or (LShiftLong(1, Address))
        Else
            val = Inputs And (Not LShiftLong(1, Address))
        End If
    
        Inputs = val
    End If
End Function
