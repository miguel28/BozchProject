Attribute VB_Name = "BitOperations"
'==============================================
' Bitwise Operands
'==============================================
' Because the Visual Basic 6 is too old as
' a fuck. Need the the bitwise implementations
' This is for use only of the IO Emulator.
'==============================================

'========================================
' Force explicit variable declaration.
'========================================
Option Explicit

Private OnBits(0 To 31) As Long

Public Function LShiftLong(ByVal value As Long, _
    ByVal Shift As Integer) As Long
  
    MakeOnBits
    
    If Shift > 0 Then
      If (value And (2 ^ (31 - Shift))) Then GoTo OverFlow
      LShiftLong = ((value And OnBits(31 - Shift)) * (2 ^ Shift))
      Exit Function
    Else
        LShiftLong = value
        Exit Function
    End If

OverFlow:
  
    LShiftLong = ((value And OnBits(31 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H80000000
  
End Function

Public Function RShiftLong(ByVal value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    
    If Shift > 0 Then
      If (value And &H80000000) Then hi = &H40000000
    
      RShiftLong = (value And &H7FFFFFFE) \ (2 ^ Shift)
      RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
    Else
        RShiftLong = value
    End If
      
End Function
 
Private Sub MakeOnBits()
    Dim j As Integer, _
        v As Long
  
    For j = 0 To 30
  
        v = v + (2 ^ j)
        OnBits(j) = v
  
    Next j
  
    OnBits(j) = v + &H80000000

End Sub

