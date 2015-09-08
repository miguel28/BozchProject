Attribute VB_Name = "WindowsNative"
'==========================
'Windows Native Functions
'==========================
'Creates a timer with a handler ID
Public Declare Function SetTimer Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal nIDEvent As Long, _
     ByVal uElapse As Long, _
     ByVal lpTimerFunc As Long) As Long

'Destroys a Timer
Public Declare Function KillTimer Lib "user32" _
    (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'Sleep Current Thread Certain number of milliseconds
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

