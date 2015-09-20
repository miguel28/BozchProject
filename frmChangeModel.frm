VERSION 5.00
Begin VB.Form frmChangeModel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Modelo"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTypeVar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "0000"
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton btnChangeModel 
      Caption         =   "Solicitar Cambio de Modelo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox cboxParts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblTypeVar 
      Caption         =   "Type Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmChangeModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================
' Force explicit variable declaration.
'========================================
Option Explicit

Private Sub Form_Load()
    ' Stops the state machine
    StopStateMachine
    
    ' Reload Parts Number from file config\PartNumbers.ini to the combo box
    cboxParts.Clear           ' Deletes previous Parts numbers
    LoadPartNumbers cboxParts ' Load the part numbers
    cboxParts.ListIndex = 0   ' Select the first element by default
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Resumes the state machine to continue with the sequence
    StartStateMachine
End Sub

'================================================================================
' btnChangeModel_Click Button Click Event
'================================================================================
' This button ask for for a change model by using plcChangeOverStarted telegram
' from MES Bosch System
' NOTE: See TelegramHandler module.

' The Sequence that this button makes is:
'   1. Sends to MES System the Type Number and the TypeVar
'   2. Waits 5 seconds until the program receives the response of the MES
'   3. When the socket data is available then Process the telegram using the
'      ReadPLCChangeOver funtion
'   4. Validates if the MES accept the new Model and uses this TypeNumber
'      and the TypeVar
'=================================================
Private Sub btnChangeModel_Click()
    ' Send plcChangeOverStarted Telegram
    SendPLCChangeOver cboxParts.List(cboxParts.ListIndex), txtTypeVar.text
    
    ' Wait until the data arraives
    Dim attempts As Integer
    Do Until machine.SocketAvailable = True
        attempts = attempts = 1
        DoEvents
        Sleep 100
        If attempts = 50 Then
            MsgBox "Error Recibir respuesta de sistema MES ", vbCritical _
            + vbOKOnly, "Error Cambio de Modelo"
            
            AppendLog ("Change Model Failed: MES Socket TIMEOUT")
            Exit Sub
        End If
    Loop
    
    ' Reads the response of the telegram
    If ReadPLCChangeOver = True Then
        machine.TypeNumber = cboxParts.List(cboxParts.ListIndex)
        machine.typevar = txtTypeVar.text
        MsgBox "MES Acepto el Cambio de Modelo", _
            vbOKOnly, "Cambio de Modelo"
        
        AppendLog ("Change Model Success")
        Unload Me
    Else
        MsgBox "MES Rechazo el Cambio de Modelo", vbCritical _
            + vbOKOnly, "Error Cambio de Modelo"
        AppendLog ("Change Model Failed: MES Reject the request")
    End If
End Sub

