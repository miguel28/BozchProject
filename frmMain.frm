VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "BOSCH Project"
   ClientHeight    =   5115
   ClientLeft      =   330
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   11730
   Begin VB.CommandButton btnAlarms 
      Caption         =   "Alarmas"
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton btnMantenaince 
      Caption         =   "Mantenimiento"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sockMES 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPartNumber 
      Caption         =   "Numero de Parte:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lblOperatorMsg 
      Alignment       =   2  'Center
      Caption         =   "Mensaje para el Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========================
'Global Variables
'==========================


'==========================
'Controls Events
'==========================
Private Sub Form_Initialize()
    InitializeProgram
    StartStateMachine
End Sub

Private Sub Form_Resize()
    ' Diseno responsivo de la forma
    If Me.Width < 10000 Then Me.Width = 10000
    'If Me.height < 10000 Then Me.height = 7000
    
    btnMantenaince.Left = Me.Width - 2500
    btnMantenaince.Top = Me.height - 1500
    
    btnAlarms.Left = Me.Width - 4500
    btnAlarms.Top = Me.height - 1500
    
    lblOperatorMsg.Left = (Me.Width / 2) - (lblOperatorMsg.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    StopStateMachine
    Unload frmPortEmulator
    Unload frmAlarms
    End
End Sub

Private Sub btnMantenaince_Click()
    Dim pass As String
    pass = InputBox("Escriba la contrasena de Mantenimiento")
    If pass = "pass" Then
        StopStateMachine
        Me.Hide
        frmMantenaince.Show
    End If
End Sub

Private Sub btnAlarms_Click()
    AddAlarmMessage "Mensaje de Error"
    frmAlarms.Show
End Sub


'==========================
'Local Defined Functions.
'==========================
Private Sub ConfigureControls()
    'Load Win Socket Configuration of config files
    ConfigureSocket sockMES, "MESSocket.ini"
End Sub

Private Sub OpenPorts()
    'Open Win Socket Configuration of config files
    sockMES.Connect
End Sub
