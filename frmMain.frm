VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "BOSCH Project"
   ClientHeight    =   7935
   ClientLeft      =   330
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11925
   Begin VB.Timer tmrUpdateStateMachine 
      Interval        =   150
      Left            =   600
      Top             =   120
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mensajes del Sistema"
      Height          =   1815
      Left            =   6120
      TabIndex        =   19
      Top             =   5280
      Width           =   5655
      Begin VB.Label lblAMS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ESTACION SIN CONEXION                     MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8280
      TabIndex        =   15
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   5775
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4080
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnUtils 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Utilerias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         Picture         =   "frmMain.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambio Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         Picture         =   "frmMain.frx":13F3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtOperador 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5775
      Begin VB.TextBox txtTypeVar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboxParts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtcounter 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtSN 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Numero de Seria de la parte, este es escaneado por el escaner manual"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtbadunits 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblTypeVar 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero de Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades Buenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Serial "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades Malas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1935
      End
   End
   Begin VB.PictureBox picBosch 
      Height          =   1455
      Left            =   2880
      Picture         =   "frmMain.frx":1CB6
      ScaleHeight     =   1395
      ScaleWidth      =   5715
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton btnMantenaince 
      Caption         =   "Mantenimiento"
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   7320
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sockMES 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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


Private Sub Command1_Click()
    Dim xml As XMLParser
    Set xml = New XMLParser
    xml.Load ("xmls\partReceived_request.xml")
    MsgBox xml.Code
    xml.SetAttribute "identifier", "50505"
    MsgBox xml.Code

End Sub

'==========================
'Controls Events
'==========================
Private Sub Form_Initialize()
    InitializeProgram
    LoadPartNumbers cboxParts
    StartStateMachine
End Sub

Private Sub Form_Resize()
    ' Diseno responsivo de la forma
    If Me.Width < 10000 Then Me.Width = 10000
    'If Me.height < 10000 Then Me.height = 7000
    
    btnMantenaince.Left = Me.Width - 2500
    btnMantenaince.Top = Me.height - 1500
    picBosch.Left = (Me.Width / 2) - (picBosch.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    StopStateMachine
    Unload frmPortEmulator
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

Private Sub tmrUpdateStateMachine_Timer()
    UpdateStateMachine
End Sub
