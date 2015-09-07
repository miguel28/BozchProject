VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "BOSCH Project"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAlarms 
      Caption         =   "Alarmas"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton btnMantenaince 
      Caption         =   "Mantenimiento"
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton btnOpenExample 
      Caption         =   "Zebra Example"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Timer tmrIOSync 
      Left            =   3120
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock sockMES 
      Left            =   2640
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm comHandScanner 
      Left            =   2040
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm comCognex 
      Left            =   1440
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm comScanner 
      Left            =   840
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm comZebra 
      Left            =   240
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
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
    ConfigureControls
    'OpenPorts
End Sub
Private Sub Form_Load()
    
End Sub

Private Sub btnMantenaince_Click()
    Me.Hide
    frmMantenaince.Show
End Sub

Private Sub btnOpenExample_Click()
    frmExample.Show
    Me.Hide
End Sub

Private Sub btnAlarms_Click()
    frmAlarms.Show
End Sub

'==========================
'Port Receiver Events
'==========================


'==========================
'Local Defined Functions.
'==========================
Private Sub ConfigureControls()
    'Load Serial COM Configuration of config files
    ConfigurePort comZebra, "ZebraPort.ini"
    ConfigurePort comScanner, "ScannerPort.ini"
    ConfigurePort comCognex, "CognexPort.ini"
    ConfigurePort comHandScanner, "HandScannerPort.ini"
    
    'Load Win Socket Configuration of config files
    ConfigureSocket sockMES, "MESSocket.ini"
    
    'Config IO Port
    Set IOPortCom = New IOPort
    UseEmulator = True
    
    If UseEmulator = True Then frmPortEmulator.Show
    
End Sub

Private Sub OpenPorts()
    'Open Serial COM
    comZebra.PortOpen
    comScanner.PortOpen
    comCognex.PortOpen
    comHandScanner.PortOpen

    'Open Win Socket Configuration of config files
    sockMES.Connect
    
End Sub
