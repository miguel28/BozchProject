VERSION 5.00
Begin VB.Form frmMantenaince 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnUnClampBoard 
      Caption         =   "Desanclar Tablilla"
      Height          =   735
      Left            =   5520
      TabIndex        =   21
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton btnClampBoard 
      Caption         =   "Anclar Tabilla"
      Height          =   735
      Left            =   5520
      TabIndex        =   20
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sensores - Actuadores"
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   6855
      Begin VB.Label Label15 
         Caption         =   "Luz 2"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Luz 1"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   18
         Left            =   5040
         Top             =   2640
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   17
         Left            =   5040
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   16
         Left            =   5040
         Top             =   1680
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   15
         Left            =   5040
         Top             =   1200
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   14
         Left            =   5040
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   13
         Left            =   5040
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Extender Piston 2"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Retaer Piston 2"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Extender Piston 1"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Retaer Piston 1"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   12
         Left            =   2280
         Top             =   2160
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   11
         Left            =   2280
         Top             =   1680
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   10
         Left            =   2280
         Top             =   1200
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   9
         Left            =   2280
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   8
         Left            =   2280
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Sensor Piston 2 Extendido"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Sensor Piston 2 Retraido"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Sensor Piston 1 Extendido"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Sensor Piston 1 Retraido"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Sensor de Tablilla"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entradas - Salidas"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.Timer tmrUpdateIO 
         Interval        =   50
         Left            =   4560
         Top             =   240
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   3
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   4
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   5
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   6
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton btnOutputs 
         Height          =   375
         Index           =   7
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "  0           1           2            3           4            5           6           7"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Entradas"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "  0           1           2            3           4            5           6           7"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Salidas"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   0
         Left            =   360
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   1
         Left            =   960
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   2
         Left            =   1560
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   3
         Left            =   2160
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   4
         Left            =   2760
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   5
         Left            =   3360
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   6
         Left            =   3960
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpInputs 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000007&
         Height          =   375
         Index           =   7
         Left            =   4560
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.CommandButton btnPrintTest 
      Caption         =   "Etiqueta De Prueba"
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMantenaince"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClampBoard_Click()
    IOPortCom.SetOutput 0, True
    IOPortCom.SetOutput 1, False
    IOPortCom.SetOutput 2, True
    IOPortCom.SetOutput 3, False
End Sub

Private Sub btnOutputs_Click(Index As Integer)
    Dim activated As Boolean
    activated = IOPortCom.GetOutput(Index)
    
    If activated = True Then
        IOPortCom.SetOutput Index, False
    Else
        IOPortCom.SetOutput Index, True
    End If
End Sub

Private Sub btnUnClampBoard_Click()
    IOPortCom.SetOutput 0, False
    IOPortCom.SetOutput 1, True
    IOPortCom.SetOutput 2, False
    IOPortCom.SetOutput 3, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

Private Sub tmrUpdateIO_Timer()
    Dim i As Integer
    
    For i = 0 To 7
        'Check Outputs
        If IOPortCom.GetOutput(i) Then
            btnOutputs(i).BackColor = &HC000&
        Else
            btnOutputs(i).BackColor = &H80000000
        End If
        
        'Check Inputs
        If IOPortCom.GetInput(i) Then
            shpInputs(i).BackColor = &HC000&
        Else
            shpInputs(i).BackColor = &H8000000F
        End If
        
    Next i
    
    'Verificar Sensor de Tablilla
    If IOPortCom.GetInput(0) Then
        shpInputs(8).BackColor = &HC000&
    Else
        shpInputs(8).BackColor = &H8000000F
    End If
    
    'Verificar Sensor Piston 1 Retraido
    If IOPortCom.GetInput(1) Then
        shpInputs(9).BackColor = &HC000&
    Else
        shpInputs(9).BackColor = &H8000000F
    End If
    
    'Verificar Sensor Piston 1 Extendido
    If IOPortCom.GetInput(2) Then
        shpInputs(10).BackColor = &HC000&
    Else
        shpInputs(10).BackColor = &H8000000F
    End If
    
    'Verificar Sensor Piston 2 Retraido
    If IOPortCom.GetInput(3) Then
        shpInputs(11).BackColor = &HC000&
    Else
        shpInputs(11).BackColor = &H8000000F
    End If
    
    'Verificar Sensor Piston 2 Extendido
    If IOPortCom.GetInput(4) Then
        shpInputs(12).BackColor = &HC000&
    Else
        shpInputs(12).BackColor = &H8000000F
    End If
    
    
    'Verifica Piston 1 Retraer
    If IOPortCom.GetOutput(0) Then
        shpInputs(13).BackColor = &HC000&
    Else
        shpInputs(13).BackColor = &H8000000F
    End If
    
    'Verifica Piston 1 Extender
    If IOPortCom.GetOutput(1) Then
        shpInputs(14).BackColor = &HC000&
    Else
        shpInputs(14).BackColor = &H8000000F
    End If
    
    'Verifica Piston 2 Retraer
    If IOPortCom.GetOutput(2) Then
        shpInputs(15).BackColor = &HC000&
    Else
        shpInputs(15).BackColor = &H8000000F
    End If
    
    'Verifica Piston 2 Extender
    If IOPortCom.GetOutput(3) Then
        shpInputs(16).BackColor = &HC000&
    Else
        shpInputs(16).BackColor = &H8000000F
    End If
    
    'Verifica Luz 1
    If IOPortCom.GetOutput(4) Then
        shpInputs(17).BackColor = &HC000&
    Else
        shpInputs(17).BackColor = &H8000000F
    End If
    
    'Verifica Luz 2
    If IOPortCom.GetOutput(5) Then
        shpInputs(18).BackColor = &HC000&
    Else
        shpInputs(18).BackColor = &H8000000F
    End If
    
End Sub
