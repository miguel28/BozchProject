VERSION 5.00
Begin VB.Form frmPortEmulator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IO Emulator"
   ClientHeight    =   7905
   ClientLeft      =   14895
   ClientTop       =   570
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   5280
   Begin VB.TextBox txtCioDio 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   5880
      Width           =   4575
   End
   Begin VB.TextBox txtZPL 
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4200
      Width           =   4575
   End
   Begin VB.CommandButton btnSendScanner 
      Caption         =   "Send"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtScanner 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Text            =   "PART123456"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Timer tmrUpdateIO 
      Interval        =   50
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   7
      Left            =   4440
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   6
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   5
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   3
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnInputs 
      Height          =   375
      Index           =   0
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Cognex Cio Dio"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Zebra ZPL"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Scanner"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "  0           1           2            3           4            5           6           7"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Inputs"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "  0           1           2            3           4            5           6           7"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Outputs"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   7
      Left            =   4440
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   3840
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   3240
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   2640
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   1440
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   840
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpOutput 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmPortEmulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnInputs_Click(Index As Integer)
    Dim activated As Boolean
    activated = IOPortCom.GetInput(Index)
    
    If activated = True Then
        IOPortCom.SetInput Index, False
    Else
        IOPortCom.SetInput Index, True
    End If
End Sub

Private Sub btnSendScanner_Click()
    ScannerAvailable = True
End Sub

Private Sub tmrUpdateIO_Timer()
    Dim i As Integer
    
    For i = 0 To 7
        'Check Outputs
        If IOPortCom.GetOutput(i) Then
            shpOutput(i).BackColor = &HC000&
        Else
            shpOutput(i).BackColor = &H80000000
        End If
        
        'Check Inputs
        If IOPortCom.GetInput(i) Then
            btnInputs(i).BackColor = &HC000&
        Else
            btnInputs(i).BackColor = &H8000000F
        End If
        
    Next i
End Sub
