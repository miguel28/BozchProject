VERSION 5.00
Begin VB.Form frmUtilities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilerias"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   120
      Picture         =   "frmUtilities.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   6435
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton btnTestPrint 
      Caption         =   "TEST PRINTER"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnTestPort 
      Caption         =   "Test Port"
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnTestScanner 
      Caption         =   "Scanner"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnResetCounter 
      Caption         =   "Reset Counter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnResetCounter_Click()
    machine.GoodParts = 0
    machine.BadParts = 0
End Sub

Private Sub btnTestPrint_Click()
    machine.PrintTestZebra
    Sleep 3000
End Sub
