VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOpenPort 
      Caption         =   "API Example"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   4695
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOpenPort_Click()
    Dim maker As ZPLMaker
    Set maker = New ZPLMaker
    
    maker.Begin
    maker.SetOrigin 250, 70
    maker.SetFontSize 11, 7
    maker.PutText "CORPORACION TECTRONIC SA de CV"
    
    maker.SetOrigin 350, 105
    maker.SetFontSize 11, 7
    maker.PutText "Prueba 1"
    
    maker.SetOrigin 30, 150
    maker.SetFontSize 11, 7
    maker.PutText "Texto de muestra 1"
    
    maker.SetOrigin 350, 200
    maker.SetFontSize 11, 7
    maker.BarCodeConfig 80, "Y", "Y", "N"
    maker.PutText "corptectr>147896325"
    maker.Terminate
    
    Text1.text = maker.Code
End Sub
