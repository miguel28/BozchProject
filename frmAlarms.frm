VERSION 5.00
Begin VB.Form frmAlarms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alarmas"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClean 
      Caption         =   "Limpiar Alarmas"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtAlarms 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAlarms.frx":0000
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmAlarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
