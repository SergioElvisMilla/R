VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmImpInf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de Impresión"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1700
      TabIndex        =   3
      Top             =   960
      Width           =   1065
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   40429
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   40429
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmImpInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bOk As Boolean

Private Sub cmdAceptar_Click()
    bOk = True
    Me.Hide
End Sub
Private Sub cmdCancelar_Click()
    bOk = False
    Me.Hide
End Sub

Public Sub ShowForm()
    Call Me.Show(1)
End Sub

Property Get FecIni() As String
    FecIni = dtpFecha(0).Value
End Property

Property Get FecFin() As String
    FecFin = dtpFecha(1).Value
End Property

Property Get Ok() As Boolean
    Ok = bOk
End Property
