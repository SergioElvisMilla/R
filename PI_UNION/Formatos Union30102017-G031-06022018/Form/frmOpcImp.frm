VERSION 5.00
Begin VB.Form frmOpcImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox chkFnl 
      Caption         =   "Imprimir Final"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CheckBox chkAde 
      Caption         =   "Imprimir Adelanto"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
End
Attribute VB_Name = "frmOpcImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bOk As Boolean
Public cc As Boolean
Property Get ImpAde()
    ImpAde = chkAde.Value = 1
End Property

Property Get ImpFnl()
    ImpFnl = chkFnl.Value = 1
End Property
Public Function ShowForm() As Boolean
    cc = False
    Screen.MousePointer = 0
    Call Me.Show(1)
    ShowForm = bOk
    Screen.MousePointer = 11
End Function
Private Sub chkAde_Click()
    If chkAde.Value = 1 Then
    chkFnl.Value = 0
    End If
End Sub
Private Sub chkFnl_Click()
    If chkFnl.Value = 1 Then
    chkAde.Value = 0
    End If
End Sub

Private Sub cmdaceptar_Click()
    bOk = True
    Me.Hide
    cc = True
End Sub

