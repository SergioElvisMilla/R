VERSION 5.00
Begin VB.Form FrmProImpCo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFinPag 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3060
      TabIndex        =   10
      Top             =   1740
      Width           =   915
   End
   Begin VB.TextBox txtIniPag 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   1740
      Width           =   915
   End
   Begin VB.CheckBox chkCom 
      Alignment       =   1  'Right Justify
      Caption         =   "Imprimir Comentario"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkNota 
      Alignment       =   1  'Right Justify
      Caption         =   "Imprimir Nota"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.PictureBox PicImagen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   3945
      TabIndex        =   13
      Top             =   3510
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2220
      TabIndex        =   11
      Top             =   2145
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3330
      TabIndex        =   12
      Top             =   2145
      Width           =   1065
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmProImpCo.frx":0000
      Left            =   1680
      List            =   "frmProImpCo.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   810
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2340
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   3
      X1              =   150
      X2              =   4380
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   150
      X2              =   4380
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Documento :"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblNumdoc 
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   205
      Left            =   2580
      TabIndex        =   2
      Top             =   160
      Width           =   1755
   End
   Begin VB.Label lblCodDoc 
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   205
      Left            =   1710
      TabIndex        =   1
      Top             =   160
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de Impresión :"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "FrmProImpCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bOk As Boolean
Private Sub cmbTipo_Click()
    Dim iTipo As Integer
    If cmbTipo.ListIndex >= 0 Then iTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
    Select Case iTipo
        Case 1, 2, 5
            chkNota.Enabled = True
            chkCom.Enabled = True
            txtIniPag.Enabled = False
            txtFinPag.Enabled = False
        Case 3, 4
            chkNota.Enabled = False
            chkCom.Enabled = False
            txtIniPag.Enabled = True
            txtFinPag.Enabled = True
    End Select
End Sub
Private Sub cmdaceptar_Click()
    bOk = True
    Me.Hide
    DoEvents
End Sub
Private Sub cmdCancelar_Click()
    bOk = False
    Me.Hide
    DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{tab}"
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub
Function ShowForm(ByVal sCodDoc As String, ByVal sNumDoc As String) As Boolean
    lblCodDoc.Caption = Trim(sCodDoc)
    lblNumDoc.Caption = Trim(sNumDoc)
    cmbTipo.ListIndex = 0: Call cmbTipo_Click
    Me.Show 1
    ShowForm = bOk
End Function
Property Get TipImp() As Byte
    TipImp = cmbTipo.ItemData(Me.cmbTipo.ListIndex)
End Property
Property Get ImpNota() As Byte
    ImpNota = CByte(chkNota.Value)
End Property
Property Get ImpCom() As Byte
    ImpCom = CByte(chkCom.Value)
End Property
Property Get Paginas() As String
    Dim iCnt As Integer, sPag As String
    For iCnt = Val(txtIniPag.Text) To Val(txtFinPag.Text)
        If iCnt > 0 Then sPag = sPag & ";" & iCnt
    Next iCnt
    If sPag <> "" Then sPag = sPag & ";"
    Paginas = sPag
End Property
Private Sub txtFinPag_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txtIniPag_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
