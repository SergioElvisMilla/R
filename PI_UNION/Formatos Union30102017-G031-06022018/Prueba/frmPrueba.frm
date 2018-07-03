VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPrueba 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   Icon            =   "frmPrueba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Empresa"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
      Begin VB.TextBox txtDato 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Text            =   "E1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Documento"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
      Begin VB.TextBox txtDato 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Text            =   "100710"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDato 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Text            =   "OPPY"
         Top             =   320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Numero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   765
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdBoton 
      Caption         =   "Invocar Formato"
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   540
      Left            =   120
      Top             =   3000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   953
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Password=nexus2002;Persist Security Info=True;User ID=sa;Initial Catalog=nexus;Data Source=venus"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=nexus2002;Persist Security Info=True;User ID=sa;Initial Catalog=nexus;Data Source=venus"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCnx As String

Const Txt_CodDoc As Byte = 0
Const Txt_NumDoc As Byte = 1
Const Txt_CodEmp As Byte = 2

Private Sub cmdBoton_Click(Index As Integer)
    Dim oCnx As ADODB.Connection, oImp As Object, cnn As String
    Select Case Index
        Case 0
            'Set oImp = CreateObject("PIElectrodata.cProImpOc")
            'Set oImp = CreateObject("PIElectrodata.cProImpNI")
            'Set oImp = CreateObject("PIElectrodata.cProImpGr03")
            'Set oImp = CreateObject("PIElectrodata.cProImpND")
            Set oImp = CreateObject("PiIntuitive.cProImpInf")
            Set oCnx = New ADODB.Connection
            oCnx.ConnectionString = sCnx
            oCnx.CursorLocation = adUseClient
            oCnx.Open
            Call oCnx.Execute("Set DateFormat Dmy")
            oImp.CodEmp = Trim$("" & txtDato(Txt_CodEmp).Text)
            On Error Resume Next
            oImp.CodDoc = Trim$("" & txtDato(Txt_CodDoc).Text)
            oImp.NumDoc = Trim$("" & txtDato(Txt_NumDoc).Text)
            oImp.MovCab = GetArcTab(txtDato(Txt_CodEmp).Text, txtDato(Txt_CodDoc).Text, 0, oCnx)
            oImp.MovDet = GetArcTab(txtDato(Txt_CodEmp).Text, txtDato(Txt_CodDoc).Text, 1, oCnx)
            Set oImp.Connection = oCnx
            oImp.PrintShow
            Set oImp = Nothing
            oCnx.Close
            Set oCnx = Nothing
    End Select
End Sub

Private Sub Form_Load()
     sCnx = "provider=MSDataShape.1;uid=sidigeuser;pwd=G2d5P0s2A1c980089c1A2s0P5d2G;driver={SQL server};server=SERVER2005;database=INTUITIVE;dsn=;,,connection=adconnectasync"
     'sCnx = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=Electrodata;Data Source=ADOC-06"
End Sub

Private Function GetArcTab(ByVal CodEmp As String, ByVal CodDoc As String, ByVal Tipo As Byte, ByVal oCnx As ADODB.Connection) As String
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = oCnx.Execute("Select CodDoc,ArcCab,ArcDet " & vbCrLf & _
                             "From CfgDoc00 " & vbCrLf & _
                             "Where CodEmp  = '" & Trim$("" & CodEmp) & "' And CodDoc = '" & Trim$("" & CodDoc) & "'")
    If Not (rsTmp.EOF And rsTmp.BOF) Then
        Select Case Tipo
            Case 0
                GetArcTab = Trim$("" & rsTmp.Fields("ArcCab").Value)
            Case 1
                GetArcTab = Trim$("" & rsTmp.Fields("ArcDet").Value)
        End Select
    End If
    Call Close_RecordSet(rsTmp)
End Function

