VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProImpLT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ICodDoc As String
Private INumDoc As String
Private IProImp As String
Private ICodEmp As String
Private ICodGru As String
Private ITitFrm As String
Private IMovCab As String
Private IMovDet As String
Private IDevice As String
Private IParam1 As String
Private RwCab As ADODB.Recordset
Private ICnx As ADODB.Connection
Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set ICnx = Cnn
End Property
Public Property Let NumDoc(ByVal Cad As String)
    INumDoc = Cad
End Property
Public Property Let CodDoc(ByVal Cad As String)
    ICodDoc = Cad
End Property
Public Property Let ProImp(ByVal Cad As String)
    IProImp = Cad
End Property
Public Property Let CodEmp(ByVal Cad As String)
    ICodEmp = Cad
End Property
Public Property Let CodGru(ByVal Cad As String)
    ICodGru = Cad
End Property
Public Property Let MovCab(ByVal Cad As String)
    IMovCab = Cad
End Property
Public Property Let MovDet(ByVal Cad As String)
    IMovDet = Cad
End Property
Public Property Let RutImp(ByVal Cad As String)
    IDevice = Cad
End Property
Public Property Let Param1(ByVal Cad As String)

End Property
Private Sub Class_Initialize()
    Set RwCab = New ADODB.Recordset
End Sub
Private Sub Class_Terminate()
    Set RwCab = Nothing
End Sub
Public Function PrintShow() As Boolean
    PrintShow = PrintDoc(True)
End Function
Public Function PrintDoc(Optional PreView As Boolean = False) As Boolean
    Screen.MousePointer = 11
    On Error GoTo Solucion
    Set RwCab = ICnx.Execute("SELECT CAB.NUMDOC,CAB.CODANE,'LIMA' AS CODLOC,DAY(CAB.FECDOC) AS DIAD,'/' AS SEP1,MONTH(CAB.FECDOC) AS MESD ,'/' " & _
                             " AS SEP2,YEAR(CAB.FECDOC) AS AÑOD,CAB.FECven,CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS XTIPMON,CAB.TOTDOC,CAB.TOTDOCL,CAB.NOMANE,RTRIM(ANE.XTIPIDE2) +'-'+ RTRIM(ANE.IDEANE2) AS IDEANE2,CAB.DIRANE,UBIZON.DESCRIPCION,(CAB.IDEANE) AS IDEANE,CAB.TELANE FROM " & IMovCab & " AS CAB  LEFT JOIN ANEXO AS ANE ON CAB.CODANE=ANE.CODANE LEFT JOIN UBIGEO AS UBIZON ON ANE.XUBIGEO = UBIZON.UBIGEO WHERE CodDoc = '" & Trim(ICodDoc) & "' And NUMDOC = '" & INumDoc & "'")
    If RwCab.EOF = False Then
        Dim rpt As rptLetra, rsEmp As ADODB.Recordset
        Set rpt = New rptLetra
        With rpt.Sections("CabPag")
            .Controls("lblNumDoc").Caption = RwCab.Fields("NumDoc").Value
            .Controls("lblCodAne").Caption = RwCab.Fields("CodAne").Value
            .Controls("lblCodLoc").Caption = RwCab.Fields("CodLoc").Value
            .Controls("lblAñoD").Caption = RwCab.Fields("AñoD").Value
            .Controls("lblSep1").Caption = RwCab.Fields("Sep1").Value
            .Controls("lblMesd").Caption = Format$(RwCab.Fields("MesD").Value, "00")
            .Controls("lblSep2").Caption = RwCab.Fields("Sep2").Value
            .Controls("lblDiaD").Caption = RwCab.Fields("DiaD").Value
            .Controls("lblFCA").Caption = Format(RwCab.Fields("Fecven").Value, "YYYY")
            .Controls("lblFCM").Caption = Format$(Format(RwCab.Fields("Fecven").Value, "MM"), "00")
            .Controls("lblFCD").Caption = Format(RwCab.Fields("Fecven").Value, "DD")
            .Controls("lblXTipMon").Caption = RwCab.Fields("XTipMon").Value
            .Controls("lblTotDoc").Caption = Format$(" " & RwCab.Fields("TotDoc").Value, "#,0.00")
            .Controls("lblTotDocL").Caption = RwCab.Fields("TotDocL").Value
            .Controls("lblNomAne").Caption = Trim$(" " & RwCab.Fields("NomAne").Value)
            .Controls("lblDirAne").Caption = Trim$(" " & RwCab.Fields("DirAne").Value)
            .Controls("lblIdeAne").Caption = Trim$(" " & RwCab.Fields("IdeAne").Value)
            .Controls("lblTelAne").Caption = Mid$(Trim$(" " & RwCab.Fields("TelAne").Value), 1, 18)
            Set rsEmp = New ADODB.Recordset
            Set rsEmp = ICnx.Execute("Select NomEmp, RutLog, DirEmp, FaxEmp From Empresas Where CodEmp='" & ICodEmp & "'", , adCmdText)
            If rsEmp.RecordCount > 0 Then
                .Controls("lblEmpresa").Caption = rsEmp.Fields("NomEmp").Value
                .Controls("lblSubTitulo1").Caption = rsEmp.Fields("NomEmp").Value
                .Controls("lblSubTitulo2").Caption = rsEmp.Fields("DirEmp").Value & IIf((Trim$(" " & rsEmp.Fields("FaxEmp").Value) <> ""), " - Telefax: " & Trim$(" " & rsEmp.Fields("FaxEmp").Value), "")
                If Dir(Trim$(" " & rsEmp.Fields("RutLog").Value)) <> "" Then Set .Controls("Logo").Picture = LoadPicture(Trim$(" " & rsEmp.Fields("RutLog").Value))
            End If
            rsEmp.Close
            Set rsEmp = Nothing
        End With
        Set rpt.DataSource = RwCab
        Screen.MousePointer = 0
        
        If PreView Then
            rpt.Show 1
        Else
            rpt.PrintReport
        End If
        
        Set rpt = Nothing
    Else
        Screen.MousePointer = 0
        MsgBox "No se pudo encontrar el documento", vbExclamation
    End If
    RwCab.Close
    Set RwCab = Nothing
    Exit Function
Solucion:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "¿Desea ver mas información?", vbExclamation + vbYesNo, "Error") = vbYes Then
            MsgBox Err.Description, vbCritical, "Error"
            Err.Clear
        End If
    End If
End Function
