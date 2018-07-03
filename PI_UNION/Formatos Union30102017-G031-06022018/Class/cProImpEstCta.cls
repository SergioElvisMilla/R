VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpEstCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private INumDoc As String
Private IProImp As String
Private ICodEmp As String
Private ICodGru As String
Private ITitFrm As String
Private IMovCab As String
Private sMovEdi As String
Private IMovDet As String
Private sDevice As String
Private sParam1 As String
Private INivel As Byte
Private ICnx As ADODB.Connection
Dim rsCab As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Dim rsEmp As ADODB.Recordset
Dim RsDetLar As ADODB.Recordset
Dim RsDocRef As ADODB.Recordset
Dim bFlagAdelanto As Boolean
Private rsDocRef2 As ADODB.Recordset
Private rsDocRef3 As ADODB.Recordset
Private RsSer As ADODB.Recordset
Private RsAde As ADODB.Recordset
Private bConAde As Boolean
Private bImpConDetalle As Boolean
Private bImpConCondiciones As Boolean
Private bImpSinFiltro As Boolean
Private sComentario As String
Private ICodDoc As String

Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set ICnx = Cnn
End Property
Public Property Let CodDoc(ByVal Cad As String)
    ICodDoc = Cad
End Property
Public Property Let NumDoc(ByVal Cad As String)
    INumDoc = Cad
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
Public Property Let MOVEDI(ByVal Valor As String)
    sMovEdi = Valor
End Property
Public Property Let RutImp(ByVal Cad As String)
    sDevice = Cad
End Property
Public Property Let Param1(ByVal Cad As String)
    sParam1 = Cad
End Property
Public Property Let Nivel(ByVal Cad As Byte)
    INivel = Cad
End Property
Public Function PrintShow() As Boolean
    PrintShow = PrintDoc(True)
End Function
Public Function PrintDoc(Optional PreView As Boolean = False) As Boolean
    Dim sCurDevice As String, bPrint As Boolean
    Screen.MousePointer = 11: bPrint = True
    
    Set rsEmp = ICnx.Execute("SELECT TOP 1 NOMEMP,DIREMP,TELEMP,FAXEMP,RUCEMP,RUtLOG FROM EMPRESAS WHERE CODEMP='" & ICodEmp & "'", , adCmdText)
    
    Set rsCab = ICnx.Execute("SELECT TOP 1 CAB.TIPCAM, CAB.NUMLOC,LOC.DIRLOC,UBL.DESCRIPCION AS UBILOC,CFG.SERRIE AS SERIE,CAB.XTIPMOV,UBIGEO.DESCRIPCION AS XUBIGEO,CAB.CODDOC,CFG.DESCRI AS DESDOC,CAB.NUMDOC, CAB.NUMREA,CAB.FECDOC,CAB.FECVEN,CAB.CODANE,CAB.NOMANE, " & vbCrLf & _
                             "CAB.DIRANE, CAB.TELANE, CAB.IDEANE, CAB.CODRES, CAB.CODVEN,ANE.NOMANE AS NOMVEN,ANE1.UBIZON AS CODUBIZON, CAB.TIPTR1, CAB.NUMRF1, CAB.TIPTR2,CAB.NUMRF2,UBI.NOMBRE AS UBIZON, " & vbCrLf & _
                             "CAB.SUBTOT - (CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) AS IMPNET, CAB.TOTIMP1+CAB.TOTIMP2+CAB.TOTIMP3+CAB.TOTIMP4 AS TOTIMP, CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4 AS TOTDSC, " & vbCrLf & _
                             "CAB.SUBTOT,CAB.TOTIMP1,CAB.TOTIMP2,CAB.TOTIMP3,CAB.TOTIMP4,CAB.TOTDSC1,CAB.TOTDSC2,CAB.TOTDSC3,CAB.TOTDSC4,CAB.TOTDOC, MAE.DESITE AS CONPAG,MAE1.DESITE AS FORPAG,CAB.NUMDIA,CAB.CODMEM,CAB.COMMEM, " & vbCrLf & _
                             "CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS SIMMON, CAB.TOTDOCL, Case When IsNull(Tmp02.CodDoc,'') = '' Then '' Else lTrim(rTrim(IsNull(CfgP.SerRie,''))) + '-' + lTrim(rTrim(IsNull(Tmp02.NumDoc,''))) End As DocPad  " & vbCrLf & _
                             "FROM " & IMovCab & " AS CAB " & vbCrLf & _
                             "LEFT JOIN MAETABDET AS MAE ON MAE.CODTAB='CONPAG' AND XCONPAG=MAE.CODITE " & vbCrLf & _
                             "LEFT JOIN MAETABDET AS MAE1 ON MAE1.CODTAB='FORPAG' AND XFORPAG=MAE1.CODITE " & vbCrLf & _
                             "LEFT JOIN ANEXO AS ANE ON CAB.CODVEN=ANE.CODANE " & vbCrLf & _
                             "LEFT JOIN CFGDOC00 AS CFG ON CAB.CODDOC=CFG.CODDOC " & vbCrLf & _
                             "LEFT JOIN ANEXO AS ANE1 ON CAB.CODANE = ANE1.CODANE " & vbCrLf & _
                             "LEFT JOIN UBIZON AS UBI ON ANE1.UBIZON = UBI.UBIZON " & vbCrLf & _
                             "LEFT JOIN UBIGEO ON ANE1.XUBIGEO = UBIGEO.UBIGEO " & vbCrLf & _
                             "LEFT JOIN LOCALES LOC ON CAB.CODANE = LOC.CODANE AND CAB.NUMLOC = LOC.NUMLOC " & vbCrLf & _
                             "LEFT JOIN UBIGEO UBL ON LOC.XUBIGEO = UBL.UBIGEO " & vbCrLf & _
                             "Left Join (Select CodDocPad,NumDocPad,CodDoc,NumDoc From MovCtaArt Group By CodDocPad,NumDocPad,CodDoc,NumDoc) Tmp On Tmp.CodDoc = Cab.CodDoc And Tmp.NumDoc = Cab.NumDoc " & vbCrLf & _
                             "Left Join (Select CodDocPad,NumDocPad,CodDoc,NumDoc From MovCtaArt Where  CodDoc In ('" & Replace(sParam1, ";", "','") & "') Group By CodDocPad,NumDocPad,CodDoc,NumDoc) Tmp02 On Tmp02.CodDocPad = Tmp.CodDocPad And Tmp02.NumDocPad = Tmp.NumDocPad" & vbCrLf & _
                             "Left Join CfgDoc00 CfgP On CfgP.CodDoc = Tmp02.CodDoc And CfgP.CodEmp = '" & Trim(ICodEmp) & "' " & vbCrLf & _
                             "WHERE Cab.CodDoc = '" & Trim(ICodDoc) & "' And Cab.NUMDOC='" & INumDoc & "'", , adCmdText)
     
    Set RsDet = ICnx.Execute("SELECT Md.CodSubAlm, Md.NumIte , MD.CodArt, A.CODARTANT1, Md.DesArt, Md.DesArt1, Md.xTipUni, Mae.DesIte As DesUni, " & vbCrLf & _
                             "Md.Dsc001, Md.Dsc002, Md.Dsc003, Md.Dsc004,  Md.PreVta, Md.Imp001," & vbCrLf & _
                             "Case When Md.CanTot != 0 Then Md.CanTot * Md.PreVta/Md.CanTot Else Md.PreVta End AS PreVta1, Md.CanTot , Md.PreVta*Md.CanTot As PreIte, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*(1-((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100)))) As TotDsc, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100))) As PreTot " & vbCrLf & _
                             "FROM " & IMovDet & " MD " & vbCrLf & _
                             "LEFT JOIN ARTICULO A ON A.CODEMP = '" & ICodEmp & "' AND MD.CODSUBALM = A.CODSUBALM AND MD.CODART=A.CODART " & vbCrLf & _
                             "LEFT JOIN ARTDET AD ON AD.CODEMP = A.CODEMP AND A.CODSUBALM = AD.CODSUBALM AND A.XTIPALM = AD.XTIPALM AND A.CODART = AD.CODART " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X17 ON AD.XTIPALM = X17.XTIPALM AND AD.XSUBART17 = X17.CODSUB AND X17.XSUBART = 'XSUBART17' " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X11 ON AD.XTIPALM = X11.XTIPALM AND AD.XSUBART11 = X11.CODSUB AND X11.XSUBART = 'XSUBART11' " & vbCrLf & _
                             "LEFT JOIN MAETABDET MAE ON MAE.CODTAB = 'XTIPUNI' AND MD.XTIPUNI = MAE.CODITE " & vbCrLf & _
                             "WHERE Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NUMDOC='" & INumDoc & "' AND ISNULL(MD.NUMITE1,'')='' AND MD.CODSUBALM <> 'OT'" & vbCrLf & _
                             "ORDER BY Md.NUMITE  Asc", , adCmdText)

        
    Set RsDetLar = ICnx.Execute("Select Mdl01.CodDoc, Mdl01.NumDoc, Mdl01.NomDetLar NomDetLar01, Mdl01.DetLar DetLar01 " & vbCrLf & _
                                ", IsNull(Mdl02.NomDetLar,'') NomDetLar02, IsNull(Mdl02.DetLar,'') DetLar02 " & vbCrLf & _
                                ", IsNull(Mdl03.NomDetLar,'') NomDetLar03, IsNull(Mdl03.DetLar,'') DetLar03 " & vbCrLf & _
                                ", IsNull(Mdl04.NomDetLar,'') NomDetLar04, IsNull(Mdl04.DetLar,'') DetLar04 " & vbCrLf & _
                                "From MovDetLar Mdl01 " & vbCrLf & _
                                "Left Join (Select CodDoc, NumDoc, NomDetLar, DetLar From MovDetLar Where CodEmp='" & Trim$("" & ICodEmp) & "' And CodDoc='" & Trim$("" & rsCab!CodDoc) & "' And NumDoc='" & INumDoc & "' And NumDetLar = 2) Mdl02 On Mdl02.CodDoc = Mdl01.CodDoc And Mdl02.NumDoc = Mdl01.NumDoc " & vbCrLf & _
                                "Left Join (Select CodDoc, NumDoc, NomDetLar, DetLar From MovDetLar Where CodEmp='" & Trim$("" & ICodEmp) & "' And CodDoc='" & Trim$("" & rsCab!CodDoc) & "' And NumDoc='" & INumDoc & "' And NumDetLar = 3) Mdl03 On Mdl03.CodDoc = Mdl01.CodDoc And Mdl03.NumDoc = Mdl01.NumDoc " & vbCrLf & _
                                "Left Join (Select CodDoc, NumDoc, NomDetLar, DetLar From MovDetLar Where CodEmp='" & Trim$("" & ICodEmp) & "' And CodDoc='" & Trim$("" & rsCab!CodDoc) & "' And NumDoc='" & INumDoc & "' And NumDetLar = 4) Mdl04 On Mdl04.CodDoc = Mdl01.CodDoc And Mdl04.NumDoc = Mdl01.NumDoc " & vbCrLf & _
                                "Where Mdl01.CodEmp='" & Trim$("" & ICodEmp) & "' And Mdl01.CodDoc='" & Trim$("" & rsCab!CodDoc) & "' And Mdl01.NumDoc='" & INumDoc & "' And Mdl01.NumDetLar = 1")
                
    Set RsDocRef = ICnx.Execute("Select Mca02.CodDocPad CodDocOp, Mca02.NumDocPad NumDocOp" & vbCrLf & _
                                "From MovCtaArt Mca01" & vbCrLf & _
                                "Left Join MovCtaArt Mca02 On Mca02.CodDoc = Mca01.CodDocPad And Mca02.NumDoc = Mca01.NumDocPad" & vbCrLf & _
                                "Where Mca01.CodDoc = '" & Trim(ICodDoc) & "' And Mca01.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf & _
                                "Group By Mca02.CodDocPad, Mca02.NumDocPad", , adCmdText)
                                
                
    If Not rsCab.EOF Then
            
        Dim pvWindow As Object, prt As Printer
        Set pvWindow = CreateObject("PrtTools.cPreView")
        Set prt = GetPrinter(sDevice)
        
        If Not prt Is Nothing Then Set Printer = prt
        Set prt = Nothing
        Set pvWindow.ActivePrinter = Printer
        
        If PreView Then
            Call PrintData(pvWindow.ActivePage)
            Screen.MousePointer = 0
            bPrint = pvWindow.ShowPreView()
        End If
        If bPrint Then
            Call PrintData(pvWindow.ActivePrinter)
            pvWindow.ActivePrinter.EndDoc
            Screen.MousePointer = 0
        End If
        Set pvWindow = Nothing
        PrintDoc = True
    Else
        Screen.MousePointer = 0
        Call MsgBox("No se pudo encontrar el documento " & Trim(ICodDoc) & "-" & Trim(INumDoc) & ".", vbExclamation, "Mensaje al Usuario")
    End If
    Call Close_RecordSet(rsCab)
    Call Close_RecordSet(RsDetLar)
    Call Close_RecordSet(RsDet)
    Call Close_RecordSet(RsDocRef)
    Exit Function
Solucion:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "�Desea ver mas informaci�n?", vbExclamation + vbYesNo, "Mensaje de Error") = vbYes Then
            Call MsgBox(Err.Description, vbCritical, "Mensaje de Error")
            Err.Clear
        End If
    End If
End Function
Private Sub PrintData(Obj As Object)
    Dim sCad As String, lPosI As Long, lPosF As Long, lLen As Long, lYPos As Long, dTotal As Double, CurY As Long
    Dim dY As Integer, sSeries As String, sCurY As Long, sTmp() As String, lTmp As Long, lCnt As Long, dImp001 As Double
    Dim dX As Double, dyP As Double
    With Obj
        dX = 0: dY = 0
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Draft 17cpi"
        Obj.Font.Size = 9
                
        Obj.Height = 5817.32
        Obj.Width = 7245.66
        
        .CurrentX = 4 + dX
        .CurrentY = 8 + dY
        Obj.Print Trim$("" & rsCab.Fields("CodDoc")) & " - " & Trim$("" & rsCab.Fields("NumDoc").Value)
        
        .CurrentX = 4 + dX
        .CurrentY = 12 + dY
        Obj.Print Format$(Now, "dd/mm/yyyy") & " " & Format(GetValorScalar("GetDate() ServerDate"), "HH:MM:SS am/pm")
        
        If RsDet.RecordCount > 0 Then RsDet.MoveFirst
        .CurrentY = 20 + dY
        
        Do Until RsDet.EOF
            CurY = .CurrentY
            .CurrentX = 1 + dX
            .CurrentY = CurY
            
            If dImp001 = 0 Then dImp001 = Val(Trim$("" & RsDet.Fields("Imp001").Value))
                        
            .CurrentX = 10 - .TextWidth(Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")) + dX
            .CurrentY = CurY
            Obj.Print Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")

            .CurrentX = 108 - .TextWidth(Format("" & Trim("" & RsDet.Fields("PreVta").Value), "#,0.00")) + dX
            .CurrentY = CurY
            Obj.Print Format("" & Trim("" & RsDet.Fields("PreVta").Value), "#,0.00")
            
            .CurrentX = 124 - .TextWidth(Format("" & Trim("" & RsDet.Fields("PreIte").Value), "#,0.00")) + dX
            .CurrentY = CurY
            Obj.Print Format("" & Trim("" & RsDet.Fields("PreIte").Value), "#,0.00")
            
            sCad = Trim("" & RsDet.Fields("CodArt").Value) & " " & Trim("" & RsDet.Fields("DesArt").Value)
            Call Cortar(sCad, 40, True, False, sTmp(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 15 + dX
                Obj.CurrentY = CurY
                Obj.Print sTmp(lCnt)
                If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt

            dTotal = dTotal + Format("" & Trim("" & RsDet.Fields("PreIte").Value), "#,0.00")
            
            CurY = CurY + 4
                
            RsDet.MoveNext
        Loop
        
        dyP = -160
    
        .CurrentX = 5 + dX
        .CurrentY = 233 + dY + dyP
        Obj.Print Trim$(" " & rsCab!TOTDOCL)
        
        .CurrentX = 124 - .TextWidth(Format$(Trim$(" " & rsCab!impnet), "#,0.00")) + dX
        .CurrentY = 241 + dY + dyP
        Obj.Print Format$(Trim$(" " & rsCab!impnet), "#,0.00")
        
        'I.G.V
        .CurrentX = 105 - .TextWidth(Format(dImp001, "#,0.00")) + dX
        .CurrentY = 246 + dY + dyP
        Obj.Print Format(dImp001, "#,0.00")
    
        .CurrentX = 124 - .TextWidth(Format$(Trim$(" " & rsCab!totimp1), "#,0.00")) + dX
        .CurrentY = 246 + dY + dyP
        Obj.Print Format$(Trim$(" " & rsCab!totimp1), "#,0.00")
        
        'TOTAL DE LA FACTURA
        .CurrentX = 104 - .TextWidth(Trim$(" " & rsCab!SIMMON)) + dX
        .CurrentY = 251 + dY + dyP
        Obj.Print Trim$(" " & rsCab!SIMMON)
        
        .CurrentX = 124 - .TextWidth(Format$(Trim$(" " & rsCab!TotDoc), "#,0.00")) + dX
        .CurrentY = 251 + dY + dyP
        Obj.Print Format$(Trim$(" " & rsCab!TotDoc), "#,0.00")
    End With
End Sub

Private Function GetValorScalar(ByVal sCampo As String, Optional ByVal sTabla As String = "", Optional ByVal sCondicion As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = ICnx.Execute("Select " & sCampo & IIf(Trim(sTabla) <> "", " From " & sTabla, "") & IIf(Trim(sCondicion) <> "", " Where " & sCondicion, ""), , adCmdText)
    If Not rsTmp.EOF Then
        GetValorScalar = Trim("" & rsTmp.Fields(0).Value)
    End If
    Call Close_RecordSet(rsTmp)
End Function

