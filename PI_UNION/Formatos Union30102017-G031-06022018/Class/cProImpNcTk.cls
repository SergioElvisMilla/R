VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpNcTk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oRsC As ADODB.Recordset
Private dMonEnt As Double
Private dVueltoMn As Double
Private dVueltoMe As Double
Private dTipMov As String
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
Private sMoneda As String
Private sCondPago As String
Private sNv01 As String
Private sPagoEfe As String
Private sVueltoEfe As String
Private sPagoTv As String
Private sPagoTm As String
Private bPagoTv As Boolean
Private bPagoTm As Boolean
Private bPagoEfe As Boolean
Private INivel As Byte
Private ICnx As ADODB.Connection
Dim rsCab As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Dim RsDet2 As ADODB.Recordset
Dim rsEmp As ADODB.Recordset
Dim rsEmpLoc As ADODB.Recordset
Dim RsDetLar As ADODB.Recordset
Dim RsDocRef As ADODB.Recordset
Dim RsDocRefPunt As ADODB.Recordset
Dim RsCodDocPunt As ADODB.Recordset
Dim RsCabDocPunt As ADODB.Recordset
Dim RsAcuDocPunt As ADODB.Recordset
Dim RsCodDocPunt1 As ADODB.Recordset
Dim bFlagAdelanto As Boolean
Private rsDocRef2 As ADODB.Recordset
Private rsDocRef3 As ADODB.Recordset
Private oRsDoc As ADODB.Recordset
Private RsSer As ADODB.Recordset
Private RsAde As ADODB.Recordset
Private bConAde As Boolean
Private bImpConDetalle As Boolean
Private bImpConCondiciones As Boolean
Private bImpSinFiltro As Boolean
Private sComentario As String
Private ICodDoc As String
Private sRndRegNeg001 As String 'CanTot
Private sRndRegNeg002 As String 'PreVta
Private sRndRegNeg003 As String 'CanTot * PreVta
Private sRndRegNeg004 As String 'Descuento
Private sRndRegNeg005 As String 'Impuestos
Private sRndTot As String 'Impuestos
Private oSysAdmin As Object
Private bImpAde As Boolean
Private bImpFnl As Boolean

Property Set SysAdmin(ByRef oObj As Object)
    Set oSysAdmin = oObj
End Property

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
    Dim CODDOCPAD2, NUMDOCPAD2, CFGDOCPADCAB, CFGDOCPADDET, NOMANEPAD As String
    Screen.MousePointer = 11: bPrint = True
    
    Set oRsC = New ADODB.Recordset
    
    Set rsEmp = ICnx.Execute("SELECT TOP 1 NOMEMP,DESEMP,DIREMP,TELEMP,BKUPEMP,RUCEMP,RUtLOG,CIUDAD FROM EMPRESAS WHERE CODEMP='" & ICodEmp & "'", , adCmdText)
    
    Set rsEmpLoc = ICnx.Execute("SELECT TOP 1 MTD.DESAUX AS DIRLOC,DESAUX2 AS TELFLOC,CFG.SERIMP AS SERIMP,MTD.CTA02 AS NROAUT FROM CFGDOC00 CFG INNER JOIN MAETABDET MTD ON CFG.XLOCAL=MTD.CODITE " & vbCrLf & _
                                 "WHERE CODTAB = 'LOCAL' AND CODDOC='" & ICodDoc & "'", , adCmdText)
                            
                            
    Set rsCab = ICnx.Execute("SELECT TOP 1 CAB.TIPCAM, CAB.XTIPMON, CAB.NUMLOC,LOC.DIRLOC,UBL.DESCRIPCION AS UBILOC,CFG.SERRIE AS SERIE,CAB.XTIPMOV,UBIGEO.DESCRIPCION AS XUBIGEO,CAB.CODDOC,CFG.DESCRI AS DESDOC,CAB.NUMDOC, CAB.NUMREA,CAB.FECDOC,CAB.FECVEN,CAB.CODANE,CAB.NOMANE, " & vbCrLf & _
                             "CAB.DIRANE, CAB.TELANE, CAB.IDEANE, CAB.CODRES, CAB.CODVEN,ANE.NOMANE AS NOMVEN,Res.NomAne AS CAJERO,ANE1.UBIZON AS CODUBIZON, CAB.TIPTR1, CAB.NUMRF1, CAB.TIPTR2,CAB.NUMRF2,UBI.NOMBRE AS UBIZON, " & vbCrLf & _
                             "CAB.SUBTOT - (CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) AS IMPNET, CAB.TOTIMP1+CAB.TOTIMP2+CAB.TOTIMP3+CAB.TOTIMP4 AS TOTIMP, CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4 AS TOTDSC, " & vbCrLf & _
                             "CAB.SUBTOT,CAB.TOTIMP1,CAB.TOTIMP2,CAB.TOTIMP3,CAB.TOTIMP4,CAB.TOTDSC1,CAB.TOTDSC2,CAB.TOTDSC3,CAB.TOTDSC4,CAB.TOTDOC, MAE.DESITE AS CONPAG,MAE1.DESITE AS FORPAG,CAB.NUMDIA,CAB.CODMEM,CAB.COMMEM, " & vbCrLf & _
                             "CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS SIMMON, CAB.TOTDOCL, CfgP.SerRie SerRieP, mV05.DesIte Tipo_D, Ven.NomAne NomVen, Cab.xEstDoc, Cab.IdMesa, Me.Descripcion MesaD" & vbCrLf & _
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
                             "Left Join (Select CodDocPad,CodDoc,NumDoc From MovCtaArt Group By CodDocPad,CodDoc,NumDoc) Tmp On Tmp.CodDoc = Cab.CodDoc And Tmp.NumDoc = Cab.NumDoc " & vbCrLf & _
                             "Left Join CfgDoc00 CfgP On CfgP.CodDoc = Tmp.CodDocPad And CfgP.CodEmp = '" & Trim(ICodEmp) & "' " & vbCrLf & _
                             "Left Join MaeTabDet mV05 On mV05.CodIte = Cab.v05 And mV05.CodTab = 'vCab05' And mV05.CodFil = 'A'" & vbCrLf & _
                             "Left Join Anexo Ven On Ven.CodAne = Cab.CodVen" & vbCrLf & _
                             "Left Join Anexo Res On Res.CodAne = Cab.CodRes" & vbCrLf & _
                             "Left Join Mesa Me On Me.CodEmp='" & ICodEmp & "' And Me.IdMesa=Cab.IdMesa" & vbCrLf & _
                             "WHERE Cab.CodDoc = '" & Trim(ICodDoc) & "' And Cab.NUMDOC='" & INumDoc & "'", , adCmdText)
    
    Call Default_Reglas_Negocio
    If Not rsCab.EOF Then Call Reglas_Negocio(Trim$("" & rsCab.Fields("FecDoc").Value))
     
    Set RsDet = ICnx.Execute("SELECT Md.CodSubAlm, Md.NumIte , MD.CodArt, A.CODARTANT1, Md.DesArt, Md.DesArt1, Md.xTipUni, Mae.DesIte As DesUni, " & vbCrLf & _
                             "Md.Dsc001, Md.Dsc002, Md.Dsc003, Md.Dsc004,A.PrecMn , Md.PreVta,Md.Imp002 , Md.Imp001," & vbCrLf & _
                             "Case When Md.CanTot != 0 Then Md.CanTot * Md.PreVta/Md.CanTot Else Md.PreVta End AS PreVta1, Md.CanTot , Md.PreVta*Md.CanTot As PreIte, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*(1-((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100)))) As TotDsc, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100))) As PreTot, Md.TotArt, " & vbCrLf & _
                             "(SELECT  X20.NOMSUBART FROM  XSUBARTDET X20 WHERE  X20.XSUBART='XSUBART20' AND X20.CODSUB=AD.XSUBART20 )+' '+" & vbCrLf & _
                             "(SELECT  X02.NOMSUBART FROM  XSUBARTDET X02 WHERE  X02.XSUBART='XSUBART02' AND X02.CODSUB=AD.XSUBART02 )+' '+" & vbCrLf & _
                             "(SELECT  X07.NOMSUBART FROM  XSUBARTDET X07 WHERE  X07.XSUBART='XSUBART07' AND X07.CODSUB=AD.XSUBART07 )+' '+" & vbCrLf & _
                             "(SELECT  X06.NOMSUBART FROM  XSUBARTDET X06 WHERE  X06.XSUBART='XSUBART06' AND X06.CODSUB=AD.XSUBART06 )AS NEWCOD" & vbCrLf & _
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
                                
           If UCase(Trim("" & rsCab.Fields("xEstDoc"))) <> "A" Then
                                Set RsDocRef = ICnx.Execute("Select isnull(Mca02.CodDocPad,'') as CodDocOp, isnull(Mca02.NumDocPad,'') as NumDocOp" & vbCrLf & _
                                "From MovCtaArt Mca01" & vbCrLf & _
                                "Left Join MovCtaArt Mca02 On Mca02.CodDoc = Mca01.CodDocPad And Mca02.NumDoc = Mca01.NumDocPad" & vbCrLf & _
                                "Where Mca01.CodDoc = '" & Trim(ICodDoc) & "' And Mca01.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf & _
                                "Group By Mca02.CodDocPad, Mca02.NumDocPad", , adCmdText)
                                
                 If RsDocRef.BOF = True Then
                                Set RsDocRef = ICnx.Execute("Select isnull(Mca01.CodDocPad,'') as CodDocOp, isnull(Mca01.NumDocPad,'') as NumDocOp" & vbCrLf & _
                                "From MovCtaArt Mca01" & vbCrLf & _
                                "Where Mca01.CodDoc = '" & Trim(ICodDoc) & "' And Mca01.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf & _
                                "Group By Mca01.CodDocPad, Mca01.NumDocPad", , adCmdText)
                                CODDOCPAD2 = Trim(RsDocRef.Fields("CodDocOp"))
                                NUMDOCPAD2 = Trim(RsDocRef.Fields("NumDocOp"))
                 Else
                 
                                CODDOCPAD2 = Trim(RsDocRef.Fields("CodDocOp").Value)
                                NUMDOCPAD2 = Trim(RsDocRef.Fields("NumDocOp").Value)
                 End If
            Else
                        Dim xcodp, xnump, xdocpag1 As String
                        Set RsDocRef = ICnx.Execute("select (mc01.TIPTR1) as TIPTR1,(mc01.NUMRF1) as NUMRF1" & vbCrLf & _
                        "FROM " & IMovCab & " AS mc01 " & vbCrLf & _
                        "Where Mc01.CodDoc = '" & Trim(ICodDoc) & "' And Mc01.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf & _
                        "Group By Mc01.tiptr1, Mc01.numrf1", , adCmdText)
                        xcodp = Trim(RsDocRef.Fields("TIPTR1").Value)
                        xnump = Trim(RsDocRef.Fields("NUMRF1").Value)
                        
                        Set RsCodDocPunt1 = ICnx.Execute("SELECT ARCCAB,ARCDET from CFGDOC00 WHERE CODDOC='" & xcodp & "'", , adCmdText)
                                                         xdocpag1 = Trim(RsCodDocPunt1.Fields("ARCCAB").Value)
                        
                        Set RsDocRef = ICnx.Execute("select (mc02.TIPTR1) as TIPTR1,(mc02.NUMRF1) as NUMRF1" & vbCrLf & _
                        "FROM " & IMovCab & " AS mc01 " & vbCrLf & _
                        "left join " & xdocpag1 & " AS mc02 on mc01.TIPTR1=mc02.CODDOC and mc01.NUMRF1=mc02.NUMDOC " & vbCrLf & _
                        "Where Mc01.CodDoc = '" & Trim(ICodDoc) & "' And Mc01.NumDoc = '" & Trim(INumDoc) & "'")
                                CODDOCPAD2 = Trim(RsDocRef.Fields("TIPTR1").Value)
                                NUMDOCPAD2 = Trim(RsDocRef.Fields("NUMRF1").Value)
            
            End If
            
    Set RsDocRefPunt = ICnx.Execute("select CODEMP,CODDOC,NUMDOC,INGSAL,CODANE,IDTAR,PUNTOS,FECHIS,IDTIP,FLKACT from MovTarDoc" & vbCrLf & _
                                     "Where CodDoc = '" & CODDOCPAD2 & "' AND NUMDOC= '" & NUMDOCPAD2 & "' AND INGSAL='I'", , adCmdText)
    
    Set RsCodDocPunt = ICnx.Execute("SELECT ARCCAB,ARCDET from CFGDOC00 WHERE CODDOC='" & CODDOCPAD2 & "'", , adCmdText)
        
           CFGDOCPADCAB = Trim(RsCodDocPunt.Fields("ARCCAB").Value)
           CFGDOCPADDET = Trim(RsCodDocPunt.Fields("ARCDET").Value)
    
    Set RsCabDocPunt = ICnx.Execute("SELECT ISNULL(NV01,'')AS NV01,CODANE,TOTDOC FROM " & CFGDOCPADCAB & " WHERE CODDOC='" & CODDOCPAD2 & "' AND NUMDOC='" & NUMDOCPAD2 & "'", adCmdText)
    
            NOMANEPAD = Trim("" & RsCabDocPunt.Fields("CODANE").Value)
    
    Set RsAcuDocPunt = ICnx.Execute("SELECT CODEMP,IDTAR,IDTIP,CODSUBALM,CODART,FECINI,FECFIN,PUNTOS,FLKACT FROM TARJETA WHERE CODANE ='" & NOMANEPAD & "'", adCmdText)
    
                                
    Set oRsDoc = ICnx.Execute("Select Md.CodDoc, Md.NumDoc, Md.NumIte, M01.CodDocPad CodDoc01, M01.NumDocPad NumDoc01, M02.CodDocPad CodDoc02, M02.NumDocPad NumDoc02, M03.CodDocPad CodDoc03, M03.NumDocPad NumDoc03" & vbCrLf & _
                              "From " & Trim(IMovDet) & " Md" & vbCrLf & _
                              "Inner Join MovCtaArt M01 On M01.CodDoc = Md.CodDoc And M01.NumDoc = Md.NumDoc And M01.NumIteCta = Md.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Inner Join MovCtaArt M02 On M02.CodDoc = M01.CodDocPad And M02.NumDoc = M01.NumDocPad And M02.NumIteCta = M01.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Inner Join MovCtaArt M03 On M03.CodDoc = M02.CodDocPad And M03.NumDoc = M02.NumDocPad And M03.NumIteCta = M02.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Where Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf, , adCmdText)
    
                
                
    If Not rsCab.EOF Then
    
        Call GetDatos("RI01','RI02','RI03','RI04','RI05','RI06','RI07','RI08','RI09','RI10','RI11','RI12','RI13','RI14','RI15','RI16','RI17','RI18','RI19','RI20','RI21','RI22','RI23','RI24", ICodDoc, INumDoc, dMonEnt, dVueltoMn, dVueltoMe)
        Call Get_RecordSet_Can(oRsC, "RI01','RI02','RI03','RI04','RI05','RI06','RI07','RI08','RI09','RI10','RI11','RI12','RI13','RI14','RI15','RI16','RI17','RI18','RI19','RI20','RI21','RI22','RI23','RI24", ICodDoc, INumDoc)
            
        Dim pvWindow As Object, prt As Printer
        Set pvWindow = CreateObject("PrtTools.cPreView")
        Set prt = GetPrinter(sDevice)
        
        If Not prt Is Nothing Then Set Printer = prt
        Set prt = Nothing
        Set pvWindow.ActivePrinter = Printer
        
        If PreView Then
            Call PreviewData(pvWindow.ActivePage)
            Screen.MousePointer = 0
            bPrint = pvWindow.ShowPreView()
        End If
        If bPrint Then
            Call PreviewData(pvWindow.ActivePrinter)
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
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "¿Desea ver mas información?", vbExclamation + vbYesNo, "Mensaje de Error") = vbYes Then
            Call MsgBox(Err.Description, vbCritical, "Mensaje de Error")
            Err.Clear
        End If
    End If
End Function
Private Sub PreviewData(Obj As Object)
    Dim sCad As String, lPosI As Long, lPosF As Long, lLen As Long, lYPos As Long, dTotal As Double, CurY As Long, CurY2 As Long
    Dim dY As Integer, sSeries As String, sCurY As Long, StmP() As String, lTmp As Long, lCnt As Long, dImp001 As Double
    Dim dX As Double, dyP As Double, xTmp As Long, xTotArt As Double, xCanTot As Double, xPreVta As Double, xSubTot As Double, xTotDsc As Double, xTotImp As Double, dImp01 As Double
    Dim bAnulado As Boolean, sIdPrt As String, dDsc01 As Double, dImp02 As Double
            Dim xPreVta2, dDescT, dSubTot As Double
            Dim xdscx01, xdscx02, xdscx03, xdscx04 As Double
            Dim dDsc02, dDsc03, dDsc04 As Double
            Dim Tdscx01, Tdscx02, Tdscx03, Tdscx04 As Double
            Dim xsubtot1, xsubtot2, xsubtot3, xsubtot4, Totxdscx, TotxSubTot As Double
            Dim TCANT, DY2 As Integer
            Dim CONTADOR As Integer
    With Obj
        dX = 0: dY = 0
        CurY = 0
        .ScaleMode = vbMillimeters
        .Font.Name = "Draft 17cpi"
        .Font.Size = 7
        '7500
        Obj.Height = 8000 + ((RsDet.RecordCount * 100) + 5000) '2100
        Obj.Width = 4251.968503937
        
        'If UCase(Trim("" & rsCab.Fields("xEstDoc"))) = "A" Then
        '    .CurrentX = 37 - ((.TextWidth("******ANULADO******")) / 2) + dX
        '    .CurrentY = 3 + dY
        '    Obj.Print "******ANULADO******"
        '    dY = dY + 8
        '    bAnulado = True
        'End If
        
        'NOMBRE DE LA EMPRESA
        .CurrentX = 20
        .CurrentY = 3 + dY
        Obj.Print Trim$("" & rsEmp.Fields("NomEmp").Value)
        
        'DIRECCION DE LA EMPRESA POR LOCAL
        sCad = Trim(Left(Trim$("" & rsEmpLoc.Fields("DIRLOC").Value), 120))
            Call Cortar(sCad, 50, True, False, StmP(), lTmp, False)
            For lCnt = 1 To lTmp
                  Obj.CurrentX = 1 + dX
                  Obj.CurrentY = 7 + CurY
                  Obj.Print StmP(lCnt)
            If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt
            CurY = CurY + 4
        'DIRECCION DEL DOMICILIO FISCAL
        sCad = "D.F :  " & Trim(Left(Trim$("" & rsEmp.Fields("DIREMP").Value), 120))
            Call Cortar(sCad, 52, True, False, StmP(), lTmp, False)
            For lCnt = 1 To lTmp
                  Obj.CurrentX = 1 + dX
                  Obj.CurrentY = 7 + CurY
                  Obj.Print StmP(lCnt)
            If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt
            
        'CIUDAD DE LA EMPRESA POR LOCAL
        .CurrentX = 30
        .CurrentY = 11 + CurY
        Obj.Print UCase(Trim$("" & rsEmp.Fields("CIUDAD").Value))

        'RUC DE LA EMPRESA
        .CurrentX = 1
        .CurrentY = 15 + CurY
        Obj.Print "RUC: " & Trim$("" & rsEmp.Fields("RUCEMP").Value)

        'SERIE Y NUMERO DE DOCUMENTO
        .CurrentX = 35
        .CurrentY = 15 + CurY
        Obj.Print "Nota Credito : " & Trim$("" & rsCab.Fields("Serie").Value) & "-" & Trim$("" & rsCab.Fields("NumDoc").Value)

        'SERIE DE LA ETIQUETERA
        .CurrentX = 1
        .CurrentY = 19 + dY + CurY
        Obj.Print "N°DE SERIE: " & Trim$("" & rsEmpLoc.Fields("SERIMP").Value)

        'NUMERO DE AUTORIZACION
        .CurrentX = 1
        .CurrentY = 23 + dY + CurY
        Obj.Print "N°DE AUTORIZACION: " & Trim$("" & rsEmpLoc.Fields("NroAut").Value)

        'TELEFONO DEl LOCAL
        .CurrentX = 1
        .CurrentY = 27 + dY + CurY
        Obj.Print "TELF:" & Trim$("" & rsEmpLoc.Fields("TELFLOC").Value) & " " & Trim$("" & rsEmp.Fields("BKUPEMP").Value)

        'DESCRIPCION DE LA EMPRESA
        .CurrentX = 1
        .CurrentY = 31 + dY + CurY
        Obj.Print Trim$("" & rsEmp.Fields("DesEmp").Value)

        'FECHA DEL DOCUMENTO
        .CurrentX = 1
        .CurrentY = 35 + dY + CurY
        Obj.Print "FECHA:" & Trim$("" & rsCab.Fields("FecDoc").Value)

        'CAJA
        .CurrentX = 35
        .CurrentY = 35 + dY + CurY
        Obj.Print "CAJA: 01-TM"

        'CAJERO
        .CurrentX = 1
        .CurrentY = 39 + dY + CurY
        Obj.Print "CAJERO: " & UCase(Trim$("" & rsCab.Fields("CAJERO").Value))

        'VENDEDOR
        .CurrentX = 1
        .CurrentY = 43 + dY + CurY
        Obj.Print "VENDEDOR:" & UCase(Trim$("" & rsCab.Fields("NomVen").Value))
        
        'HORA
        .CurrentX = 20 + dX
        .CurrentY = 47 + dY + CurY
        Obj.Print "HORA : " & Time

        .CurrentX = 1 + dX
        .CurrentY = 51 + dY + CurY
        Obj.Print "..........................................................................."
        
                
        If RsDet.RecordCount > 0 Then RsDet.MoveFirst
        .CurrentY = 55 + dY
         CurY = .CurrentY + 4 '55
         
        Do Until RsDet.EOF
            Dim xPreTot As Double
            'CALCULO DE DESCUENTOS Y MONTOS
            dImp01 = Trim("" & RsDet.Fields("Imp001").Value)
            dDsc01 = Trim("" & RsDet.Fields("Dsc001").Value)
            dDsc02 = Trim("" & RsDet.Fields("Dsc002").Value)
            dDsc03 = Trim("" & RsDet.Fields("Dsc003").Value)
            dDsc04 = Trim("" & RsDet.Fields("Dsc004").Value)

            xCanTot = Format(Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)
            xPreVta = Format(Trim("" & RsDet.Fields("PreVta").Value), sRndRegNeg002)
'            xTotDsc = Trim("" & RsDet.Fields("TotDsc").Value)
'            xPreTot = Trim("" & RsDet.Fields("PreTot").Value)
'            xTotArt = Trim("" & RsDet.Fields("TotArt").Value)
            
            xPreVta2 = xPreVta * (1 + (dImp01 / 100))
            xSubTot = Format(xCanTot * xPreVta2, sRndRegNeg003)  'total precio incluido igv
            
            xdscx01 = Format((dDsc01 / 100), sRndRegNeg004)
            xdscx02 = Format((dDsc02 / 100), sRndRegNeg004)
            xdscx03 = Format((dDsc03 / 100), sRndRegNeg004)
            xdscx04 = Format((dDsc04 / 100), sRndRegNeg004)
            
            Tdscx01 = xdscx01 * xPreVta2   'total descuento
            Tdscx02 = xdscx02 * xPreVta2   'total descuento
            Tdscx03 = xdscx03 * xPreVta2  'total descuento
            Tdscx04 = xdscx04 * xPreVta2  'total descuento
            
            
            Totxdscx = (Tdscx01 + Tdscx02 + Tdscx03 + Tdscx04)
            xsubtot1 = (xSubTot - Totxdscx)
            TotxSubTot = (xsubtot1)
            

            'CANTIDAD
            .CurrentX = 1 + dX
            .CurrentY = CurY + 4
            Obj.Print Format("" & xCanTot, sRndRegNeg001)
            
            'PRECIO DEL ARITUCLO
            .CurrentX = 50 + dX
            .CurrentY = CurY + 4
            Obj.Print Trim("" & rsCab.Fields("SIMMON").Value) & Format(xSubTot, "#,0.00") 'correcto
            
            'ARTICULO
            sCad = Trim(Left(Trim("" & RsDet.Fields("NewCod").Value), 120))
            Call Cortar2(sCad, 30, True, False, StmP(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 4 + dX
                Obj.CurrentY = CurY + 4
                Obj.Print StmP(lCnt)
                If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt
            
            CurY = CurY + 4
            'PRECIO DEL ARITUCLO UNIDAD
             .CurrentX = 10 + dX
             .CurrentY = CurY + 4
             Obj.Print "PRECIO UNIDAD :  " & Trim("" & rsCab.Fields("SIMMON").Value) & "  " & Format((xPreVta2), "#,0.00") 'correcto
            
            
            CONTADOR = CONTADOR + 1
            CurY = CurY + 4
            TCANT = TCANT + xCanTot
            
            dSubTot = dSubTot + xSubTot
            dDescT = dDescT + Totxdscx
            dTotal = dTotal + Format(TotxSubTot, sRndRegNeg003)
            

            RsDet.MoveNext
        Loop
        CurY = CurY + 4
        .CurrentX = 1 + dX
        .CurrentY = CurY
        Obj.Print "......................................................................................"

        
        If bAnulado Then dY = dY - 8
        
        If dDescT > 0 Or UCase(Trim("" & rsCab.Fields("CONPAG").Value)) = "EFECTIVO" Then
        
        'TOTAL CANTIDAD
        .CurrentX = 1 + dX
        .CurrentY = CurY + 4 + dY
        Obj.Print "TOTAL CANTIDAD  :  " & TCANT
        
'        CurY = CurY + 4
        
'        'SUBTOTAL
'        .CurrentX = 1 + dX
'        .CurrentY = CurY + 4 + dY
'        Obj.Print "SUBTOTAL:  " & Trim("" & rsCab.Fields("SIMMON").Value) & Format(dSubTot, "#,0.00")
'
'        'DESCUENTO
'        .CurrentX = 1 + dX
'        .CurrentY = CurY + 8 + dY
'        Obj.Print "DESCUENTO: " & Trim("" & rsCab.Fields("SIMMON").Value) & Format(dDescT, "#,0.00")
        
        'TOTAL DE DOCUMENTO
        .CurrentX = 4 + dX
        .CurrentY = CurY + 12 + dY
        Obj.Print "TOTAL: " & Trim("" & rsCab.Fields("SIMMON").Value) & Format(dTotal, "#,0.00")
        CurY = CurY + 4
        
        Else
    
        'TOTAL DE DOCUMENTO
        .CurrentX = 4 + dX
        .CurrentY = CurY + 8 + dY
        Obj.Print "TOTAL: " & Trim("" & rsCab.Fields("SIMMON").Value) & Format(dTotal, "#,0.00")
        CurY = CurY + 4
        
        End If
        
        
        
        
        
        
        
        .CurrentX = 1 + dX
        .CurrentY = CurY + 12
        Obj.Print "......................................................................................"
        
        If UCase(Trim("" & rsCab.Fields("IDEANE").Value)) = "SDGVAR" Then
        
        'NOMBRE DE CLIENTE
         Obj.CurrentX = 1 + dX
         Obj.CurrentY = CurY + 16
         Obj.Print "CLIENTE: "

        
        Else
        
        'DNI DE CLIENTE
        .CurrentX = 1 + dX
        .CurrentY = CurY + 16
        Obj.Print "DNI: " & Trim("" & rsCab.Fields("IDEANE").Value)
        
        'NOMBRE DE CLIENTE
         Obj.CurrentX = 1 + dX
         Obj.CurrentY = CurY + 20
         CurY2 = CurY + 20
            
            sCad = "CLIENTE: " & Trim("" & rsCab.Fields("NOMANE").Value)
            Call Cortar(sCad, 50, True, False, StmP(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 1 + dX
                Obj.CurrentY = CurY2
                Obj.Print StmP(lCnt)
                If lTmp > 1 And lCnt < lTmp Then CurY2 = CurY2 + 4
            Next lCnt

            'DIRECCION DE CLIENTE
                Obj.CurrentX = 1 + dX
                Obj.CurrentY = CurY + 24
                CurY2 = CurY2 + 4

            sCad = "DIRECCION: " & Trim("" & rsCab.Fields("DIRANE").Value)
            Call Cortar(sCad, 50, True, False, StmP(), lTmp, False)
            
            For lCnt = 1 To lTmp
                Obj.CurrentX = 1 + dX  '32
                Obj.CurrentY = CurY2
                Obj.Print StmP(lCnt)
                If lTmp > 1 And lCnt < lTmp Then CurY2 = CurY2 + 4
            Next lCnt
        End If
        
        'DETALLE DE PIE
        '--------------------------------------
        
        If CurY2 > CurY Then
        CurY = CurY + 4
        If (lCnt) - 1 >= lTmp Then
        CurY = CurY + 4
        End If
        Else
        CurY = CurY - 8
        End If
        


       'Imprimir el detalle de pago multiple y simple
       If oRsC.State <> adStateClosed Then
          If oRsC.RecordCount > 0 Then oRsC.MoveFirst
          
            Do Until oRsC.EOF
            sCondPago = ""
             sPagoEfe = "0.00"
             sPagoTv = "0.00"
             sPagoTm = "0.00"
             bPagoEfe = False
             bPagoTv = False
             bPagoTm = False
            'sCondPago & "/" &
            sCondPago = UCase(Trim("" & oRsC.Fields("XTIPMOV").Value))
            
                If (UCase(Trim("" & oRsC.Fields("XTIPMON").Value)) = "MN") Then
                 sMoneda = "S/."
                End If
                
                If (UCase(Trim("" & oRsC.Fields("XTIPMON").Value)) = "ME") Then
                 sMoneda = "US $."
                End If
                
                If (UCase(Trim("" & oRsC.Fields("XTIPMOV").Value)) = "EFE") Then
                 sPagoEfe = sMoneda & Format(oRsC.Fields("MonEnt").Value, "#,0.00")
                 sVueltoEfe = sMoneda & Format(oRsC.Fields("VueltoMn").Value, "#,0.00")
                 bPagoEfe = True
                End If
                
                If (UCase(Trim("" & oRsC.Fields("XTIPMOV").Value)) = "TCV") Then
                 sPagoTv = sMoneda & Format(oRsC.Fields("MonEnt").Value, "#,0.00")
                 bPagoTv = True
                End If
                
                If (UCase(Trim("" & oRsC.Fields("XTIPMOV").Value)) = "TCM") Then
                 sPagoTm = sMoneda & Format(oRsC.Fields("MonEnt").Value, "#,0.00")
                 bPagoTm = True
                End If
                
                .CurrentX = 1 + dX
                .CurrentY = CurY + 28 + dY
                Obj.Print "FORMA DE PAGO: " & sCondPago

                'Pagos multiples incluyendo efectivo
                DY2 = CurY + 32 + dY
                If bPagoTv = True Then
                    'Monto Tarjeta Visa
                    .CurrentX = 1 + dX
                    .CurrentY = DY2
                    Obj.Print "TV: " & sPagoTv
                    DY2 = DY2 + 4
                    CurY = CurY + 4
                End If
                If bPagoTm = True Then
                    'Monto Tarjeta Master Card
                    .CurrentX = 1 + dX
                    .CurrentY = DY2
                    Obj.Print "TM: " & sPagoTm
                    DY2 = DY2 + 4
                    CurY = CurY + 4
                End If
                If bPagoEfe = True Then
                    'Monto Pago Efectivo
                    .CurrentX = 1 + dX
                    .CurrentY = DY2
                    Obj.Print "EF: " & sPagoEfe
                    DY2 = DY2 + 4
                    CurY = CurY + 4
                    'VUELTO
                    .CurrentX = 1 + dX
                    .CurrentY = DY2
                    Obj.Print "VUELTO: " & sVueltoEfe
                    DY2 = DY2 + 4
                    CurY = CurY + 4
                End If
                DY2 = DY2 + 4
                CurY = CurY + 4
                oRsC.MoveNext
            Loop
           'End If
        End If
        
        
        
        
       .CurrentX = 1 + dX
       .CurrentY = CurY + 34 + dY
       Obj.Print "T.C. S/. " & Trim$("" & rsCab.Fields("TIPCAM").Value)
        
       'TARJETA BABY PUNTOS
     If (Trim$(sNv01)) = "1" Then
        .CurrentX = 1 + dX
        .CurrentY = CurY + 38 + dY
        Obj.Print ".......BABY PUNTOS........"
        
        .CurrentX = 1 + dX
        .CurrentY = CurY + 42 + dY
        Obj.Print "Puntos Por Esta Compra:" & Trim("" & RsDocRefPunt.Fields("PUNTOS").Value)
        
        
        .CurrentX = 1 + dX
        .CurrentY = CurY + 46 + dY
        Obj.Print "Puntos Acumulados:" & Trim("" & RsAcuDocPunt.Fields("PUNTOS").Value)
        
        .CurrentX = 1 + dX
        .CurrentY = CurY + 50 + dY
        Obj.Print "........................"
        
        CurY = CurY + 18
       End If
        
    .CurrentX = 1 + dX
    .CurrentY = CurY + 38 + dY
     Obj.Print "Item(S): " & CONTADOR
        
        
          
    .CurrentX = 1 + dX
    .CurrentY = CurY + 42 + dY
     Obj.Print "   GRACIAS POR SU COMPRA"
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 46 + dY
     Obj.Print "*CONSERVE ESTE TICKET PARA DEVOLUCIONES "
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 50 + dY
     Obj.Print " O CAMBIO DENTRO DE LOS 7 DIAS CALENDARIO"
         
    .CurrentX = 1 + dX
     .CurrentY = CurY + 54 + dY
     Obj.Print " A LA FECHA DE COMPRA."
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 58 + dY
     Obj.Print "* LAS DEVOLUCIONES SE REALIZARAN DE"
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 62 + dY
     Obj.Print " LUNES A VIERNES DE 9:00 AM A 5:00PM"
     

     .CurrentX = 1 + dX
     .CurrentY = CurY + 66 + dY
     Obj.Print "*NO SE ACEPTAN CAMBIO PARA PRODUCTOS"
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 70 + dY
     Obj.Print " COMPRADOS EN LIQUDACION."
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 74 + dY
     Obj.Print "*LOS PRODUCTOS EN PROMOCION NO ACUMULAN"
     
     .CurrentX = 1 + dX
     .CurrentY = CurY + 78 + dY
     Obj.Print " BABY PUNTOS."

  End With
End Sub

Private Sub GetDatos(ByVal xCodDocCan As String, ByVal xCodDoc As String, ByVal sNumDoc As String, ByRef dMonEnt As Double, ByRef dVueltoMn As Double, ByRef dVueltoMe As Double)
    Dim oRs As ADODB.Recordset, xCodDocP As String, xNumDocP As String, xArcCab As String
    Set oRs = ICnx.Execute("Select Top 1 Mpd.CodDocDes, Mpd.NumDocDes, Cfg.ArcCab, Cfg.ArcDet From MovPlaDet Mpd Inner Join CfgDoc00 Cfg On Cfg.CodEmp = Mpd.CodEmp And Cfg.CodDoc = Mpd.CodDocDes Where Mpd.CodEmp = '" & Trim(ICodEmp) & "' And Mpd.CodDocOrg = '" & Trim(xCodDoc) & "' And Mpd.NumDocOrg = '" & Trim(sNumDoc) & "' And Mpd.CodDocDes In ('" & Replace(Trim(xCodDocCan), ";", "','") & "');", , adCmdText)
    If Not oRs.EOF Then
        xCodDocP = Trim("" & oRs.Fields("CodDocDes").Value)
        xNumDocP = Trim("" & oRs.Fields("NumDocDes").Value)
        xArcCab = Trim("" & oRs.Fields("ArcCab").Value)
        Call Close_RecordSet(oRs)
        Set oRs = ICnx.Execute("Select MonEnt, VueltoMn, VueltoMe, XTIPMOV From " & xArcCab & " Where CodDoc = '" & Trim(xCodDocP) & "' And NumDoc = '" & Trim(xNumDocP) & "';", , adCmdText)
        If Not oRs.EOF Then
            dMonEnt = Val(Trim("" & oRs.Fields("MonEnt").Value))
            dVueltoMn = Val(Trim("" & oRs.Fields("VueltoMn").Value))
            dVueltoMe = Val(Trim("" & oRs.Fields("VueltoMe").Value))
            dTipMov = Trim("" & oRs.Fields("XTIPMOV").Value)
        End If
    End If
    Call Close_RecordSet(oRs)
End Sub

Private Function GetValorScalar(ByVal sCampo As String, Optional ByVal sTabla As String = "", Optional ByVal sCondicion As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = ICnx.Execute("Select " & sCampo & IIf(Trim(sTabla) <> "", " From " & sTabla, "") & IIf(Trim(sCondicion) <> "", " Where " & sCondicion, ""), , adCmdText)
    If Not rsTmp.EOF Then
        GetValorScalar = Trim("" & rsTmp.Fields(0).Value)
    End If
    Call Close_RecordSet(rsTmp)
End Function

Private Sub Default_Reglas_Negocio()
    If Trim(sRndRegNeg001) = "" Then sRndRegNeg001 = "#,0.00"
    If Trim(sRndRegNeg002) = "" Then sRndRegNeg002 = "#,0.00"
    If Trim(sRndRegNeg003) = "" Then sRndRegNeg003 = "#,0.00"
    If Trim(sRndRegNeg004) = "" Then sRndRegNeg004 = "#,0.00"
    If Trim(sRndRegNeg005) = "" Then sRndRegNeg005 = "#,0.00"
    If Trim(sRndTot) = "" Then sRndTot = "#,0.00"
End Sub

Private Sub Reglas_Negocio(ByVal FecDoc As String)
    Call oSysAdmin.Reglas_Negocios("Reg0000001", , sRndRegNeg001, , FecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000002", , sRndRegNeg002, , FecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000003", , sRndRegNeg003, , FecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000004", , sRndRegNeg004, , FecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000005", , sRndRegNeg005, , FecDoc)
End Sub

Private Sub Get_RecordSet_Can(ByRef oRsCan As ADODB.Recordset, ByVal xCodDocCan As String, ByVal xCodDoc As String, ByVal sNumDoc As String)
    Dim oRs As ADODB.Recordset, xCodDocP As String, xNumDocP As String, xArcCab As String, sSql As String
    Set oRs = ICnx.Execute("Select Mpd.CodDocDes, Mpd.NumDocDes, Cfg.ArcCab, Cfg.ArcDet From MovPlaDet Mpd Inner Join CfgDoc00 Cfg On Cfg.CodEmp = Mpd.CodEmp And Cfg.CodDoc = Mpd.CodDocDes Where Mpd.CodEmp = '" & Trim(ICodEmp) & "' And Mpd.CodDocOrg = '" & Trim(xCodDoc) & "' And Mpd.NumDocOrg = '" & Trim(sNumDoc) & "' And Mpd.CodDocDes In ('" & Replace(Trim(xCodDocCan), ";", "','") & "') Group By Mpd.CodDocDes, Mpd.NumDocDes, Cfg.ArcCab, Cfg.ArcDet;", , adCmdText)
    Do Until oRs.EOF
        xCodDocP = Trim("" & oRs.Fields("CodDocDes").Value)
        xNumDocP = Trim("" & oRs.Fields("NumDocDes").Value)
        xArcCab = Trim("" & oRs.Fields("ArcCab").Value)
        sSql = sSql & _
               "Union All " & _
               "Select Mc.MonEnt, Mc.VueltoMn, Mc.VueltoMe, Mc.XtipMon, Mc.xTipMov, Mtd.DesIte xTipMov_D" & vbCrLf & _
               "From " & xArcCab & " Mc" & vbCrLf & _
               "Left Join MaeTabDet Mtd On Mtd.CodIte = Mc.xTipMov And Mtd.CodFil = 'A' And Mtd.CodTab = 'MOV-CH'" & vbCrLf & _
               "Where Mc.CodDoc = '" & Trim(xCodDocP) & "' And Mc.NumDoc = '" & Trim(xNumDocP) & "'" & vbCrLf
        oRs.MoveNext
    Loop
    sSql = Mid(sSql, 11)
    If Trim(sSql) <> "" Then
        Set oRsCan = ICnx.Execute(sSql, , adCmdText)
    End If
    Call Close_RecordSet(oRs)
End Sub



