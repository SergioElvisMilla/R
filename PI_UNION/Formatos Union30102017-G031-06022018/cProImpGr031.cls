VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpGr031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Private sMOVCABE1 As String
Private sCodDoc As String, sTIPTR1 As String, sNUMRF1 As String, sNumDoc As String, sProImp As String
Private sCodEmp As String, sCodGru As String, sTitFrm As String
Private sMovCab As String, sMovDet As String, sDevice As String
Private pai As String, dep As String, pro As String, dis As String, CodUbi As String, referencia As String
Private sParam1 As String, INivel As Byte, ICnx As ADODB.Connection
Private Const NFIL_ = 35
Private sNomDoc As String, sCodLoc As String, sNomLoc As String, sNomSuc As String, sDirLoc As String, sUbiLoc As String
Private sTelLoc As String
Private oSysAdmin As Object
Private sComentario As String
Private bPrint As Boolean
Dim RsDirAne As ADODB.Recordset
Dim RsLoc As ADODB.Recordset
Dim RsCab As ADODB.Recordset, RsCabLoc As ADODB.Recordset, RsEmp As ADODB.Recordset, RsDetLar As ADODB.Recordset
Dim Rspai As ADODB.Recordset, rsdep As ADODB.Recordset, rspro As ADODB.Recordset, rsdis As ADODB.Recordset, rsreferencia As ADODB.Recordset
Dim pvWindow As Object
Dim nPag As Long, bDet As Boolean, bCant As Boolean
Dim Cx As Long, CXX As Long
Dim Cy As Long
Public Property Let TIPTR1(ByVal Cad As String)
    sTIPTR1 = Trim$(Cad)
End Property
Public Property Let NUMRF1(ByVal Cad As String)
    sNUMRF1 = Trim$(Cad)
End Property

Property Set SysAdmin(ByRef Valor As Object)
    Set oSysAdmin = Valor
End Property
Property Get SysAdmin() As Object
    Set SysAdmin = oSysAdmin
End Property
Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set ICnx = Cnn
End Property
Public Property Let CodDoc(ByVal Cad As String)
    sCodDoc = Trim$(Cad)
End Property
Public Property Let NumDoc(ByVal Cad As String)
    sNumDoc = Trim$(Cad)
End Property
Public Property Let ProImp(ByVal Cad As String)
    sProImp = Cad
End Property
Public Property Let CodEmp(ByVal Cad As String)
    sCodEmp = Trim$(Cad)
End Property
Public Property Let CodGru(ByVal Cad As String)
    sCodGru = Trim$(Cad)
End Property
Public Property Let MovCab(ByVal Cad As String)
    sMovCab = Trim$(Cad)
End Property
Public Property Let MovDet(ByVal Cad As String)
    sMovDet = Trim$(Cad)
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
Public Sub PrintShow()
    Call PrintDoc(True)
End Sub
Public Function PrintDoc(Optional PreView As Boolean = False)
    Dim bImpDesArt1 As Boolean
    Dim campoB As String
    Dim Pos As Integer
    Dim iLonDet As Integer 'Longitud maxima por linea de Articulo
    Dim Valor As String
    Dim sCadAux As String
    Dim lPosI As Integer
    Dim lPosF As Integer
    Dim lLen As Integer
    Dim sCad As String
    Dim rpt As Object
    Dim Igv As String, sTmp() As String, lTmp As Long, lCnt As Long
    Dim bPrint As Boolean
    Dim bImpRec As Boolean
    Dim rsTmp As ADODB.Recordset
    bPrint = True
    Set rsTmp = ICnx.Execute("Select Top 1 Descri,xlocal From CfgDoc00 Where CodEmp='" & sCodEmp & "' And CodDoc='" & sCodDoc & "'")
    If Not rsTmp.EOF Or Not rsTmp.BOF Then sNomDoc = Trim$("" & rsTmp!DESCRI): sCodLoc = Trim$("" & rsTmp!xlocal)
    rsTmp.Close: Set rsTmp = Nothing
    Set rsTmp = ICnx.Execute("Select Top 1 DESITE,DESAUX,DESAUX1,Param2,DESAUX2 From MAETABDET Where CodTab='LOCAL' And CodIte='" & sCodLoc & "'")
    If Not rsTmp.EOF Or Not rsTmp.BOF Then sNomLoc = Trim$("" & rsTmp!DesIte): sNomSuc = Trim$("" & rsTmp!DesAux1): sDirLoc = Trim$("" & rsTmp!DesAux): sUbiLoc = Trim$("" & rsTmp!Param2): sTelLoc = Trim$("" & rsTmp!DESAUX2)
    rsTmp.Close: Set rsTmp = Nothing
    Set rsTmp = ICnx.Execute("Select Top 1 * From " & sMovCab & " " & _
                            "Where CodDoc = '" & sCodDoc & "' And NumDoc = '" & sNumDoc & "' And xEstDoc != 'A'", , adCmdText)
    If rsTmp.EOF Or rsTmp.BOF Then
        Screen.MousePointer = 11
        Call MsgBox("No se encontro el documento", vbExclamation)
        rsTmp.Close: Set rsTmp = Nothing: Exit Function
    End If
    Screen.MousePointer = 11
    '=============================
    Set RsEmp = ICnx.Execute("Select Top 1 CodEmp,NomEmp,DirEmp,TelEmp,FAXEmp,BKUPEmp,RUCEmp,RUtLOG,EMAIL,CIUDAD From EMPRESAS Where CodEmp='" & sCodEmp & "'", , adCmdText)
    If RsEmp.EOF And RsEmp.BOF Then
        Screen.MousePointer = 0
        MsgBox "No se puede localizar la empresa", vbExclamation, "Error"
        Exit Function
    End If
   '================================
   '================================
    Set pvWindow = CreateObject("PrtTools.cPreView")
    Dim prt As Printer
    Set prt = GetPrinter(sDevice)
    If Not prt Is Nothing Then Set Printer = prt
    Set prt = Nothing
    Set pvWindow.ActivePrinter = Printer
    pvWindow.ScaleMode = vbMillimeters
    If PreView Then
        Call PrintData(pvWindow.ActivePage, True)
        Screen.MousePointer = 0
        bPrint = pvWindow.ShowPreView()
    End If
    If bPrint = True Then
        Call PrintData(pvWindow.ActivePrinter, False)
        pvWindow.ActivePrinter.EndDoc
        Screen.MousePointer = 0
    End If
    Set pvWindow = Nothing
    rsTmp.Close: Set rsTmp = Nothing
    RsEmp.Close: Set RsEmp = Nothing
    PrintDoc = True
End Function
Private Sub PrintData(Obj As Object, PreView As Boolean)
    Dim sMOVCABE1 As String
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long
    Dim RsDet As ADODB.Recordset, RsCab As ADODB.Recordset, RsPrecio As ADODB.Recordset, RsSec As ADODB.Recordset, sSqlCab As String, sSqlLoc As String, sSqlCabLoc As String, sSqlDirAne As String, sSqlDet As String, sSqlSec As String, lTop As Long
    Dim sLeft As Single, lItem As Integer, sRow As Single, sRowY As Long, sPos(5) As Single, sAumRow As Single, sMaxHeight As Single
    Dim dSubTot As Double, sLmtPag As Single, Col As Long, sLnx() As String, lLnx As Long
    Dim X As Integer, Y As Integer, GRAV As Double, EXO As Double, GRAT As Double, DETRA As Double
    GRAV = 0: EXO = 0: GRAT = 0: DETRA = 0
    sMOVCABE1 = "MOVCABE1" 'agregado para concatenar cabecera de OP
    Cx = -5 - 5: Cy = -4
    sLmtPag = 220
    
    Obj.Font.Name = "Draft 17cpi"

    Screen.MousePointer = 11
    sLeft = 32: sAumRow = 4: sMaxHeight = 245
    
    sPos(0) = sLeft - 15: sPos(1) = sLeft + 8: sPos(2) = sLeft: sPos(3) = sLeft + 105: sPos(4) = sLeft + 127: sPos(5) = sLeft + 159
    Obj.Font.Name = "Draft 17cpi": Obj.Width = 8910: Obj.ScaleMode = 6: Obj.Height = 16000 '15000
    Printer.Font.Name = "Draft 17cpi": Printer.ScaleMode = 6: Printer.Font.Size = 9: Printer.Font.Bold = True
    lTop = 63
                       

    sSqlCab = "Select top (1) Mc.CodDoc,Mc.NumDoc,ANE1.XTIPIDE1,mc.NomAne,mc.IdeAne as Ruc,ANE1.XUBIGEO AS UBIGEO, mc.TelAne, isnull(mc.FecDoc,'')FecDoc,isnull(mc.Fecven,'')FecVen, isnull(Mc.NUMREA,'') as OrdenCliente,isnull(Mae.Desite,'') as CondPago,isnull(COMMEM,'')COMMEM,isnull(CODMEM,'')CODMEM,(case when isnull(Mc.TipTr4,'')<>'' then upper(Mc.TipTr4) when isnull(Mc.TipTr1,'')<>'' then upper(Cfg2.SerImp) else upper(isnull(Cta.CODDOCPAD,'')) End) as CodDocGuia ,(case when isnull(Mc.NumRf4,'')<>'' then Mc.NumRf4 when isnull(Mc.NumRf1,'')<>'' then Mc.NumRf1 else isnull(Cta.NUMDOCPAD,'') End) as NumDocGuia,TipTr1,NumRf1,Mc.DirAne as DirAne, RefInt, isnull(mc.CodVen,'')CodVen,isnull(mc.GLOSA,'')GLOSA, NumRf2, NumRf2, IsNull(TotDsc1,0) TotDsc1, IsNull(TotDsc2,0) TotDsc2, IsNull(TotDsc3,0) TotDsc3, IsNull(TotDsc4,0) TotDsc4,(IsNull(SubTot,0)-IsNull(TotDsc1,0)+IsNull(TotDsc2,0)+IsNull(TotDsc3,0)+IsNull(TotDsc4,0)) SubTot, (TotImp1) TotImp, TotDoc, TotDocL, " & _
            " isnull(TipCam,'')TipCam, Case When (TotDoc-TotImp1) = 0 Then 0 Else Case When Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) = Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) Then Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) Else Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) End End Igv, Mtd.DesAux Sig,Mtd.CTA01 as MONEDAS,MtI.CTA01,ISNULL(Cfg.SerImp,'')SERIE, An3.IdeAne1 RucTra, Con.NomCon, Con.TelCon AS Marca, Con.mailcon AS CertifMTC, Con.CelCon AS Placa, Con.DniCon  AS Licencia,Mc.codane02 as CodAneTra,Mc.NomAne02 AS NomAneTra " & _
            " From " & sMovCab & " Mc" & _
            " Left Join MovCtaArt Cta On cta.coddoc=Mc.CodDoc and cta.NUMDOC=Mc.NumDoc" & _
            " Left Join ANEXO ANE1 ON Mc.CODANE = ANE1.CODANE" & _
            " Left Join Anexo AN2 On  Mc.CodVen = AN2.CodAne " & _
            " Left Join Anexo AN3 On  Mc.CodAne02 = AN3.CodAne " & _
            " Left Join MaeTabDet Mtd On Mtd.CodTab = 'xTipMon' And Mtd.CodIte = Mc.xTipMon" & _
            " Left Join MaeTabDet MtI On MtI.CodTab = 'XTIPIDE' And MtI.CodIte = ANE1.XTIPIDE1" & _
            " Left Join MaeTabDet Mae On Mae.CodTab = 'ConPag' And Mae.CodIte = Mc.xConPag" & _
            " Left Join CfgDoc00 Cfg On Cfg.CodEmp='" & sCodEmp & "' And Cfg.CodDoc='" & sCodDoc & "'" & _
            " Left Join CfgDoc00 Cfg2 On Cfg2.CodEmp='" & sCodEmp & "' And Cfg2.CodDoc=Mc.TipTr1" & _
            " Left Join Contactos CON On Con.CodAne = An3.CodAne And Mc.NumCon02 = Con.NumCon " & _
            " Where Mc.CodDoc = '" & sCodDoc & "' And Mc.NumDoc = '" & sNumDoc & "' And xEstDoc != 'A'"
            
         
    sSqlDet = "Select MD.CODART,MD.DESART,MD.CANTOT,MD.IMP001,MD.TOTART,ISNULL(MD.XTIPUNI,'')XTIPUNI,MD.PREVTA,(MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004)DESCUENTO,(MD.PREVTA*MD.CANTOT)-((MD.PREVTA*MD.CANTOT) * (CASE WHEN(MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004)> 0 THEN (MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004) ELSE 0 END/100))TOTALSINIGV,MD.TOTART,MD.V01,(ISNULL(A.PORDET,0)/100)DETRACCION " & _
            " From " & sMovDet & " Md" & _
            " LEFT JOIN ARTICULO A ON A.CODART=MD.CODART AND A.CODSUBALM=MD.CODSUBALM " & _
            " Where Md.CodDoc = '" & sCodDoc & "' And Md.NumDoc = '" & sNumDoc & "' order by Md.numite"
    

    Set RsDetLar = ICnx.Execute("Select NumDetLar,DetLar Comentario From MovDetLar " & _
                                " Where CodEmp = '" & sCodEmp & "' And CodDoc = '" & sCodDoc & "' And NumDoc='" & sNumDoc & "'  ", adCmdText)

    RsDetLar.Filter = "NumDetLar = '1'"

    If Not RsDetLar.EOF Then sComentario = Trim$("" & RsDetLar.Fields("Comentario").Value)

    Set RsCab = New ADODB.Recordset: Set RsCab = ICnx.Execute(sSqlCab, , adCmdText)
    Set RsDet = New ADODB.Recordset: Set RsDet = ICnx.Execute(sSqlDet, , adCmdText)
    
    
    
    If Not RsCab.EOF Then
    sSqlCabLoc = "Select top (1) Rtrim(Mc.CodAne)CodAne ,IsNull(Mc.NumLoc,'') NumLoc ,A.DirAne+' - ' +isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,5)),'')+ '/' +  isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,7)),'')+ '/' + isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,10)),'') AS DirAne1, " & _
            "A.XUBIGEO,RTRIM(Mc.TipTr1)+' - '+RTRIM(Mc.NUMRF1) AS OP," & _
            "CASE Mc.XTIPMON WHEN 'MN' THEN 'SOLES' ELSE 'DOLARES' END AS TipMon,Mc.CodDoc,Mc.NumDoc,mc.NomAne,mc.IdeAne as Ruc, mc.TelAne, isnull(mc.FecDoc,'')FecDoc,isnull(mc.FecEnt,'')FecVen,Mc.xTipMov, isnull(Mc.NUMREA,'') as OrdenCliente,isnull(Mae.Desite,'') as CondPago,isnull(COMMEM,'')COMMEM,isnull(CODMEM,'')CODMEM,(case when isnull(Mc.TipTr4,'')<>'' then upper(rtrim(Mc.TipTr4)) when isnull(Mc.TipTr1,'')<>'' then upper(rtrim(Mc.TipTr1)) else upper(isnull(rtrim(Cta.CODDOCPAD),'')) End) as CodDocGuia , (case when isnull(Mc.TipTr4,'')<>'' then upper(rtrim(Mc.TipTr4)) when isnull(Mc.TipTr1,'')<>'' then upper(rtrim(Mc.TipTr1)) else upper(isnull(rtrim(Cta.CODDOCPAD),'')) End) + ' - ' +(case when isnull(Mc.NumRf4,'')<>'' then Mc.NumRf4 when isnull(Mc.NumRf1,'')<>'' then Mc.NumRf1 else isnull(Cta.NUMDOCPAD,'') End) as NumDocGuia, " & _
            "TipTr1,NumRf1, RefInt, isnull(mc.CodVen,'')CodVen, NumRf2, NumRf2, IsNull(TotDsc1,0) TotDsc1, IsNull(TotDsc2,0) TotDsc2, IsNull(TotDsc3,0) TotDsc3, IsNull(TotDsc4,0) TotDsc4,(IsNull(SubTot,0)-IsNull(TotDsc1,0)+IsNull(TotDsc2,0)+IsNull(TotDsc3,0)+IsNull(TotDsc4,0)) SubTot, (TotImp1) TotImp, TotDoc, TotDocL, " & _
            " isnull(Mc.TipCam,'')TipCam,Mc.xTipMon, Case When (TotDoc-TotImp1) = 0 Then 0 Else Case When Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) = Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) Then Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) Else Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) End End Igv, Mtd.DesAux Sig,ISNULL(Cfg.SerRie,'')SERIE," & _
            " Case When Mc.XTIPMON='ME' Then Mc.TOTDOC* Mc.TipCam Else Mc.TOTDOC End as TOTDET " & _
            " From " & sMOVCABE1 & Trim$("" & RsCab.Fields("TIPTR1").Value) & " Mc" & _
            " Left Join MovCtaArt Cta On cta.coddoc=Mc.CodDoc and cta.NUMDOC=Mc.NumDoc" & _
            " Left Join MaeTabDet Mtd On Mtd.CodTab = 'xTipMon' And Mtd.CodIte = Mc.xTipMon" & _
            " Left Join MaeTabDet Mae On Mae.CodTab = 'ConPag' And Mae.CodIte = Mc.xConPag" & _
            " LEFT JOIN ANEXO A ON MC.CODANE=A.CODANE" & _
            " LEFT JOIN UBIGEO UB ON A.XUBIGEO=UB.UBIGEO" & _
            " Left Join CfgDoc00 Cfg On Cfg.CodEmp='" & sCodEmp & "' And Cfg.CodDoc='" & sCodDoc & "'" & _
            " Left Join ANEXO AS ANE1 ON Mc.CODANE = ANE1.CODANE " & _
            " Where Mc.CodDoc = '" & Trim$("" & RsCab.Fields("TIPTR1").Value) & "' And Mc.NumDoc = '" & Trim$("" & RsCab.Fields("NUMRF1").Value) & "' And xEstDoc != 'A'"
            End If
            
    Set RsCabLoc = New ADODB.Recordset: Set RsCabLoc = ICnx.Execute(sSqlCabLoc, , adCmdText)
       
        If Not RsCabLoc.EOF Then
    sSqlLoc = "Select CodAne,NumLoc,NomLoc,TelLoc,xUbigeo,xConLoc,CodLocAnt,DirLoc + ' - ' + isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(L.XUBIGEO,5)),'')" & _
              "+ '/' +  isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(L.XUBIGEO,7)),'')+ '/' + " & _
              "isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(L.XUBIGEO,10)),'') as DirLoc " & _
              "From Locales L LEFT JOIN UBIGEO UB ON L.XUBIGEO=UB.UBIGEO " & _
              "Where CodAne='" & Trim$("" & RsCabLoc.Fields("CodAne").Value) & "' and NumLoc ='" & Trim$("" & RsCabLoc.Fields("NumLoc").Value) & "'"
    End If
    
    Set RsLoc = New ADODB.Recordset: Set RsLoc = ICnx.Execute(sSqlLoc, , adCmdText)

    sSqlDirAne = "select DIRANE from anexo where CODANE='" & Trim$("" & RsCabLoc.Fields("CodAne").Value) & "'"
    
    Set RsDirAne = New ADODB.Recordset: Set RsDirAne = ICnx.Execute(sSqlDirAne, , adCmdText)
    
'    CodUbi = Trim$("" & rsCab.Fields("UBIGEO").Value)
'    Set Rspai = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,3)='" & Left(CodUbi, 3) & "'  ", adCmdText)
'        If Trim$("" & Rspai.Fields("DESCRIPCION").Value) = "" Then pai = "" Else pai = " -" + Trim$("" & Rspai.Fields("DESCRIPCION").Value)
'    Set rsdep = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,5)='" & Left(CodUbi, 5) & "'  ", adCmdText)
'        If Trim$("" & rsdep.Fields("DESCRIPCION").Value) = "" Then dep = "" Else dep = " - " + Trim$("" & rsdep.Fields("DESCRIPCION").Value)
'    Set rspro = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,7)='" & Left(CodUbi, 7) & "'  ", adCmdText)
'        If Trim$("" & rspro.Fields("DESCRIPCION").Value) = "" Then pro = "" Else pro = " - " + Trim$("" & rspro.Fields("DESCRIPCION").Value)
'    Set rsdis = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,9)='" & Left(CodUbi, 9) & "'  ", adCmdText)
'        If Trim$("" & rsdis.Fields("DESCRIPCION").Value) = "" Then dis = "" Else dis = " - " + Trim$("" & rsdis.Fields("DESCRIPCION").Value)
''------------------------------------------
'    Set rsreferencia = ICnx.Execute("Select DISTINCT tipdoc_p,num_p  FROM " & sMovDet & " Where CodDoc = '" & sCodDoc & "' And NumDoc='" & sNumDoc & "'  ", adCmdText)
'    Dim n As Long
'    n = 0
'    Do Until rsreferencia.EOF
'       n = n + 1
'       If n = 1 Then
'          If Len(rsreferencia.Fields("tipdoc_p").Value) > 1 Then
'            referencia = Trim$("" & rsreferencia.Fields("tipdoc_p").Value) + Trim$("-" & rsreferencia.Fields("num_p").Value)
'            referencia = referencia
'          Else
'            referencia = ""
'          End If
'       End If
'       If n > 1 Then
'          If Len(referencia) > 4 Then
'            referencia = referencia + ", " + Trim$("" & rsreferencia.Fields("tipdoc_p").Value) + Trim$("-" & rsreferencia.Fields("num_p").Value)
'          Else
'            referencia = Trim$("" & rsreferencia.Fields("tipdoc_p").Value) + Trim$("-" & rsreferencia.Fields("num_p").Value)
'          End If
'       End If
'        rsreferencia.MoveNext
'    Loop
'    If Len(referencia) > 1 Then
'        referencia = referencia
'        Else
'        referencia = ""
'    End If
'
'------------------------------------
    
   If RsCab.EOF Then Exit Sub
    
    If TypeName(Obj) = "PictureBox" Then
        Obj.Width = 12000: Obj.Height = 15200 '9500
    Else
    Obj.PaperSize = 1
     'Obj.PaperSize = 134  cambio momentaneo para impresora de facruracion
    End If
    
    sRow = 85: Call Encabezado(Obj, RsCab, sLeft, lTop)
    
    RsCab.MoveFirst
    RsDet.MoveFirst

    Obj.Font.Size = 8
'    Dim Pag As Long, P As Long
'    P = 0: sRowY = 93
'    Do Until RsDet.EOF
'        Call Cortar(Trim$("" & RsDet.Fields("DesArt").Value), 87, False, False, sLn(), lLn, True, 0)
'        sAumRow = 0
'        For lCnt = 1 To lLn
'            If lCnt > 1 Then sAumRow = sAumRow + 4
'        Next lCnt
'        sRowY = sRowY + sAumRow + 4
'    RsDet.MoveNext
'    Loop
    RsDet.MoveFirst
   
'    If sRowY > 180 Then Pag = (sRowY / 225) + 1 Else Pag = sRowY / 180
'    If Pag = 0 Then Pag = 1 Else Pag = Pag
    ''=============================================================================
    ''=======================DETALLE DEL DOCUMENTO=================================
    ''=============================================================================
     Obj.Font.Size = 8: Obj.Font.Bold = True
    Dim sPosXTot(5) As Single
    sPosXTot(0) = 15: sPosXTot(1) = 42: sPosXTot(2) = 125: sPosXTot(3) = 140: sPosXTot(4) = 155: sPosXTot(5) = 175
'    sRow = 174
    Dim CYC As Long, CXC As Long
    CXC = 15
    
    Obj.CurrentX = Cx + 165 - Printer.TextWidth("UM"): Obj.CurrentY = Cy + 74: Obj.Print "UM"
    
    ''=============================================================================
    ''=============================================================================
    sRow = sRow - 7
        'P = 1
        Cx = Cx
        Dim PRECIOTARI As Double
        Dim CODI As String
    Do Until RsDet.EOF
       Obj.Font.Size = 8: Obj.Font.Bold = False
                       ''CANTIDAD DE ARTICULOS
                        CXC = 15: sAumRow = 0
                        Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!CODART)
                        
                        ''DESCRIPCION DE ARTICULOS
                        Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
                        
                        Obj.CurrentX = Cx + 55 - Printer.TextWidth(Format(RsDet.Fields("CANTOT").Value, "#,0.00")): Obj.CurrentY = Cy + sRow
                        Obj.Print Format(RsDet.Fields("CANTOT").Value, "#,0.00")
                        
'                        Call Cortar(Trim$("" & RsDet.Fields("DesArt").Value), 87, False, False, sLn(), lLn, True, 0)
'                        For lCnt = 1 To lLn
'                            If lCnt > 1 Then sAumRow = sAumRow + 4
'                            Obj.CurrentX = Cx + 50: Obj.CurrentY = Cy + sRow + sAumRow: Obj.Print sLn(lCnt)
'                        Next lCnt
                        Obj.CurrentX = Cx + 60: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!DESART)

                        
                        Obj.CurrentX = Cx + 162: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!xTipUni)
            
                        ''PRECIO UNITARIO
                        Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
                        'Obj.CurrentX = Cx + 153: Obj.CurrentY = Cy + sRow: Obj.Print rsCab!Sig
'                        CODI = Trim$(RsDet!CODART)
'                        Set RsPrecio = ICnx.Execute("Select PREVEN From TARIFDET Where  CODTAR='TAR003' AND CODART='" & CODI & "'", adCmdText)
'                        If RsPrecio.RecordCount() = 0 Then
'                        PRECIOTARI = 0
'                        Else
                        PRECIOTARI = IIf(IsNull(Trim(RsDet!PREVTA)), "", Trim(RsDet!PREVTA))
'                       End If
                       ' If Trim(RsPrecio.Fields("PREVEN").Value) = "" Then PRECIOTARI = "" Else PRECIOTARI = Trim$(RsPrecio.Fields("PREVEN").Value)
                       ' Obj.CurrentX = Cx + 184 - Printer.TextWidth(Format(RsDet.Fields("PreVta").Value, "#,0.00")): Obj.CurrentY = Cy + sRow
                       ' Obj.Print Format(RsDet.Fields("PreVta").Value, "#,0.000")
                        Obj.CurrentX = Cx + 189 - Printer.TextWidth(Format(PRECIOTARI, "#,0.00")): Obj.CurrentY = Cy + sRow
                        Obj.Print Format(PRECIOTARI, "#,0.000")
                        
                
                        

                        sRow = sRow + sAumRow + 4
       
                RsDet.MoveNext
    Loop
Cx = Cx

    CYC = CYC + 263
    Obj.Font.Size = 10: Obj.Font.Bold = True
    Obj.CurrentX = Cx + 92: Obj.CurrentY = CYC: Obj.Print "1"

    Obj.Font.Size = 9: Obj.Font.Bold = False

'An3.IdeAne1 RucTra, Con.NomCon, Con.TelCon AS Marca, Con.mailcon AS CertifMTC, Con.CelCon AS Placa, Con.DniCon  AS Licencia,Mc.codane02 as CodAneTra,Mc.NomAne02 AS NomAneTra
    CYC = 243
    Obj.CurrentX = Cx + 117: Obj.CurrentY = CYC: Obj.Print Trim$("" & UCase(RsCab!NomAneTra))

    CYC = CYC + 4
    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "RUC"
    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!RucTra)

    CYC = CYC + 4
    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Conductor"
    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!NOMCON)


    CYC = CYC + 4
    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Licencia N°"
    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!LICENCIA)


    CYC = CYC + 4
    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Vehiculo"
    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!Marca)
    Obj.CurrentX = Cx + 145: Obj.CurrentY = CYC: Obj.Print "Placa :"
    Obj.CurrentX = Cx + 155: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!PLACA)

    CYC = CYC + 4
    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Certif. MTC"
    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab!CertifMTC)


    

   
    RsCab.Close: Set RsCab = Nothing
    RsDet.Close: Set RsDet = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub Encabezado(ByRef Obj As Object, ByRef RsCab As ADODB.Recordset, sLeft As Single, ByRef lTop As Long)
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long, iPst As Integer
    Dim CYL As Long, CYS As Long, CYC As Long, CYE As Long, sRutLog As String
    Dim CXL As Long, CXS As Long, CXC As Long, CXCD As Long, CXE As Long, Direccion As String, Direccion2 As String, Consignado2 As String, Direccion3 As String
    Dim fechaTexto As String
    Dim longi, inicio, final As Long
        
    CYL = 25 - 3: CYS = 25: CYC = 47: CYE = 40
    CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    
    If bPrint = True Then
        CXX = 5
        CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    Else
        CXX = 5
        CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    End If
    
    ''=============================================================================
    ''=======================LOGO DE LA EMPRESA====================================
    ''=============================================================================
'    If Trim$("" & RsEmp!rutlog) <> "" Then
'            sRutLog = Trim$("" & RsEmp!rutlog)
'            On Error GoTo erlog
 '          Dim std As StdPicture
'            Set std = GetLogo(sRutLog)
'            If Not std Is Nothing Then
'                Obj.PaintPicture std, 10, 10, 30, 30
'            End If
'    End If
    ''=============================================================================
    ''=======================DATOS DE LA EMPRESA===================================
    ''=============================================================================
'    Obj.Font.Size = 20: Obj.Font.Bold = True
'    Obj.CurrentX = 70 + Cx: Obj.CurrentY = 15: Obj.Print RsEmp!NomEmp
'
'    Obj.Font.Size = 10: Obj.Font.Bold = True
'    Obj.CurrentX = 76 + Cx: Obj.CurrentY = 24: Obj.Print RsEmp!DirEmp
'    Obj.Font.Size = 8: Obj.Font.Bold = True
'    Obj.CurrentX = 75.5 + Cx: Obj.CurrentY = 28: Obj.Print "LIMA - LIMA - LIMA - PERU"
'    Obj.Font.Size = 10: Obj.Font.Bold = True
'    Obj.CurrentX = 79 + Cx: Obj.CurrentY = 32: Obj.Print Trim$("Teléfono : " & RsEmp!TelEmp)
    

    ''=============================================================================
    ''=======================DATOS DE CABECERA=====================================
    ''=============================================================================
    CYC = 38
    Cx = Cx - 10
    Obj.Font.Size = 9: Obj.Font.Bold = False: Printer.Font.Size = 9: Printer.Font.Bold = False
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("FECHA")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("" & RsCab.Fields("FecDoc").Value)
    CYC = CYC + 4.5
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Doc. Identif.")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("" & RsCab.Fields("Ruc").Value);
    CYC = CYC + 4.5
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("CLIENTE")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("" & RsCab.Fields("NomAne").Value);
    Obj.CurrentX = Cx + CXC + 160: Obj.CurrentY = CYC - 3: Obj.Print Trim$("" & RsCab!SERIE) & "-" & Trim$("" & sNumDoc)

    CYC = CYC + 4.5
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("DIRECCIÓN")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    '    Direccion2 = Trim$(rsCab.Fields("DirAne").Value) + dis + pro + dep + pai
  '  Direccion2 = Trim$(RsCab.Fields("DirAne").Value)
    
   
        
              
              '  If IsNull(RsDirAne.Fields("DirAne").Value) = True Then
                   ' Direccion2 = ""
               ' Else
                  Direccion2 = Trim$(RsDirAne.Fields("DirAne").Value)
               ' End If
          
          
     
        Call Cortar(Direccion2, 90, False, False, sLn(), lLn, True, 0)
        For lCnt = 1 To lLn
            If lCnt > 1 Then CYC = CYC + 4.5
            Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print sLn(lCnt)
        Next lCnt
    
    CYC = CYC + 4.5
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("P. Partida")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Av Bolognesi Nro 551 Santa Anita")
    
    Obj.Font.Size = 14: Obj.Font.Bold = True: Printer.Font.Size = 14: Printer.Font.Bold = True
    If Trim$("" & RsCab!CodDoc) = "G043" Or Trim$("" & RsCab!CodDoc) = "G044" Then
        Obj.CurrentX = Cx + CXC + 155: Obj.CurrentY = CYC - 3: Obj.Print Trim$("GUIA DE CARGA")
    Else
        Obj.CurrentX = Cx + CXC + 155: Obj.CurrentY = CYC - 3: Obj.Print Trim$("GUIA DE REMISION")
    End If
    Obj.Font.Size = 9: Obj.Font.Bold = False: Printer.Font.Size = 9: Printer.Font.Bold = False
    CYC = CYC + 4.5
    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("P. Llegada")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    
    'CONDICION SI NO EXISTE DIRECCION Y SI NO HAY NUMERO DE LOCAL
    If IsNull(RsCab.Fields("DirAne").Value) = True Then
                  Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Destinos Varios")
                Else
                If (RsCabLoc.Fields("NumLoc").Value) = 0 Then
                Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Destinos Varios")
                Else
                Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print RsLoc.Fields("DirLoc").Value
                End If
                   
                End If
    
    
    'Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print RsLoc.Fields("DirLoc").Value
    CYC = CYC + 4.5

    Obj.CurrentX = Cx + CXC + 7: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Consignado")
    Obj.CurrentX = Cx + CXC + 27: Obj.CurrentY = CYC - 3: Obj.Print Trim$(":")
    Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print RsCab.Fields("COMMEM").Value
    
    
    
    ' COnsignado cortado
'        Consignado2 = Trim$(RsCab.Fields("COMMEM").Value)
'
'        Call Cortar(Consignado2, 180, False, False, sLn(), lLn, True, 0)
'        For lCnt = 1 To lLn
'            If lCnt > 1 Then CYC = CYC + 4.5
'            Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print sLn(lCnt)
'        Next lCnt
    'Terminad Consignado cORTADO
    
    
    
    
    'Obj.CurrentX = Cx + CXC + 29: Obj.CurrentY = CYC - 3: Obj.Print Trim$("Lima - Destinos Varios")
    
'    longi = Len(Trim(rsCab.Fields("OrdenCliente").Value))
'    inicio = Cx + CXC + 27
'    final = inicio + 22
'    Obj.CurrentX = Cx + CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC - 14: Obj.Print Trim(rsCab.Fields("OrdenCliente").Value)
    ''=============================================================================
    ''=============================================================================

erlog:
            
            Err.Clear
             'Set std = Nothing
    
End Sub
Public Sub CortarX(ByVal sCad As String, ByRef sLn() As String, ByRef lLn As Long, ByRef sLnx() As String, ByRef lLnx As Long)
Dim sTmp As String, lIni As Long, Lx As Long, Ly As Long: ReDim sLn(1 To 1) As String
Dim X As Integer
    sCad = Replace(sCad, vbTab, "")
  Do
        X = 0
        
        lIni = InStr(1, sCad, vbCrLf)
        If lIni = 0 Then
            sTmp = Trim$(sCad): sCad = ""
        Else
            sTmp = Trim$(Mid(sCad, 1, lIni - 1))
            sCad = Mid(sCad, lIni + 2)
        End If
        
        If (Len(sTmp) > 50) Then
            X = Len(sTmp) \ 50
            X = X + IIf(Len(sTmp) Mod 50 > 0, 1, 0)
            X = X - 1
        End If

            Lx = Lx + X + 1
            lLnx = lLnx + 1: ReDim Preserve sLnx(1 To lLnx)
            sLnx(lLnx) = sTmp
        
  Loop Until sCad = ""
End Sub















