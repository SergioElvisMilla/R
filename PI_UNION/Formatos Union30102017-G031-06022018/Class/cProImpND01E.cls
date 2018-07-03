VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpND01E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private sCodDoc As String, sNumDoc As String, sProImp As String
Private sCodEmp As String, sCodGru As String, sTitFrm As String
Private sMovCab As String, sMovDet As String, sDevice As String
Private sParam1 As String, INivel As Byte, ICnx As ADODB.Connection
Private Const NFIL_ = 35
Private sNomDoc As String, sCodLoc As String, sNomLoc As String, sNomSuc As String, sDirLoc As String, sUbiLoc As String
Private sTelLoc As String
Private oSysAdmin As Object
Private sComentario As String
Private bPrint As Boolean
Dim Doc As Integer
Dim RsCab As ADODB.Recordset, RsEmp As ADODB.Recordset, RsDetLar As ADODB.Recordset
Dim pvWindow As Object
Dim nPag As Long, bDet As Boolean, bCant As Boolean
Dim Cx As Long, CXX As Long
Dim Cy As Long
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
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long
    Dim RsDet As ADODB.Recordset, RsCab As ADODB.Recordset, RsSec As ADODB.Recordset, sSqlCab As String, sSqlDet As String, sSqlSec As String, lTop As Long
    Dim sLeft As Single, lItem As Integer, sRow As Single, sRowY As Long, sPos(5) As Single, sAumRow As Single, sMaxHeight As Single
    Dim dSubTot As Double, sLmtPag As Single, Col As Long, sLnx() As String, lLnx As Long
    Dim X As Integer, Y As Integer, GRAV As Double, EXO As Double, GRAT As Double, DETRA As Double
    GRAV = 0: EXO = 0: GRAT = 0: DETRA = 0
    
    Cx = -5 - 5: Cy = -4
    sLmtPag = 220
    
    Obj.Font.Name = "Draft 17cpi"
   ''If TypeName(Obj) <> "PictureBox" Then Obj.PaperSize = 256
    Screen.MousePointer = 11
    sLeft = 32: sAumRow = 4: sMaxHeight = 245
    sPos(0) = sLeft - 15: sPos(1) = sLeft + 8: sPos(2) = sLeft: sPos(3) = sLeft + 105: sPos(4) = sLeft + 127: sPos(5) = sLeft + 159
    Obj.Font.Name = "Draft 17cpi": Obj.Width = 8910: Obj.ScaleMode = 6: Obj.Height = 15000 '15000
    Printer.Font.Name = "Draft 17cpi": Printer.ScaleMode = 6: Printer.Font.Size = 9: Printer.Font.Bold = True
    lTop = 63
                       
'    sSqlCab = "Select top (1) Mc.CodDoc,Mc.NumDoc,NomAne,IdeAne as Ruc, TelAne, isnull(FecDoc,'')FecDoc,isnull(Fecven,'')FecVen, isnull(Mc.NUMREA,'') as OrdenCliente,isnull(Mae.Desite,'') as CondPago,isnull(COMMEM,'')COMMEM,isnull(CODMEM,'')CODMEM,(case when isnull(Mc.TipTr4,'')<>'' then Mc.TipTr4 else isnull(Cta.CODDOCPAD,'') End) as CodDocGuia ,(case when isnull(Mc.NumRf4,'')<>'' then Mc.NumRf4 else isnull(Cta.NUMDOCPAD,'') End) as NumDocGuia,TipTr1,NumRf1,DirAne, " & _
'            " isnull(RefInt,'')RefInt, isnull(CodVen,'')CodVen, NumRf2, NumRf2, IsNull(TotDsc1,0) TotDsc1, IsNull(TotDsc2,0) TotDsc2, IsNull(TotDsc3,0) TotDsc3, IsNull(TotDsc4,0) TotDsc4, IsNull(SubTot,0) SubTot, (TotImp1) TotImp, TotDoc, TotDocL, isnull(TipCam,'')TipCam, Case When (TotDoc-TotImp1) = 0 Then 0 Else Case When Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) = Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) Then Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) Else Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) End End Igv, Mtd.DesAux Sig " & _
'            " From " & sMovCab & " Mc" & _
'            " Left Join MovCtaArt Cta On cta.coddoc=Mc.CodDoc and cta.NUMDOC=Mc.NumDoc" & _
'            " Left Join MaeTabDet Mtd On Mtd.CodTab = 'xTipMon' And Mtd.CodIte = Mc.xTipMon" & _
'            " Left Join MaeTabDet Mae On Mae.CodTab = 'ConPag' And Mae.CodIte = Mc.xConPag" & _
'            " Where Mc.CodDoc = '" & sCodDoc & "' And Mc.NumDoc = '" & sNumDoc & "' And xEstDoc != 'A'"

''    sSqlCab = "Select top (1) Mc.CodDoc,Mc.NumDoc,mc.NomAne,mc.IdeAne as Ruc, mc.TelAne, isnull(mc.FecDoc,'')FecDoc,isnull(mc.Fecven,'')FecVen, isnull(Mc.NUMREA,'') as OrdenCliente,isnull(Mae.Desite,'') as CondPago,isnull(COMMEM,'')COMMEM,isnull(CODMEM,'')CODMEM,(case when isnull(Mc.TipTr4,'')<>'' then upper(Mc.TipTr4) when isnull(Mc.TipTr1,'')<>'' then upper(Cfg2.SerImp) else upper(isnull(Cta.CODDOCPAD,'')) End) as CodDocGuia ,(case when isnull(Mc.NumRf4,'')<>'' then Mc.NumRf4 when isnull(Mc.NumRf1,'')<>'' then Mc.NumRf1 else isnull(Cta.NUMDOCPAD,'') End) as NumDocGuia,TipTr1,NumRf1,(Mc.DirAne+ ' - ' + dbo.fnx_UbigeoComprobante(ANE1.XUBIGEO))as DirAne, RefInt, isnull(mc.CodVen,'')CodVen, NumRf2, NumRf2, IsNull(TotDsc1,0) TotDsc1, IsNull(TotDsc2,0) TotDsc2, IsNull(TotDsc3,0) TotDsc3, IsNull(TotDsc4,0) TotDsc4,(IsNull(SubTot,0)-IsNull(TotDsc1,0)+IsNull(TotDsc2,0)+IsNull(TotDsc3,0)+IsNull(TotDsc4,0)) SubTot, (TotImp1) TotImp, TotDoc, TotDocL, " & _

      sSqlCab = "Select top (1) A.DirAne +' - ' +isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,5)),'')+ '/' +  isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,7)),'')+ '/' + isnull((SELECT rtrim(DESCRIPCION) FROM UBIGEO WHERE UBIGEO=LEFT(A.XUBIGEO,10)),'') AS DirAne1, " & _
            "A.xUBIGEO,RTRIM(Mc.TipTr1)+'-'+RTRIM(Mc.NUMRF1) AS OP," & _
            "CASE Mc.XTIPMON WHEN 'MN' THEN 'SOLES' ELSE 'DOLARES' END AS TipMon,Mc.CodDoc,Mc.NumDoc,mc.NomAne,mc.IdeAne as Ruc, mc.TelAne, isnull(mc.FecDoc,'')FecDoc,isnull(mc.Fecven,'')FecVen,Mc.xTipMov, isnull(Mc.NUMREA,'') as OrdenCliente,isnull(Mae.Desite,'') as CondPago,isnull(Mam.Desite,'') as Motivo,isnull(COMMEM,'')COMMEM,isnull(CODMEM,'')CODMEM,(case when isnull(Mc.TipTr4,'')<>'' then upper(rtrim(Mc.TipTr4)) when isnull(Mc.TipTr1,'')<>'' then upper(rtrim(Mc.TipTr1)) else upper(isnull(rtrim(Cta.CODDOCPAD),'')) End) as CodDocGuia , (case when isnull(Mc.TipTr4,'')<>'' then upper(rtrim(Mc.TipTr4)) when isnull(Mc.TipTr1,'')<>'' then upper(rtrim(Mc.TipTr1)) else upper(isnull(rtrim(Cta.CODDOCPAD),'')) End) +' - '+(case when isnull(Mc.NumRf4,'')<>'' then Mc.NumRf4 when isnull(Mc.NumRf1,'')<>'' then Mc.NumRf1 else isnull(Cta.NUMDOCPAD,'') End) as NumDocGuia," & _
            "TipTr1,NumRf1, RefInt, isnull(mc.CodVen,'')CodVen, NumRf2, NumRf2, IsNull(TotDsc1,0) TotDsc1, IsNull(TotDsc2,0) TotDsc2, IsNull(TotDsc3,0) TotDsc3, IsNull(TotDsc4,0) TotDsc4,(IsNull(SubTot,0)-IsNull(TotDsc1,0)+IsNull(TotDsc2,0)+IsNull(TotDsc3,0)+IsNull(TotDsc4,0)) SubTot, (TotImp1) TotImp, TotDoc, TotDocL, " & _
            " isnull(Mc.TipCam,'')TipCam,Mc.xTipMon, Case When (TotDoc-TotImp1) = 0 Then 0 Else Case When Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) = Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) Then Convert(Decimal(18,0),100*((TotDoc/(TotDoc-TotImp1))-1)) Else Convert(Decimal(18,1),100*((TotDoc/(TotDoc-TotImp1))-1)) End End Igv, Mtd.DesAux Sig,ISNULL(Cfg.SerRie,'')SERIE," & _
            " Case When Mc.XTIPMON='ME' Then Mc.TOTDOC* Mc.TipCam Else Mc.TOTDOC End as TOTDET " & _
            " From " & sMovCab & " Mc" & _
            " Left Join MovCtaArt Cta On cta.coddoc=Mc.CodDoc and cta.NUMDOC=Mc.NumDoc" & _
            " Left Join MaeTabDet Mtd On Mtd.CodTab = 'xTipMon' And Mtd.CodIte = Mc.xTipMon" & _
            " Left Join MaeTabDet Mae On Mae.CodTab = 'ConPag' And Mae.CodIte = Mc.xConPag" & _
            " Left Join MaeTabDet Mam On Mam.CodTab = 'MOV-NC' And Mam.CodIte = Mc.xTipMov" & _
            " LEFT JOIN ANEXO A ON MC.CODANE=A.CODANE" & _
            " LEFT JOIN UBIGEO UB ON A.XUBIGEO=UB.UBIGEO" & _
            " Left Join CfgDoc00 Cfg On Cfg.CodEmp='" & sCodEmp & "' And Cfg.CodDoc='" & sCodDoc & "'" & _
            " Left Join ANEXO AS ANE1 ON Mc.CODANE = ANE1.CODANE " & _
            " Where Mc.CodDoc = '" & sCodDoc & "' And Mc.NumDoc = '" & sNumDoc & "' And xEstDoc != 'A'"
    sSqlDet = "Select MD.CODART,(A.DESART+' / '+A.CodArtAnt1 +' / '+ A.CodArtAnt2) AS DESART,MD.CANTOT,ISNULL(MD.XTIPUNI,'')XTIPUNI,MD.PREVTA,(MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004)DESCUENTO,(MD.PREVTA*MD.CANTOT)-((MD.PREVTA*MD.CANTOT) * (CASE WHEN (MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004)> 0 THEN (MD.DSC001+MD.DSC002+MD.DSC003+MD.DSC004) ELSE 0 END/100))TOTALSINIGV,MD.TOTART,MD.V01,(ISNULL(A.PORDET,0)/100)DETRACCION,ISNULL(A.XTIPART,'')XTIPART, " & _
            " Md.IngSal" & _
            " From " & sMovDet & " Md" & _
            " LEFT JOIN ARTICULO A ON A.CODEMP = '" & sCodEmp & "' AND MD.CODSUBALM = A.CODSUBALM AND MD.CODART=A.CODART " & _
            " LEFT JOIN ARTDET AD ON AD.CODEMP = A.CODEMP AND A.CODSUBALM = AD.CODSUBALM AND A.XTIPALM = AD.XTIPALM AND A.CODART = AD.CODART " & _
            " LEFT JOIN XSUBARTDET X03 ON AD.XTIPALM = X03.XTIPALM AND AD.XSUBART03 = X03.CODSUB AND X03.XSUBART = 'XSUBART03' " & _
            " Where Md.CodDoc = '" & sCodDoc & "' And Md.NumDoc = '" & sNumDoc & "' and Md.CodArt!='DET001' AND MD.INGSAL!='N' order by Md.numite"
            
    Set RsDetLar = ICnx.Execute("Select NumDetLar,DetLar Comentario From MovDetLar " & _
                                " Where CodEmp = '" & sCodEmp & "' And CodDoc = '" & sCodDoc & "' And NumDoc='" & sNumDoc & "'  ", adCmdText)

    RsDetLar.Filter = "NumDetLar = '1'"

    If Not RsDetLar.EOF Then sComentario = Trim$("" & RsDetLar.Fields("Comentario").Value)

    Set RsCab = New ADODB.Recordset: Set RsCab = ICnx.Execute(sSqlCab, , adCmdText)
    Set RsDet = New ADODB.Recordset: Set RsDet = ICnx.Execute(sSqlDet, , adCmdText)
    If RsCab.EOF Then Exit Sub
    
    If TypeName(Obj) = "PictureBox" Then
        Obj.Width = 12000: Obj.Height = 15500 '9500
    Else
        Obj.PaperSize = 9
    End If
    
    sRow = 85: Call Encabezado(Obj, RsCab, sLeft, lTop)
    
    RsCab.MoveFirst
    RsDet.MoveFirst

    Obj.Font.Size = 8
    Dim Pag As Long, P As Long
    P = 0: sRowY = 93
    Do Until RsDet.EOF
        Call Cortar(Trim$("" & RsDet.Fields("DesArt").Value), 90, False, False, sLn(), lLn, True, 0)
        sAumRow = 0
        For lCnt = 1 To lLn
            If lCnt > 1 Then sAumRow = sAumRow + 4
        Next lCnt
        sRowY = sRowY + sAumRow + 4
    RsDet.MoveNext
    Loop
    RsDet.MoveFirst
    
    'david 180
    If sRowY > 184 Then Pag = (sRowY / 225) + 1 Else Pag = sRowY / 184
    If Pag = 0 Then Pag = 1 Else Pag = Pag
    ''=============================================================================
    ''=======================DETALLE DEL DOCUMENTO=================================
    ''=============================================================================
    Dim sPosXTot(5) As Single
    sPosXTot(0) = 15: sPosXTot(1) = 42: sPosXTot(2) = 125: sPosXTot(3) = 140: sPosXTot(4) = 155: sPosXTot(5) = 175
'    sRow = 174
    Dim CYC As Long, CXC As Long
    CXC = 15
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 42 - 19, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 21.5 - Printer.TextWidth("CÓDIGO"): Obj.CurrentY = Cy + sRow: Obj.Print "CÓDIGO"
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 42 - 19, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 42 - 19, Cy + 181), , B
    CXC = CXC + 42 - 19
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 68 + 19, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 48 - Printer.TextWidth("DESCRIPCIÓN"): Obj.CurrentY = Cy + sRow: Obj.Print "DESCRIPCIÓN"
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 68 + 19, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 68 + 19, Cy + 181), , B
    CXC = CXC + 68 + 19
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 15, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 12 - Printer.TextWidth("CANT."): Obj.CurrentY = Cy + sRow: Obj.Print "CANT."
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 15, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 15, Cy + 181), , B
    CXC = CXC + 15
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 12, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 10 - Printer.TextWidth("UNID."): Obj.CurrentY = Cy + sRow: Obj.Print "UNID."
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 181), , B
    CXC = CXC + 12
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 19, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 18.5 - Printer.TextWidth("PRECIO UNIT."): Obj.CurrentY = Cy + sRow: Obj.Print "PRECIO UNIT."
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 19, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 19, Cy + 181), , B
    CXC = CXC + 19
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 12, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 12 - Printer.TextWidth("DSCTO.%"): Obj.CurrentY = Cy + sRow: Obj.Print "DSCTO.%"
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 181), , B
    CXC = CXC + 12
    Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 27, Cy + sRow + 5), , B
    Obj.CurrentX = Cx + CXC + 22.5 - Printer.TextWidth("VALOR VENTA"): Obj.CurrentY = Cy + sRow: Obj.Print "VALOR VENTA"
    If sRowY >= 184 Then Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 27, Cy + 225), , B Else Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 27, Cy + 181), , B
    ''=============================================================================
    ''=============================================================================
    sRow = sRow + 8
        P = 1
        Do Until RsDet.EOF
        Obj.Font.Size = 8: Obj.Font.Bold = False
            ''CANTIDAD DE ARTICULOS
            CXC = 15: sAumRow = 0
            Obj.CurrentX = Cx + CXC + 2: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!CODART)
            ''DESCRIPCION DE ARTICULOS
            Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
            Call Cortar(Trim$("" & RsDet.Fields("DesArt").Value), 90, False, False, sLn(), lLn, True, 0)
            For lCnt = 1 To lLn
                If lCnt > 1 Then sAumRow = sAumRow + 4
                Obj.CurrentX = Cx + 59 - 20: Obj.CurrentY = Cy + sRow + sAumRow: Obj.Print sLn(lCnt)
            Next lCnt

            Obj.CurrentX = Cx + 134 - Printer.TextWidth(RsDet.Fields("CanTot").Value): Obj.CurrentY = Cy + sRow: Obj.Print RsDet.Fields("CanTot").Value
            
            Obj.CurrentX = Cx + 143: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!xTipUni)

            ''PRECIO UNITARIO
            Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
            'Obj.CurrentX = Cx + 153: Obj.CurrentY = Cy + sRow: Obj.Print rsCab!Sig
            Obj.CurrentX = Cx + 168.5 - Printer.TextWidth(Format(RsDet.Fields("PreVta").Value, "#,0.00")): Obj.CurrentY = Cy + sRow
            Obj.Print Format(RsDet.Fields("PreVta").Value, "#,0.00")
            
            Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
            Obj.CurrentX = Cx + 176.5: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!DESCUENTO)
            
            ''VALOR VENTA
            Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
            'Obj.CurrentX = Cx + 185: Obj.CurrentY = Cy + sRow: Obj.Print rsCab!Sig
            Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsDet.Fields("TOTALSINIGV").Value, "#,0.00")): Obj.CurrentY = Cy + sRow
            Obj.Print Format(RsDet.Fields("TOTALSINIGV").Value, "#,0.00")
            Obj.Font.Size = 10: Printer.Font.Size = 10
            sRow = sRow + sAumRow + 4
            
            '======ACUMULADOR======
'            If Trim$("" & RsDet!V01) = "G" Then
'            GRAV = GRAV + Format(RsDet.Fields("TOTALSINIGV").Value, "#,0.00")
'            ElseIf Trim$("" & RsDet!V01) = "E" Then
'            EXO = EXO + Format(RsDet.Fields("TOTALSINIGV").Value, "#,0.00")
'            ElseIf Trim$("" & RsDet!V01) = "T" Then
'            GRAT = GRAT + Format(RsDet.Fields("TOTALSINIGV").Value, "#,0.00")
'            End If
            If Trim$("" & RsDet!V01) = "G" Then
            GRAV = GRAV + Format(RsDet.Fields("PreVta").Value * RsDet.Fields("CanTot").Value, "#,0.00")
            ElseIf Trim$("" & RsDet!V01) = "E" Then
            EXO = EXO + Format(RsDet.Fields("PreVta").Value * RsDet.Fields("CanTot").Value, "#,0.00")
            ElseIf Trim$("" & RsDet!V01) = "T" Then
            GRAT = GRAT + Format(RsDet.Fields("PreVta").Value * RsDet.Fields("CanTot").Value, "#,0.00")
            End If
            'DETRA = DETRA + Format(RsDet!TOTART * RsDet!DETRACCION, "#,0.00")
            '======================
            
    If (sRow >= 223) Or (sRow >= sRowY And sRowY >= 184) Then
        Obj.CurrentX = Cx + 178: Obj.CurrentY = Cy + 268.5: Obj.Print "Pagina " & P & " de " & Pag
        P = P + 1
        If TypeName(Obj) = "PictureBox" Then pvWindow.NewPage: Set Obj = pvWindow.ActivePage Else Obj.NewPage
        Call Encabezado(Obj, RsCab, sLeft, lTop)
        sRow = 85: CXC = 15: Obj.Font.Bold = True: Obj.Font.Size = 8
        
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 42 - 19, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 21.5 - Printer.TextWidth("CÓDIGO"): Obj.CurrentY = Cy + sRow: Obj.Print "CÓDIGO"
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 42 - 19, Cy + 181), , B
        CXC = CXC + 42 - 19
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 68 + 19, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 51 - Printer.TextWidth("DESCRIPCIÓN"): Obj.CurrentY = Cy + sRow: Obj.Print "DESCRIPCIÓN"
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 68 + 19, Cy + 181), , B
        CXC = CXC + 68 + 19
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 15, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 13 - Printer.TextWidth("CANT."): Obj.CurrentY = Cy + sRow: Obj.Print "CANT."
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 15, Cy + 181), , B
        CXC = CXC + 15
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 12, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 11 - Printer.TextWidth("UNID."): Obj.CurrentY = Cy + sRow: Obj.Print "UNID."
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 181), , B
        CXC = CXC + 12
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 19, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 18.5 - Printer.TextWidth("PRECIO UNIT."): Obj.CurrentY = Cy + sRow: Obj.Print "PRECIO UNIT."
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 19, Cy + 181), , B
        CXC = CXC + 19
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 12, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 12 - Printer.TextWidth("DSCTO.%"): Obj.CurrentY = Cy + sRow: Obj.Print "DSCTO.%"
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 12, Cy + 181), , B
        CXC = CXC + 12
        Obj.Line (Cx + CXC, Cy + sRow - 1)-(Cx + CXC + 27, Cy + sRow + 5), , B
        Obj.CurrentX = Cx + CXC + 22.5 - Printer.TextWidth("VALOR VENTA"): Obj.CurrentY = Cy + sRow: Obj.Print "VALOR VENTA"
        Obj.Line (Cx + CXC, Cy + sRow + 5)-(Cx + CXC + 27, Cy + 181), , B
        sRow = sRow + 8
    End If
    RsDet.MoveNext
    Loop

    ''=============================================================================
    ''=======================TOTAL EN LETRAS=======================================
    ''=============================================================================
    sRow = 182: sAumRow = 0
    Call Cortar(Mid(UCase(Trim$("" & RsCab.Fields("TotDocL").Value)), 1, Len(Trim$("" & RsCab.Fields("TotDocL").Value))), 150, False, False, sLn(), lLn, True, 0)
    For lCnt = 1 To lLn
    If lCnt > 1 Then sAumRow = sAumRow + 4
        Obj.Font.Size = 9
        Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow + sAumRow: Obj.Print sLn(lCnt)
    Next lCnt
    
    Obj.Line (Cx + 15, Cy + sRow - 1)-(Cx + 210, Cy + sRow + 4), , B
    ''=============================================================================
    ''=======================OBSERVACIONES=========================================
    ''=============================================================================
    sRow = 190: sAumRow = 0: Obj.Font.Bold = True: sRowY = sRow
    Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow: Obj.Print "CONSIDERACIONES :"
    Obj.Font.Bold = False: Printer.Font.Bold = False: Obj.Font.Size = 7

    sComentario = UCase(Trim$("" & RsCab.Fields("COMMEM").Value))
'    If sComentario = "" Then
'    sComentario = "MOTIVOS DE LA EMISION DE LA NOTA DE CREDITO ELECTRONICA" & _
'    vbCrLf & "01 ANULACION DE OPERACION " & _
'    vbCrLf & "02 ANULACION POR ERROR EN EL RUC" & _
'    vbCrLf & "03 CORRECION POR ERROR EN LA DESCRIPCION" & _
'    vbCrLf & "04 DESCUENTO GLOBAL" & _
'    vbCrLf & "05 DESCUENTO POR ITEM" & _
'    vbCrLf & "06 DEVOLUCION TOTAL" & _
'    vbCrLf & "07 DEVOLUCION POR ITEM" & _
'    vbCrLf & "08 BONIFICACION" & _
'    vbCrLf & "09 DISMINUCION EN EL VALOR"
'    End If
'
    Call CortarX(sComentario, sLn(), lLn, sLnx(), lLnx)
         For X = 1 To lLnx
             If (Len(sLnx(X)) > 115) Then
                 Call Cortar(sLnx(X), 115, True, False, sLn(), lLn, False)
                      For Y = 1 To lLn
                          Obj.CurrentX = Cx + 20: Obj.CurrentY = sRow + sAumRow: Obj.Print sLn(Y)
                          sAumRow = sAumRow + 4
                      Next
              Else
                  Obj.CurrentX = Cx + 20: Obj.CurrentY = sRow + sAumRow: Obj.Print sLnx(X)
                  sAumRow = sAumRow + 4
              End If
         Next

    ''=============================================================================
    ''=======================MONTOS TOTALES DEL DOCUMENTO==========================
    ''=============================================================================
    Obj.Font.Size = 8
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Op. Gratuitas"): Obj.CurrentY = Cy + sRow: Obj.Print "Op. Gratuitas"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(GRAT, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(GRAT, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Op. Exoneradas"): Obj.CurrentY = Cy + sRow: Obj.Print "Op. Exoneradas"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(EXO, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(EXO, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Op. Gravadas"): Obj.CurrentY = Cy + sRow: Obj.Print "Op. Gravadas"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(GRAV, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(GRAV, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Total Descuento"): Obj.CurrentY = Cy + sRow: Obj.Print "Total Descuento"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsCab!TotDsc1 + RsCab!TotDsc2 + RsCab!TotDsc3 + RsCab!TotDsc4, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(RsCab!TotDsc1 + RsCab!TotDsc2 + RsCab!TotDsc3 + RsCab!TotDsc4, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Sub-total"): Obj.CurrentY = Cy + sRow: Obj.Print "Sub-total"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsCab.Fields("SubTot").Value, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(RsCab.Fields("SubTot").Value, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("I.G.V."): Obj.CurrentY = Cy + sRow: Obj.Print "I.G.V."
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsCab.Fields("TotImp").Value, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(RsCab.Fields("TotImp").Value, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Importe Total"): Obj.CurrentY = Cy + sRow: Obj.Print "Importe Total"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsCab.Fields("TotDoc").Value, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(RsCab.Fields("TotDoc").Value, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Detracción"): Obj.CurrentY = Cy + sRow: Obj.Print "Detracción"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(DETRA, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(DETRA, "#,0.00")
    sRow = sRow + 4
    Obj.CurrentX = Cx + 166 - Printer.TextWidth("Importe Neto"): Obj.CurrentY = Cy + sRow: Obj.Print "Importe Neto"
    Obj.CurrentX = Cx + 170: Obj.CurrentY = Cy + sRow: Obj.Print ": " & RsCab!Sig
    Obj.CurrentX = Cx + 207.5 - Printer.TextWidth(Format(RsCab.Fields("TotDoc").Value - DETRA, "#,0.00")): Obj.CurrentY = Cy + sRow: Obj.Print Format(RsCab.Fields("TotDoc").Value - DETRA, "#,0.00")
    
    Obj.Line (Cx + 15, Cy + sRowY - 1)-(Cx + 135, Cy + sRow + 5), , B
    
    Obj.Line (Cx + 140, Cy + sRowY - 1)-(Cx + 210, Cy + sRow + 5), , B
    
    ''=============================================================================
    ''=============================================================================
    sRow = 235
    'Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow: Obj.Print "Agente de Retención R.S. 139-2015/SUNAT - 01/07/2015"
    sRow = sRow + 8: sAumRow = 0
    sComentario = "Representación impresa de la " & LCase(sNomDoc) & ", el comprobante puede ser descargado en: www.upeu.edu.pe"
    Call Cortar(Trim$("" & sComentario), 95, False, False, sLn(), lLn, True, 0)
        For lCnt = 1 To lLn
            If lCnt > 1 Then sAumRow = sAumRow + 4
                Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow + sAumRow: Obj.Print sLn(lCnt)
        Next lCnt
    sRow = sRow + sAumRow + 5.5
    Obj.CurrentX = Cx + 20: Obj.CurrentY = Cy + sRow: Obj.Print "Autorizado mediante resolucion SUNAT RI: 0180050002235/SUNAT"
    sRow = sRow + 8

    '============================================================================
    '                   IMPRIMIR CODIGO DE BARRA ADJUNTO
    '============================================================================
Dim xHelp As Object, rsTmp As ADODB.Recordset, lIdObj As Long, sRuta As String, Tool As Object
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = ICnx.Execute("Select FlkObj, IdObj, Ruta From ArcDoc Where CodEmp='" & sCodEmp & "' And CodDoc='" & sCodDoc & "' And NumDoc='" & sNumDoc & "' And xTipAdj='QR';", , adCmdText)
    If Not rsTmp.EOF Then
        If CBool(rsTmp.Fields("FlkObj").Value) Then
            lIdObj = CLng(Val(Trim$("" & rsTmp.Fields("IdObj").Value)))
        End If

        If (lIdObj > 0) Then
            Set Tool = CreateObject("vbxTools.xSystemTools")
            sRuta = Tool.TempDirectory
            Set Tool = Nothing
            If Not oSysAdmin.Obj_Leer(ICnx, lIdObj, sRuta) Then
                Exit Sub
            End If
        Else
            'Buscar la ruta guardada en ArcDoc
            sRuta = Trim$("" & rsTmp.Fields("Ruta").Value)
        End If

    '==================IMAGEN=======================
    If sRuta <> "" Then
            'On Error GoTo erlog
            Dim std As StdPicture
            Set std = GetLogo(sRuta)
            If Not std Is Nothing Then
                Obj.PaintPicture std, 174 - 5, Cy + 235, 30, 30
            End If
            'Obj.Line (175, Cy + 235.5)-(203, Cy + 264), , B
            Obj.Line (173 - 5, Cy + 233.5)-(205 - 5, Cy + 266), , B
    End If
    '===============================================
    End If
    Obj.CurrentX = Cx + 178: Obj.CurrentY = Cy + 269.5: Obj.Print "Pagina " & P & " de " & Pag
    rsTmp.Close: Set rsTmp = Nothing
    If sRuta <> "" Then
    If Dir$(sRuta) <> "" Then
        Call Kill(sRuta)
    End If
    End If
    '============================================================================
    '============================================================================
   
    RsCab.Close: Set RsCab = Nothing
    RsDet.Close: Set RsDet = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub Encabezado(ByRef Obj As Object, ByRef RsCab As ADODB.Recordset, sLeft As Single, ByRef lTop As Long)
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long, iPst As Integer
    Dim CYL As Long, CYS As Long, CYC As Long, CYE As Long, sRutLog As String
    Dim CXL As Long, CXS As Long, CXC As Long, CXCD As Long, CXE As Long, Direccion As String
    Dim fechaTexto As String
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
    If Trim$("" & RsEmp!rutlog) <> "" Then
            sRutLog = Trim$("" & RsEmp!rutlog)
            On Error GoTo erlog
            Dim std As StdPicture
            Dim screenwidth, screenheight As Single
            Set std = GetLogo(sRutLog)
                'screenwidth = Screen.Width \ Screen.TwipsPerPixelX
                'screenheight = Screen.Width \ Screen.TwipsPerPixelX
            If Not std Is Nothing Then
               Obj.PaintPicture std, 8.5 - 3, 15, 45, 20 '75 20
            End If
    End If
    ''=============================================================================
    ''=======================DATOS DE LA EMPRESA===================================
    ''=============================================================================
    'Obj.Font.Name = "Microsoft Sans Serif": Printer.Font.Name = "Microsoft Sans Serif"
    'Obj.Font.Size = 7.5: Obj.Font.Bold = False: Printer.Font.Size = 7.5: Printer.Font.Bold = False
    'Obj.CurrentX = Cx + CXE: Obj.CurrentY = Cy + CYE: Obj.Print Trim$("" & RsEmp!NOMEMP)
    'CYE = CYE + 3
    'Direccion = Trim$("Domicilio Fiscal: " & RsEmp!DirEmp)
    'Call Cortar(Direccion, 70, False, False, sLn(), lLn, True, 0)
    'For lCnt = 1 To lLn
    '    If lCnt > 1 Then CYE = CYE + 3
    '    Obj.CurrentX = Cx + CXE: Obj.CurrentY = Cy + CYE: Obj.Print sLn(lCnt)
    'Next lCnt
    ''=============================================================================
    ''=============================================================================
    
    ''=============================================================================
    ''=======================DATOS DEL LOCAL SUCURSAL==============================
    ''=============================================================================
    'Obj.Font.Name = "Microsoft Sans Serif": Printer.Font.Name = "Microsoft Sans Serif"
    'Obj.Font.Size = 7.5: Obj.Font.Bold = False: Printer.Font.Size = 7.5: Printer.Font.Bold = False
    'Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print "Local Comercial :"
    'CYL = CYL + 3
    'Call Cortar(Trim$("Cal. República de Ecuador 557 - Urb. Zona Industrial"), 55, False, False, sLn(), lLn, True, 0)
    'For lCnt = 1 To lLn
    '    If lCnt > 1 Then CYL = CYL + 3
    '    Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print sLn(lCnt)
    'Next lCnt
    'CYL = CYL + 3
    'Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print Trim$("Lima – Lima – Lima – Perú")
    'CYL = CYL + 3
    'Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print Trim$("Central : " & "+(511) 617-0400")
    'CYL = CYL + 3
    'Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print "Sucursal :" & Trim$("" & sNomSuc)
    'CYL = CYL + 3
    'Call Cortar(Trim$("Dirección:" & sDirLoc), 55, False, False, sLn(), lLn, True, 0)
    'For lCnt = 1 To lLn
    '    If lCnt > 1 Then CYL = CYL + 3
    '    Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print sLn(lCnt)
    'Next lCnt
    'CYL = CYL + 3
    'Obj.CurrentX = Cx + CXL: Obj.CurrentY = Cy + CYL: Obj.Print Trim$("" & sUbiLoc)
    
    ''=============================================================================
    ''=============================================================================
    
    Obj.Font.Size = 13.5: Obj.Font.Bold = True: Printer.Font.Size = 13.5: Printer.Font.Bold = True
    
    ''=============================================================================
    ''=======================DATOS DEL CORRELATIVO=================================
    ''=============================================================================
    Obj.Line (Cx + CXS, CYS - 7)-(Cx + 210, CYS + 38), , B
    Obj.CurrentX = Cx + CXS + 4: Obj.CurrentY = Cy + CYS: Obj.Print "R.U.C. N° "
    Obj.CurrentX = Cx + CXS + 30.5: Obj.CurrentY = Cy + CYS: Obj.Print Trim$("" & RsEmp!RUCEMP)
    CYS = CYS + 12
    sNomDoc = "NOTA DE DEBITO ELECTRONICA"
     Obj.CurrentX = Cx + CXS + 42 + 10 - Printer.TextWidth("NOTA DE DEBITO"): Obj.CurrentY = Cy + CYS: Obj.Print ("NOTA DE DEBITO")
    Obj.CurrentX = Cx + CXS + 42 + 6 - Printer.TextWidth("ELECTRONICA"): Obj.CurrentY = Cy + CYS + 5: Obj.Print ("ELECTRONICA")
    
    'Call Cortar(Trim$("" & sNomDoc), 50, False, False, sLn(), lLn, True, 0)
    'For lCnt = 1 To lLn
    '    If lCnt > 1 Then CYS = CYS + 5
    '    If lCnt > 1 Then
    '    Obj.CurrentX = Cx + CXS + 38 + 10 - Printer.TextWidth(sLn(lCnt)): Obj.CurrentY = Cy + CYS: Obj.Print sLn(lCnt)
    '    Else
    '    Obj.CurrentX = Cx + CXS + 38 + 17 - Printer.TextWidth(sLn(lCnt)): Obj.CurrentY = Cy + CYS: Obj.Print sLn(lCnt)
    '    End If
    'Next lCnt
    CYS = CYS + 14
    
    
      Dim longi, inicio, final, formulax As Integer
      longi = 0: inicio = 0: final = 0
                        longi = Printer.TextWidth(Trim$("" & RsCab!SERIE) & " - " & Trim$("" & sNumDoc))
                        inicio = Cx + CXS   'inicio del cuadro
                        
                        final = inicio + 65 'final del cuadro
                        formulax = inicio + (((final - inicio) / 2) - (longi / 2))
    'Obj.CurrentX = inicio: Obj.CurrentY = CYC + 1: Obj.Print "I"
    'Obj.CurrentX = final: Obj.CurrentY = CYC + 1: Obj.Print "I"
    
    
    Obj.CurrentX = formulax: Obj.CurrentY = CYC + 1: Obj.Print Trim$("" & RsCab!SERIE) & " - " & Trim$("" & sNumDoc)
    ''=============================================================================
    ''=============================================================================
    
    Obj.Font.Size = 8: Obj.Font.Bold = False: Printer.Font.Size = 8: Printer.Font.Bold = False
    
    ''=============================================================================
    ''=======================DATOS DE CABECERA=====================================
    ''=============================================================================
    Obj.Line (Cx + CXC, CYC)-(Cx + CXC + 125, CYC + 17), , B
    CYC = CYC + 3.5
    'Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC: Obj.Print Trim$("Señores: " & rsCab.Fields("NomAne").Value)
    Call Cortar(Trim$("Señores: " & RsCab.Fields("NomAne").Value), 120, False, False, sLn(), lLn, True, 0)
    
    For lCnt = 1 To lLn
        If lLn > 1 Then
            If lCnt = 1 Then
                Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8: Obj.Print sLn(lCnt)
            ElseIf lCnt = 2 Then
                Obj.CurrentX = Cx + CXC + 14: Obj.CurrentY = CYC + 1.8: Obj.Print sLn(lCnt)
            End If
        Else
        Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8: Obj.Print sLn(lCnt)
        End If
    Next lCnt
    CYC = CYC + 3.5
    Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8: Obj.Print Trim$("R.U.C.: " & RsCab.Fields("Ruc").Value);
    CYC = CYC + 3.5
    Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8 ': Obj.Print Mid(Trim$("Dirección: " & rsCab.Fields("DirAne1").Value), 1, 78)
    
    Call Cortar(Trim$("Dirección: " & RsCab.Fields("DirAne1").Value), 120, False, False, sLn(), lLn, True, 0)
    
    For lCnt = 1 To lLn
        If lLn > 1 Then
            If lCnt = 1 Then
                Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8: Obj.Print sLn(lCnt)
            ElseIf lCnt = 2 Then
                Obj.CurrentX = Cx + CXC + 14: Obj.CurrentY = CYC + 1.8: Obj.Print sLn(lCnt)
            End If
        Else
        Obj.CurrentX = Cx + CXC + 2.5: Obj.CurrentY = CYC - 1.8: Obj.Print sLn(lCnt)
        End If
    Next lCnt
    
    
    ''=============================================================================
    ''=============================================================================
    Obj.Font.Name = "Calibri": Printer.Font.Name = "Calibri"
    Obj.Font.Size = 8.2: Obj.Font.Bold = True: Printer.Font.Size = 8.2: Printer.Font.Bold = True
    ''=============================================================================
    ''=======================DATOS DE CUADROS======================================
    ''=============================================================================
    CYC = CYC + 8
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 27 + 28, CYC + 5), , B
    Obj.CurrentX = Cx + CXC + 2.5 + 43.5 - CXX - Printer.TextWidth("NÚMERO DE FACTURA"): Obj.CurrentY = CYC: Obj.Print "NÚMERO DE FACTURA"
    CXC = CXC + 28
    'Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 27, CYC + 5), , B
    'Obj.CurrentX = Cx + CXC + 2.5 + 18 - Printer.TextWidth("FECHA"): Obj.CurrentY = CYC: Obj.Print "FECHA"
    CXC = CXC + 28
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 22, CYC + 5), , B
    Obj.CurrentX = Cx + CXC + 2.5 + 23 - CXX - Printer.TextWidth("FECHA EMISION"): Obj.CurrentY = CYC: Obj.Print "FECHA EMISION"
    CXC = CXC + 23
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 27, CYC + 5), , B
    Obj.CurrentX = Cx + CXC + 2.5 + 28.5 - CXX - Printer.TextWidth("CONDICIÓN DE PAGO"): Obj.CurrentY = CYC: Obj.Print "CONDICIÓN DE PAGO"
    CXC = CXC + 28
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 22, CYC + 5), , B
    Obj.CurrentX = Cx + CXC + 2.5 + 20 - CXX - Printer.TextWidth("FECHA VEN."): Obj.CurrentY = CYC: Obj.Print "FECHA VEN."
    CXC = CXC + 23
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 44, CYC + 5), , B
    '26 DAVID
    Obj.CurrentX = Cx + CXC + 2.5 + 29 - CXX - Printer.TextWidth("MOTIVO"): Obj.CurrentY = CYC: Obj.Print "MOTIVO"
    CXC = CXC + 45
    Obj.Line (Cx + CXC, CYC - 1.5)-(Cx + CXC + 20, CYC + 5), , B
    Obj.CurrentX = Cx + CXC + 2.5 + 21 - CXX - Printer.TextWidth("TIPO MONEDA"): Obj.CurrentY = CYC: Obj.Print "TIPO MONEDA"
    
    ''=============================================================================
    ''=======================DETALLE DE CUADROS====================================
    ''=============================================================================
    CYC = CYC + 6.5: CXC = CXCD
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 55, CYC + 7), , B
    
           longi = 0: inicio = 0: final = 0
                        longi = Printer.TextWidth(Trim$("" & RsCab!NumDocGuia))
                        inicio = Cx + CXC   'inicio del cuadro
                        final = inicio + 54 'final del cuadro
                        formulax = inicio + (((final - inicio) / 2) - (longi / 2))
'    Obj.CurrentX = inicio: Obj.CurrentY = CYC + 1: Obj.Print "I"
'    Obj.CurrentX = final: Obj.CurrentY = CYC + 1: Obj.Print "I"
    If Len(RsCab!CodDocGuia) >= 2 Then
       Obj.CurrentX = formulax: Obj.CurrentY = CYC + 1: Obj.Print Trim$("" & RsCab!NumDocGuia)
    End If
   
    CXC = CXC + 28
    
    CXC = CXC + 28
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 22, CYC + 7), , B
    If (RsCab!FecDoc) <> "" Then
    Obj.CurrentX = Cx + CXC + 17 - Printer.TextWidth(RsCab.Fields("FecDoc").Value): Obj.CurrentY = CYC + 1: Obj.Print RsCab.Fields("FecDoc").Value
    End If
    CXC = CXC + 23
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 27, CYC + 7), , B
    If (RsCab!CondPago) <> "" Then
'    Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC + 1: Obj.Print rsCab.Fields("CondPago").Value
    Call Cortar(RsCab.Fields("CondPago").Value, 20, False, False, sLn(), lLn, True, 0)
    For lCnt = 1 To lLn
        If lLn > 1 Then
        Obj.Font.Size = 7.5: Printer.Font.Size = 7.5
            If lCnt = 1 Then
                Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC - 0.5: Obj.Print sLn(lCnt)
            ElseIf lCnt = 2 Then
                Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC + 2.5: Obj.Print sLn(lCnt)
            End If
        Else
        Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC + 1: Obj.Print sLn(lCnt)
        End If
    Next lCnt
    End If
    Obj.Font.Size = 8.2: Printer.Font.Size = 8.2
    CXC = CXC + 28
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 22, CYC + 7), , B
    If (RsCab!Fecven) <> "" Then
    Obj.CurrentX = Cx + CXC + 2.5 + 14 - Printer.TextWidth(RsCab.Fields("FecVen").Value): Obj.CurrentY = CYC + 1: Obj.Print RsCab.Fields("FecVen").Value
    End If
    CXC = CXC + 23
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 44, CYC + 7), , B
    
    Dim desmotivo As String
    desmotivo = Trim$("" & RsCab!xTipMov) & " - " & Trim$("" & RsCab!motivo)
    If Trim(desmotivo) <> "" Then
        Call Cortar(desmotivo, 45, False, False, sLn(), lLn, True, 0)
        For lCnt = 1 To lLn
            If lLn > 1 Then
            Obj.Font.Size = 7.5: Printer.Font.Size = 7.5
                If lCnt = 1 Then
                    Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC - 0.5: Obj.Print sLn(lCnt)
                ElseIf lCnt = 2 Then
                    Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC + 2.5: Obj.Print sLn(lCnt)
                End If
            Else
            Obj.CurrentX = Cx + CXC + 1.5: Obj.CurrentY = CYC + 1: Obj.Print sLn(lCnt)
            End If
        Next lCnt
    End If

    CXC = CXC + 45
    Obj.Line (Cx + CXC, CYC - 1)-(Cx + CXC + 20, CYC + 7), , B
    
    
    If (RsCab!TipMon) <> "" Then
       longi = 0: inicio = 0: final = 0
                        longi = Printer.TextWidth(Trim$("" & RsCab!TipMon))
                        inicio = Cx + CXC   'inicio del cuadro
                        final = inicio + 19 'final del cuadro
                        formulax = inicio + (((final - inicio) / 2) - (longi / 2))
'    Obj.CurrentX = inicio: Obj.CurrentY = CYC + 1: Obj.Print "I"
'    Obj.CurrentX = final: Obj.CurrentY = CYC + 1: Obj.Print "I"
'
    Obj.CurrentX = formulax: Obj.CurrentY = CYC + 1: Obj.Print Trim$("" & RsCab!TipMon)
    'Obj.CurrentX = Cx + CXC + 2.5 + 11 - Printer.TextWidth(RsCab.Fields("TipMon").Value): Obj.CurrentY = CYC + 1: Obj.Print RsCab.Fields("TipMon").Value
    End If
    ''=============================================================================
    ''=============================================================================

erlog:
            Err.Clear
            Set std = Nothing
    
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






