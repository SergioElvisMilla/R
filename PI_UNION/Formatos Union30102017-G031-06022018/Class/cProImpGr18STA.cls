VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpGr18STA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private sCodDoc As String, sNumDoc As String, sProImp As String
Private sCodEmp As String, sCodGru As String, sTitFrm As String
Private sMovCab As String, sMovDet As String, sDevice As String
Private pai As String, dep As String, pro As String, dis As String, CodUbi As String, referencia As String
Private longi, inicio, final As Long

Private sParam1 As String, INivel As Byte, ICnx As ADODB.Connection
Private Const NFIL_ = 35
Private sNomDoc As String, sCodLoc As String, sNomLoc As String, sNomSuc As String, sDirLoc As String, sUbiLoc As String
Private sTelLoc As String
Private oSysAdmin As Object
Private sComentario As String
Private bPrint As Boolean
Dim RsCab As ADODB.Recordset, RsEmp As ADODB.Recordset, RsDetLar As ADODB.Recordset
Dim Rspai As ADODB.Recordset, rsdep As ADODB.Recordset, rspro As ADODB.Recordset, rsdis As ADODB.Recordset, rsreferencia As ADODB.Recordset

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
    Set RsEmp = ICnx.Execute("Select Top 1 CodEmp,NomEmp,DirEmp,TelEmp,FAXEmp,BKUPEmp,RUCEmp,RUtLOG,EMAIL,CIUDAD,XUBIGEO From EMPRESAS Where CodEmp='" & sCodEmp & "'", , adCmdText)
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
    Dim RsDet As ADODB.Recordset, RsCab As ADODB.Recordset, RsPrecio As ADODB.Recordset, RsSec As ADODB.Recordset, sSqlCab As String, sSqlDet As String, sSqlSec As String, lTop As Long
    Dim Rspai As ADODB.Recordset, rsdep As ADODB.Recordset, rspro As ADODB.Recordset, rsdis As ADODB.Recordset, rsreferencia As ADODB.Recordset
    Dim sLeft As Single, lItem As Integer, sRow As Single, sRowY As Long, sPos(5) As Single, sAumRow As Single, sMaxHeight As Single
    Dim dSubTot As Double, sLmtPag As Single, Col As Long, sLnx() As String, lLnx As Long
    Dim X As Integer, Y As Integer, GRAV As Double, EXO As Double, GRAT As Double, DETRA As Double
    GRAV = 0: EXO = 0: GRAT = 0: DETRA = 0
    
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
    
'------------------------------------
    
    If RsCab.EOF Then Exit Sub
    
    If TypeName(Obj) = "PictureBox" Then
        Obj.Width = 12000: Obj.Height = 15200 '9500
    Else
        Obj.PaperSize = 9
        'Obj.PaperSize = 1
    End If
    
    sRow = 85: Call Encabezado(Obj, RsCab, sLeft, lTop)
    
    RsCab.MoveFirst
    RsDet.MoveFirst

    Obj.Font.Size = 8
    RsDet.MoveFirst
   
    ''=============================================================================
    ''=======================DETALLE DEL DOCUMENTO=================================
    ''=============================================================================
     Obj.Font.Size = 8: Obj.Font.Bold = True
    Dim sPosXTot(5) As Single
    sPosXTot(0) = 15: sPosXTot(1) = 42: sPosXTot(2) = 125: sPosXTot(3) = 140: sPosXTot(4) = 155: sPosXTot(5) = 175
'    sRow = 174
    Dim CYC As Long, CXC As Long
   
    
    ''=============================================================================
    ''=============================================================================
    sRow = sRow + 6
        'P = 1
        
        Dim PRECIOTARI As Double
        Dim CODI As String
    Do Until RsDet.EOF
       Obj.Font.Size = 8: Obj.Font.Bold = False
                       ''CANTIDAD DE ARTICULOS
                        CXC = 15: sAumRow = 0
                        Obj.CurrentX = Cx + 18: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!CODART)
                        
                        ''DESCRIPCION DE ARTICULOS
                        Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
                        
                        Obj.CurrentX = Cx + 58 - Printer.TextWidth(Format(RsDet.Fields("CANTOT").Value, "#,0.00")): Obj.CurrentY = Cy + sRow
                        Obj.Print Format(RsDet.Fields("CANTOT").Value, "#,0.00")
                        
'                        Call Cortar(Trim$("" & RsDet.Fields("DesArt").Value), 87, False, False, sLn(), lLn, True, 0)
'                        For lCnt = 1 To lLn
'                            If lCnt > 1 Then sAumRow = sAumRow + 4
'                            Obj.CurrentX = Cx + 50: Obj.CurrentY = Cy + sRow + sAumRow: Obj.Print sLn(lCnt)
'                        Next lCnt
                        Obj.CurrentX = Cx + 62: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!DESART)

                        
                        Obj.CurrentX = Cx + 155: Obj.CurrentY = Cy + sRow: Obj.Print Trim$("" & RsDet!xTipUni)
            
                        ''PRECIO UNITARIO
                        Obj.Font.Size = 8.5: Printer.Font.Size = 8.5
                        'Obj.CurrentX = Cx + 153: Obj.CurrentY = Cy + sRow: Obj.Print rsCab!Sig
                        CODI = Trim$(RsDet!CODART)
                        Set RsPrecio = ICnx.Execute("Select PREVEN From TARIFDET Where  CODTAR='TAR003' AND CODART='" & CODI & "'", adCmdText)
                        If RsPrecio.RecordCount() = 0 Then
                        PRECIOTARI = 0
                        Else
                        PRECIOTARI = IIf(IsNull(Trim(RsPrecio!PREVEN)), "", Trim(RsPrecio!PREVEN))
                       End If
                        Obj.CurrentX = Cx + 186 - Printer.TextWidth(Format(PRECIOTARI, "#,0.00")): Obj.CurrentY = Cy + sRow
                        Obj.Print Format(PRECIOTARI, "#,0.000")
                        
                
                        

                        sRow = sRow + sAumRow + 3
       
                RsDet.MoveNext
    Loop

Cx = Cx
'
'    CYC = CYC + 263
'    Obj.Font.Size = 10: Obj.Font.Bold = True
'    Obj.CurrentX = Cx + 92: Obj.CurrentY = CYC: Obj.Print "9"
'
'    Obj.Font.Size = 9: Obj.Font.Bold = False
'
''An3.IdeAne1 RucTra, Con.NomCon, Con.TelCon AS Marca, Con.mailcon AS CertifMTC, Con.CelCon AS Placa, Con.DniCon  AS Licencia,Mc.codane02 as CodAneTra,Mc.NomAne02 AS NomAneTra
'    CYC = 243
'    Obj.CurrentX = Cx + 117: Obj.CurrentY = CYC: Obj.Print Trim$("" & UCase(rsCab!NomAneTra))
'
'    CYC = CYC + 4
'    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "RUC"
'    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
'    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!RucTra)
'
'    CYC = CYC + 4
'    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Conductor"
'    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
'    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!NomCon)
'
'
'    CYC = CYC + 4
'    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Licencia N°"
'    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
'    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!Licencia)
'
'
'    CYC = CYC + 4
'    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Vehiculo"
'    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
'    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!Marca)
'    Obj.CurrentX = Cx + 145: Obj.CurrentY = CYC: Obj.Print "Placa :"
'    Obj.CurrentX = Cx + 155: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!Placa)
'
'    CYC = CYC + 4
'    Obj.CurrentX = Cx + 100: Obj.CurrentY = CYC: Obj.Print "Certif. MTC"
'    Obj.CurrentX = Cx + 118: Obj.CurrentY = CYC: Obj.Print ":"
'    Obj.CurrentX = Cx + 120: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!CertifMTC)

    

   
    RsCab.Close: Set RsCab = Nothing
    RsDet.Close: Set RsDet = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub Encabezado(ByRef Obj As Object, ByRef RsCab As ADODB.Recordset, sLeft As Single, ByRef lTop As Long)
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long, iPst As Integer
    Dim CYL As Long, CYS As Long, CYC As Long, CYE As Long, sRutLog As String
    Dim CXL As Long, CXS As Long, CXC As Long, CXCD As Long, CXE As Long, Direccion As String, Direccion2 As String
    Dim fechaTexto As String
    
    '---------------------------------MOVER PARA QUE SE VEA IGUAL VISTA PREVIA Y IMPRESION
    CYL = 25 - 3: CYS = 25: CYC = 47: CYE = 40
    CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    
    If bPrint = True Then                     ' SI IMPRIMES
        CXX = 5
        CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    Else                                       'SINO VISTA PREVIA
        CXX = 5
        CXL = 90: CXS = 144: CXC = 15: CXCD = 15: CXE = 16.5
    End If
    
    ''=============================================================================
    ''=======================LOGO DE LA EMPRESA====================================
    ''=============================================================================
'    If Trim$("" & RsEmp!rutlog) <> "" Then
'            sRutLog = Trim$("" & RsEmp!rutlog)
'            On Error GoTo erlog
'           Dim std As StdPicture
'            Set std = GetLogo(sRutLog)
'            If Not std Is Nothing Then
'                Obj.PaintPicture std, 10, 10, 30, 30
'            End If
'    End If
    
    
    ''=============================================================================
    ''=======================DATOS DE LA EMPRESA===================================
    ''=============================================================================
     inicio = 1 '--DATOS PARA CENTRAR TEXTO
     final = 100 '--DATOS PARA CENTRAR TEXTO
    
    Obj.Font.Size = 11: Obj.Font.Bold = True
    
    CYC = 15
    CXC = 40
    longi = 24 + Len(Trim(RsEmp.Fields("NomEmp").Value))
    Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print Trim(RsEmp.Fields("NomEmp").Value)
    'Obj.CurrentX = 70 + Cx: Obj.CurrentY = CYC: Obj.Print RsEmp!NomEmp

    Obj.Font.Size = 8: Obj.Font.Bold = False
    CYC = CYC + 6
    longi = Len(Trim(RsEmp.Fields("DirEmp").Value))
    Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print Trim(RsEmp.Fields("DirEmp").Value)
    'Obj.CurrentX = 76 + Cx: Obj.CurrentY = CYC: Obj.Print RsEmp!DirEmp

    CodUbi = Trim$("" & RsEmp.Fields("XUBIGEO").Value)
    Set Rspai = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,3)='" & Left(CodUbi, 3) & "'  ", adCmdText)
        If Trim$("" & Rspai.Fields("DESCRIPCION").Value) = "" Then pai = "" Else pai = " -" + Trim$("" & Rspai.Fields("DESCRIPCION").Value)
    Set rsdep = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,5)='" & Left(CodUbi, 5) & "'  ", adCmdText)
        If Trim$("" & rsdep.Fields("DESCRIPCION").Value) = "" Then dep = "" Else dep = Trim$("" & rsdep.Fields("DESCRIPCION").Value)
    Set rspro = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,7)='" & Left(CodUbi, 7) & "'  ", adCmdText)
        If Trim$("" & rspro.Fields("DESCRIPCION").Value) = "" Then pro = "" Else pro = Trim$("" & rspro.Fields("DESCRIPCION").Value)
    Set rsdis = ICnx.Execute("Select DESCRIPCION From UBIGEO Where LEFT(UBIGEO,9)='" & Left(CodUbi, 9) & "'  ", adCmdText)
        If Trim$("" & rsdis.Fields("DESCRIPCION").Value) = "" Then dis = "" Else dis = Trim$("" & rsdis.Fields("DESCRIPCION").Value)

    Dim UBIEMP, TELEMP, COREMP As String
    
    CYC = CYC + 4
    UBIEMP = dis + " - " + pro + " - " + dep
    longi = Len(Trim(UBIEMP))
    Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print UBIEMP
    'Obj.CurrentX = 75.5 + Cx: Obj.CurrentY = CYC: Obj.Print UBIEMP
    
    CYC = CYC + 4
    TELEMP = Trim$("Telf.: " & RsEmp!TELEMP) + " - " + Trim$("Fax: " & RsEmp!FaxEmp)
    longi = Len(Trim(TELEMP))
    Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print TELEMP
    'Obj.CurrentX = 79 + Cx: Obj.CurrentY = CYC: Obj.Print TELEMP
    
    CYC = CYC + 4
    COREMP = Trim(RsEmp!BKUPEMP) + " - " + Trim(RsEmp!EMAIL)
    longi = 10 + Len(Trim(COREMP))
    Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print COREMP

   ' Obj.CurrentX = 79 + Cx: Obj.CurrentY = CYC: Obj.Print COREMP
   
    ''=============================================================================
    ''=======================DATOS DEL CORRELATIVO=================================
    ''=============================================================================
              'x1 empieza, y1 empieza, x2 termina, y2 termina
    Obj.Line (140, 13)-(203, 45), , B
    
    Dim RUCEMP, DOCEMP1, DOCEMP2, NUMDC As String
    
    Obj.Font.Size = 14: Obj.Font.Bold = True
    
    inicio = 145 '--DATOS PARA CENTRAR TEXTO
    final = 200 '--DATOS PARA CENTRAR TEXTO
    CYC = 15
    CXC = 127
   
   RUCEMP = Trim("R.U.C. N° " & Trim$(RsEmp.Fields("RucEmp").Value))
   longi = Len(Trim(RUCEMP))
   Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print RUCEMP

   CYC = CYC + 8
   DOCEMP1 = "GUIA DE REMISION"
   longi = Len(Trim(DOCEMP1)) - 1
   Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print DOCEMP1
   
   CYC = CYC + 6
   DOCEMP2 = "REMITENTE"
   longi = Len(Trim(DOCEMP2)) - 10
   Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print DOCEMP2
   
   CYC = CYC + 8
   NUMDC = Trim$("N° " & RsCab!SERIE) & Trim$(" - " & sNumDoc)
   longi = Len(Trim(NUMDC)) - 8
   Obj.CurrentX = CXC + (((final - inicio) / 2) - (longi / 2)): Obj.CurrentY = CYC: Obj.Print NUMDC


    ''=============================================================================
    ''=======================DATOS DE CABECERA=====================================
    ''=============================================================================
    CYC = 48
    CXC = 0
    Obj.Font.Size = 9: Obj.Font.Bold = False: Printer.Font.Size = 9: Printer.Font.Bold = False
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("FECHA")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab.Fields("FecDoc").Value)
    CYC = CYC + 4.5
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("DNI/RUC")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab.Fields("Ruc").Value);
    CYC = CYC + 4.5
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("CLIENTE")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print Trim$("" & RsCab.Fields("NomAne").Value);
    'Obj.CurrentX = CXC + 160: Obj.CurrentY = CYC: Obj.Print Trim$("" & rsCab!SERIE) & "-" & Trim$("" & sNumDoc)

    CYC = CYC + 4.5
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("DIRECCIÓN")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    '    Direccion2 = Trim$(rsCab.Fields("DirAne").Value) + dis + pro + dep + pai
            Direccion2 = Trim$(RsCab.Fields("DirAne").Value)
        Call Cortar(Direccion2, 90, False, False, sLn(), lLn, True, 0)
        For lCnt = 1 To lLn
            If lCnt > 1 Then CYC = CYC + 4
            Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print sLn(lCnt)
        Next lCnt
    
    CYC = CYC + 4.5
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("P. Partida")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print Trim$("Av Bolognesi 551 Sta Anita")
    
    'Obj.Font.Size = 14: Obj.Font.Bold = True: Printer.Font.Size = 14: Printer.Font.Bold = True
    'Obj.CurrentX = CXC + 150: Obj.CurrentY = CYC: Obj.Print Trim$("GUIA DE CARGA")

    Obj.Font.Size = 9: Obj.Font.Bold = False: Printer.Font.Size = 9: Printer.Font.Bold = False
    CYC = CYC + 4.5
    Obj.CurrentX = CXC + 7: Obj.CurrentY = CYC: Obj.Print Trim$("P. Llegada")
    Obj.CurrentX = CXC + 27: Obj.CurrentY = CYC: Obj.Print Trim$(":")
    Obj.CurrentX = CXC + 29: Obj.CurrentY = CYC: Obj.Print Trim$("Lima - Destinos Varios")
    
    ''=============================================================================
    ''=============================================================================
    
    Obj.Line (5, 80)-(203, 85), , B  ' CUADRO CABECERA
    
    Obj.CurrentX = 9: Obj.CurrentY = 81: Obj.Print Trim$("CODIGO")
    Obj.CurrentX = 30: Obj.CurrentY = 81: Obj.Print Trim$("CANTIDAD")
    Obj.CurrentX = 85: Obj.CurrentY = 81: Obj.Print Trim$("DESCRIPCION")
    Obj.CurrentX = 146: Obj.CurrentY = 81: Obj.Print Trim$("UM")
    Obj.CurrentX = 159: Obj.CurrentY = 81: Obj.Print Trim$("P. UNITARIO")
    Obj.CurrentX = 184: Obj.CurrentY = 81: Obj.Print Trim$("IMPORTE")
    
    Obj.Line (5, 80)-(25, 250), , B  ' CUADRO 1
    Obj.Line (25, 80)-(50, 250), , B  ' CUADRO 2
    Obj.Line (50, 80)-(140, 250), , B  ' CUADRO 3
    Obj.Line (140, 80)-(155, 250), , B  ' CUADRO 4
    Obj.Line (155, 80)-(180, 250), , B  ' CUADRO 5
    Obj.Line (180, 80)-(203, 250), , B  ' CUADRO 5
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













