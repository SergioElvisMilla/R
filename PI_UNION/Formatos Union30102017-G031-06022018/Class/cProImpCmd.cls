VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpCmd"
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

Private sRndRegNeg001 As String 'CanTot
Private sRndRegNeg002 As String 'PreVta
Private sRndRegNeg003 As String 'CanTot * PreVta
Private sRndRegNeg004 As String 'Descuento
Private sRndRegNeg005 As String 'Impuestos
Private sRndTot As String 'Impuestos
Private oSysAdmin As Object
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
    Dim sCurDevice As String, bPrint As Boolean, lCnt As Integer, lNroCop As Integer
    Screen.MousePointer = 11: bPrint = True
    lNroCop = Val(sParam1)
    If lNroCop < 1 Then lNroCop = 1
    
    Set rsEmp = ICnx.Execute("SELECT TOP 1 NOMEMP,DESEMP,DIREMP,TELEMP,FAXEMP,RUCEMP,RUtLOG FROM EMPRESAS WHERE CODEMP='" & ICodEmp & "'", , adCmdText)
    
    Set rsCab = ICnx.Execute("SELECT TOP 1 CAB.TIPCAM, CAB.NUMLOC,LOC.DIRLOC,UBL.DESCRIPCION AS UBILOC,CFG.SERRIE AS SERIE,CAB.XTIPMOV,UBIGEO.DESCRIPCION AS XUBIGEO,CAB.CODDOC,CFG.DESCRI AS DESDOC,CAB.NUMDOC, CAB.NUMREA,CAB.FECDOC,CAB.FECVEN,CAB.CODANE,CAB.NOMANE, " & vbCrLf & _
                             "CAB.DIRANE, CAB.TELANE, CAB.IDEANE, CAB.CODRES, CAB.CODVEN,ANE.NOMANE AS NOMVEN,ANE1.UBIZON AS CODUBIZON, CAB.TIPTR1, CAB.NUMRF1, CAB.TIPTR2,CAB.NUMRF2,UBI.NOMBRE AS UBIZON, " & vbCrLf & _
                             "CAB.SUBTOT - (CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) AS IMPNET, CAB.TOTIMP1+CAB.TOTIMP2+CAB.TOTIMP3+CAB.TOTIMP4 AS TOTIMP, CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4 AS TOTDSC, " & vbCrLf & _
                             "CAB.SUBTOT,CAB.TOTIMP1,CAB.TOTIMP2,CAB.TOTIMP3,CAB.TOTIMP4,CAB.TOTDSC1,CAB.TOTDSC2,CAB.TOTDSC3,CAB.TOTDSC4,CAB.TOTDOC, MAE.DESITE AS CONPAG,MAE1.DESITE AS FORPAG,CAB.NUMDIA,CAB.CODMEM,CAB.COMMEM, " & vbCrLf & _
                             "CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS SIMMON, CAB.TOTDOCL, mV05.DesIte Tipo_D, Ven.NomAne NomVen, Cab.RefInt, Tmp.CodDocPad, Tmp.NumDocPad, CfgP.SerRie SerieP, Cab.IdMesa, Me.Descripcion MesaD" & vbCrLf & _
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
                             "Left Join MaeTabDet mV05 On mV05.CodIte = Cab.v05 And mV05.CodTab = 'vCab05' And mV05.CodFil = 'A'" & vbCrLf & _
                             "Left Join Anexo Ven On Ven.CodAne = Cab.CodVen" & vbCrLf & _
                             "Left Join (Select CodDocPad,NumDocPad,CodDoc,NumDoc From MovCtaArt Group By CodDocPad,NumDocPad,CodDoc,NumDoc) Tmp On Tmp.CodDoc = Cab.CodDoc And Tmp.NumDoc = Cab.NumDoc " & vbCrLf & _
                             "Left Join CfgDoc00 CfgP On CfgP.CodEmp = '" & Trim(ICodEmp) & "' And CfgP.CodDoc = Tmp.CodDocPad" & vbCrLf & _
                             "Left Join Mesa Me On Me.CodEmp='" & ICodEmp & "' And Me.IdMesa=Cab.IdMesa" & vbCrLf & _
                             "WHERE Cab.CodDoc = '" & Trim(ICodDoc) & "' And Cab.NUMDOC='" & INumDoc & "'", , adCmdText)
    
    Call Default_Reglas_Negocio
    If Not rsCab.EOF Then Call Reglas_Negocio(Trim$("" & rsCab.Fields("FecDoc").Value))
     
    Set RsDet = ICnx.Execute("SELECT Md.CodSubAlm, Md.NumIte , MD.CodArt, A.CODARTANT1, Md.DesArt, Md.DesArt1, Md.xTipUni, Mae.DesIte As DesUni, " & vbCrLf & _
                             "Md.Dsc001, Md.Dsc002, Md.Dsc003, Md.Dsc004,  Md.PreVta, Md.Imp001," & vbCrLf & _
                             "Case When Md.CanTot != 0 Then Md.CanTot * Md.PreVta/Md.CanTot Else Md.PreVta End AS PreVta1, Md.CanTot , Md.PreVta*Md.CanTot As PreIte, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*(1-((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100)))) As TotDsc, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100))) As PreTot, Md.NumOrd " & vbCrLf & _
                             "FROM " & IMovDet & " MD " & vbCrLf & _
                             "LEFT JOIN ARTICULO A ON A.CODEMP = '" & ICodEmp & "' AND MD.CODSUBALM = A.CODSUBALM AND MD.CODART=A.CODART " & vbCrLf & _
                             "LEFT JOIN ARTDET AD ON AD.CODEMP = A.CODEMP AND A.CODSUBALM = AD.CODSUBALM AND A.XTIPALM = AD.XTIPALM AND A.CODART = AD.CODART " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X17 ON AD.XTIPALM = X17.XTIPALM AND AD.XSUBART17 = X17.CODSUB AND X17.XSUBART = 'XSUBART17' " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X11 ON AD.XTIPALM = X11.XTIPALM AND AD.XSUBART11 = X11.CODSUB AND X11.XSUBART = 'XSUBART11' " & vbCrLf & _
                             "LEFT JOIN MAETABDET MAE ON MAE.CODTAB = 'XTIPUNI' AND MD.XTIPUNI = MAE.CODITE " & vbCrLf & _
                             "WHERE Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NUMDOC='" & INumDoc & "' AND ISNULL(MD.NUMITE1,'')='' AND MD.CODSUBALM <> 'OT'" & vbCrLf & _
                             "ORDER BY Md.NumOrd Asc", , adCmdText)

        
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
            Call PreviewData(pvWindow.ActivePage)
            Screen.MousePointer = 0
            bPrint = pvWindow.ShowPreView()
        End If
        If bPrint Then
            For lCnt = 1 To lNroCop
                Call PreviewData(pvWindow.ActivePrinter)
                pvWindow.ActivePrinter.EndDoc
                Screen.MousePointer = 0
            Next lCnt
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

'Private Sub PrintData(Obj As Object)
'    Dim sCad As String, lPosI As Long, lPosF As Long, lLen As Long, lYPos As Long, dTotal As Double, CurY As Long
'    Dim dY As Integer, sSeries As String, sCurY As Long, sTmp() As String, lTmp As Long, lCnt As Long, dImp001 As Double
'    Dim dX As Double, dyP As Double, xTmp As Long, xTotArt As Double, xCanTot As Double, xPreVta As Double, xSubTot As Double, xTotDsc As Double, xTotImp As Double, dImp01 As Double
'    Dim lPos As Long, sTipCmd As String
'    With Obj
'        dX = 0: dY = 0
'
'        Obj.ScaleMode = vbMillimeters
'        Obj.Font.Name = "Draft 17cpi"
'        Obj.Font.Size = 8
'
'        'Obj.Height = 3200 + ((RsDet.RecordCount * 1) + 2100)
'        Obj.Height = 6000 + (RsDet.RecordCount * 215)
'        Obj.Width = 6400 '4400
'
'        .CurrentX = 50 + dX - (Obj.TextWidth("** PARA MESA **") / 2)
'        .CurrentY = 4 + dY
'        Obj.Print "** PARA MESA **"
'
'        dY = dY + 4
'
'        lPos = InStr(1, Trim("" & rsCab.Fields("CodDoc").Value), "TB", vbTextCompare)
'        If lPos > 0 Then sTipCmd = "Comanda para Bar"
'
'        lPos = InStr(1, Trim("" & rsCab.Fields("CodDoc").Value), "TC", vbTextCompare)
'        If lPos > 0 Then sTipCmd = "Comanda para Cocina"
'
'        .CurrentX = 50 + dX - (Obj.TextWidth(sTipCmd) / 2)
'        .CurrentY = 4 + dY
'        Obj.Print sTipCmd
'
'        dY = dY + 4
'
'        .CurrentX = 50 + dX - (Obj.TextWidth(Trim("" & rsCab.Fields("RefInt").Value)) / 2)
'        .CurrentY = 4 + dY
'        Obj.Print Trim("" & rsCab.Fields("RefInt").Value)
'
'        dY = dY + 8
'
'        .CurrentX = 4 + dX
'        .CurrentY = 4 + dY
'        Obj.Print "Nro. / Comanda"
'
'        .CurrentX = 39 + dX
'        .CurrentY = 3 + dY
'        Obj.Print ":" & "T/ " & Trim$("" & rsCab.Fields("NumDoc").Value)
'
'        .CurrentX = 4 + dX
'        .CurrentY = 8 + dY
'        Obj.Print "Fecha"
'
'        .CurrentX = 20 + dX
'        .CurrentY = 8 + dY
'        Obj.Print ":" & Trim$("" & rsCab.Fields("FecDoc").Value)
'
'        .CurrentX = 4 + dX
'        .CurrentY = 13 + dY
'        Obj.Print "Mozo"
'
'        .CurrentX = 20 + dX
'        .CurrentY = 13 + dY
'        Obj.Print ":" & Trim$("" & rsCab.Fields("NomVen").Value)
'
'
'        .CurrentX = 4 + dX
'        .CurrentY = 18 + dY
'        Obj.Print "Mesa"
'
'        .CurrentX = 20 + dX
'        .CurrentY = 18 + dY
'        Obj.Print ":" & Trim$("" & rsCab.Fields("IdMesa").Value)
'
'        .CurrentX = 2 + dX
'        .CurrentY = 23 + dY
'        Obj.Print "Cant."
'
'        .CurrentX = 19 + dX
'        .CurrentY = 23 + dY
'        Obj.Print "Producto"
'
'
'        .CurrentX = 4 + dX
'        .CurrentY = 28 + dY
'        Obj.Print "---------------------------------------"
'
'        If RsDet.RecordCount > 0 Then RsDet.MoveFirst
'        .CurrentY = 33 + dY
'        CurY = .CurrentY
'
'        Do Until RsDet.EOF
'
'            .CurrentX = 8 + dX
'            .CurrentY = CurY
'
'            .CurrentX = 11 - .TextWidth(Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")) + dX
'            .CurrentY = CurY
'            Obj.Print Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")
'
'            'xCanTot = Format(Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)
'            'xPreVta = Format(Trim("" & RsDet.Fields("PreVta").Value), sRndRegNeg002)
'            'xSubTot = Format(xCanTot * xPreVta, sRndRegNeg003)
'
'            'xTotDsc = Val(Format(xSubTot * (Val(Trim("" & RsDet.Fields("Dsc001").Value)) / 100), sRndRegNeg004))
'            'xSubTot = xSubTot - xTotDsc
'
'            'xTotImp = Val(Format(xSubTot * (Val(Trim("" & RsDet.Fields("Imp001").Value)) / 100), sRndRegNeg005))
'            'xSubTot = xSubTot + xTotImp
'
'            '.CurrentX = 105 - .TextWidth(Format(xSubTot, "#,0.00")) + dX
'            '.CurrentY = CurY
'            'Obj.Print Format(xSubTot, "#,0.00")
'
'            If Trim(Trim("" & RsDet.Fields("DesArt1").Value)) <> "" Then
'                sCad = Trim("" & RsDet.Fields("DesArt").Value) & vbCrLf & _
'                       Trim("" & RsDet.Fields("DesArt1").Value)
'            Else
'                sCad = Trim("" & RsDet.Fields("DesArt").Value)
'            End If
'
'            Call Cortar(sCad, 35, True, False, sTmp(), lTmp, False)
'            For lCnt = 1 To lTmp
'                Obj.CurrentX = 17 + dX
'                Obj.CurrentY = CurY
'                Obj.Print sTmp(lCnt)
'                'Obj.Print Left(Trim("" & RsDet.Fields("Desart1").Value), 23)
'                If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
'            Next lCnt
'
'            'dTotal = dTotal + Format(xSubTot, "#,0.00")
'
'            CurY = CurY + 4
'
'            RsDet.MoveNext
'        Loop
'    End With
'End Sub

Private Sub PreviewData(Obj As Object)
    Dim sCad As String, lPosI As Long, lPosF As Long, lLen As Long, lYPos As Long, dTotal As Double, CurY As Long
    Dim dY As Integer, sSeries As String, sCurY As Long, sTmp() As String, lTmp As Long, lCnt As Long, dImp001 As Double
    Dim dX As Double, dyP As Double, xTmp As Long, xTotArt As Double, xCanTot As Double, xPreVta As Double, xSubTot As Double, xTotDsc As Double, xTotImp As Double, dImp01 As Double
    Dim lPos As Long, sTipCmd As String
    With Obj
        dX = 0: dY = 0
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Draft 17cpi"
        Obj.Font.Size = 8
                
        Obj.Height = 6000 + (RsDet.RecordCount * 215)
        Obj.Width = 4251.968503937
        
        .CurrentX = 37 + dX - (Obj.TextWidth("** PARA MESA **") / 2)
        .CurrentY = 4 + dY
        Obj.Print "** PARA MESA **"

        dY = dY + 4
        
        lPos = InStr(1, Trim("" & rsCab.Fields("CodDoc").Value), "TB", vbTextCompare)
        If lPos > 0 Then sTipCmd = "Comanda para Bar"
        
        lPos = InStr(1, Trim("" & rsCab.Fields("CodDoc").Value), "TC", vbTextCompare)
        If lPos > 0 Then sTipCmd = "Comanda para Cocina"
        
        .CurrentX = 37 + dX - (Obj.TextWidth(sTipCmd) / 2)
        .CurrentY = 4 + dY
        Obj.Print sTipCmd
        
        dY = dY + 4
        
        .CurrentX = 37 + dX - (Obj.TextWidth(Trim("" & rsCab.Fields("RefInt").Value)) / 2)
        .CurrentY = 4 + dY
        Obj.Print Trim("" & rsCab.Fields("RefInt").Value)
        
        dY = dY + 4
        
        .CurrentX = 4 + dX
        .CurrentY = 4 + dY
        Obj.Print "Nro. / Comanda"
        
        .CurrentX = 24 + dX
        .CurrentY = 4 + dY
        Obj.Print ":" & "T/ " & Trim$("" & rsCab.Fields("NumDoc").Value)
              
        
        .CurrentX = 4 + dX
        .CurrentY = 8 + dY
        Obj.Print "Fecha"
        
        .CurrentX = 13 + dX
        .CurrentY = 8 + dY
        Obj.Print ":" & Trim$("" & rsCab.Fields("FecDoc").Value)
        
        .CurrentX = 4 + dX
        .CurrentY = 13 + dY
        Obj.Print "Mozo"
        
        .CurrentX = 13 + dX
        .CurrentY = 13 + dY
        Obj.Print ":" & Trim$("" & rsCab.Fields("NomVen").Value)
        
        
        .CurrentX = 4 + dX
        .CurrentY = 18 + dY
        Obj.Print "Mesa"
        
        .CurrentX = 13 + dX
        .CurrentY = 18 + dY
        Obj.Print ":" & Trim$("" & rsCab.Fields("IdMesa").Value)
       
        .CurrentX = 4 + dX
        .CurrentY = 23 + dY
        Obj.Print "Cant."
        
        .CurrentX = 17 + dX
        .CurrentY = 23 + dY
        Obj.Print "Producto"
        
        
        .CurrentX = 4 + dX
        .CurrentY = 28 + dY
        Obj.Print "--------------------------------------------------------------"
        
        If RsDet.RecordCount > 0 Then RsDet.MoveFirst
        .CurrentY = 33 + dY
        CurY = .CurrentY
        Do Until RsDet.EOF
            'CurY = .CurrentY
            .CurrentX = 8 + dX
            .CurrentY = CurY

            .CurrentX = 11 - .TextWidth(Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")) + dX
            .CurrentY = CurY
            Obj.Print Format("" & Trim("" & RsDet.Fields("CanTot").Value), "#,0.00")
            
            'xCanTot = Format(Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)
            'xPreVta = Format(Trim("" & RsDet.Fields("PreVta").Value), sRndRegNeg002)
            'xSubTot = Format(xCanTot * xPreVta, sRndRegNeg003)
            
            'xTotDsc = Val(Format(xSubTot * (Val(Trim("" & RsDet.Fields("Dsc001").Value)) / 100), sRndRegNeg004))
            'xSubTot = xSubTot - xTotDsc
            
            'xTotImp = Val(Format(xSubTot * (Val(Trim("" & RsDet.Fields("Imp001").Value)) / 100), sRndRegNeg005))
            'xSubTot = xSubTot + xTotImp

            '.CurrentX = 105 - .TextWidth(Format(xSubTot, "#,0.00")) + dX
            '.CurrentY = CurY
            'Obj.Print Format(xSubTot, "#,0.00")
            If Trim(Trim("" & RsDet.Fields("DesArt1").Value)) <> "" Then
                sCad = Trim("" & RsDet.Fields("DesArt").Value) & vbCrLf & _
                       Trim("" & RsDet.Fields("DesArt1").Value)
            Else
                sCad = Trim("" & RsDet.Fields("DesArt").Value)
            End If
            Call Cortar(sCad, 35, True, False, sTmp(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 17 + dX
                Obj.CurrentY = CurY
                Obj.Print sTmp(lCnt)
                'Obj.Print Left(Trim("" & RsDet.Fields("Desart1").Value), 23)
                If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt

            'dTotal = dTotal + Format(xSubTot, "#,0.00")

            CurY = CurY + 4

            RsDet.MoveNext
        Loop
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










