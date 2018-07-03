VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpPreCtaSum"
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
    Screen.MousePointer = 11: bPrint = True
    
    Set oRsC = New ADODB.Recordset
    
    Set rsEmp = ICnx.Execute("SELECT TOP 1 NOMEMP,DESEMP,DIREMP,TELEMP,FAXEMP,RUCEMP,RUtLOG FROM EMPRESAS WHERE CODEMP='" & ICodEmp & "'", , adCmdText)
    
    Set rsCab = ICnx.Execute("SELECT TOP 1 CAB.TIPCAM, CAB.XTIPMON, CAB.NUMLOC,LOC.DIRLOC,UBL.DESCRIPCION AS UBILOC,CFG.SERRIE AS SERIE,CAB.XTIPMOV,UBIGEO.DESCRIPCION AS XUBIGEO,CAB.CODDOC,CFG.DESCRI AS DESDOC,CAB.NUMDOC, CAB.NUMREA,CAB.FECDOC,CAB.FECVEN,CAB.CODANE,CAB.NOMANE, " & vbCrLf & _
                             "CAB.DIRANE, CAB.TELANE, CAB.IDEANE, CAB.CODRES, CAB.CODVEN,ANE.NOMANE AS NOMVEN,ANE1.UBIZON AS CODUBIZON, CAB.TIPTR1, CAB.NUMRF1, CAB.TIPTR2,CAB.NUMRF2,UBI.NOMBRE AS UBIZON, " & vbCrLf & _
                             "CAB.SUBTOT - (CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) AS IMPNET, CAB.TOTIMP1+CAB.TOTIMP2+CAB.TOTIMP3+CAB.TOTIMP4 AS TOTIMP, CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4 AS TOTDSC, " & vbCrLf & _
                             "CAB.SUBTOT,CAB.TOTIMP1,CAB.TOTIMP2,CAB.TOTIMP3,CAB.TOTIMP4,CAB.TOTDSC1,CAB.TOTDSC2,CAB.TOTDSC3,CAB.TOTDSC4,CAB.TOTDOC, MAE.DESITE AS CONPAG,MAE1.DESITE AS FORPAG,CAB.NUMDIA,CAB.CODMEM,CAB.COMMEM, " & vbCrLf & _
                             "CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS SIMMON, CAB.TOTDOCL, mV05.DesIte Tipo_D, Ven.NomAne NomVen, Cab.xEstDoc, Cab.FecEve, Cab.IdMesa, Me.Descripcion MesaD" & vbCrLf & _
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
                             "Left Join Mesa Me On Me.CodEmp='" & ICodEmp & "' And Me.IdMesa =Cab.IdMesa" & vbCrLf & _
                             "WHERE Cab.CodDoc = '" & Trim(ICodDoc) & "' And Cab.NUMDOC='" & INumDoc & "'", , adCmdText)
    
    Call Default_Reglas_Negocio
    If Not rsCab.EOF Then Call Reglas_Negocio(Trim$("" & rsCab.Fields("FecDoc").Value))
     
    Set RsDet = ICnx.Execute("SELECT Md.CodSubAlm, Md.NumIte , MD.CodArt, A.CODARTANT1, Md.DesArt, Md.DesArt1, Md.xTipUni, Mae.DesIte As DesUni, " & vbCrLf & _
                             "Md.Dsc001, Md.Dsc002, Md.Dsc003, Md.Dsc004,  Md.PreVta,Md.Imp002 , Md.Imp001," & vbCrLf & _
                             "Case When Md.CanTot != 0 Then Md.CanTot * Md.PreVta/Md.CanTot Else Md.PreVta End AS PreVta1, Md.CanTot , Md.PreVta*Md.CanTot As PreIte, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*(1-((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100)))) As TotDsc, " & vbCrLf & _
                             "Md.PreVta*Md.CanTot*((1-(Dsc001/100))*(1-(Dsc002/100))*(1-(Dsc003/100))*(1-(Dsc004/100))) As PreTot, Md.TotArt, Md.CanRea " & vbCrLf & _
                             "FROM " & IMovDet & " MD " & vbCrLf & _
                             "LEFT JOIN ARTICULO A ON A.CODEMP = '" & ICodEmp & "' AND MD.CODSUBALM = A.CODSUBALM AND MD.CODART=A.CODART " & vbCrLf & _
                             "LEFT JOIN ARTDET AD ON AD.CODEMP = A.CODEMP AND A.CODSUBALM = AD.CODSUBALM AND A.XTIPALM = AD.XTIPALM AND A.CODART = AD.CODART " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X17 ON AD.XTIPALM = X17.XTIPALM AND AD.XSUBART17 = X17.CODSUB AND X17.XSUBART = 'XSUBART17' " & vbCrLf & _
                             "LEFT JOIN XSUBARTDET X11 ON AD.XTIPALM = X11.XTIPALM AND AD.XSUBART11 = X11.CODSUB AND X11.XSUBART = 'XSUBART11' " & vbCrLf & _
                             "LEFT JOIN MAETABDET MAE ON MAE.CODTAB = 'XTIPUNI' AND MD.XTIPUNI = MAE.CODITE " & vbCrLf & _
                             "WHERE Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NUMDOC='" & INumDoc & "' AND ISNULL(MD.NUMITE1,'')='' AND MD.CODSUBALM <> 'OT'" & vbCrLf & _
                             "ORDER BY Md.CanRea Asc", , adCmdText)

        
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
                                
    Set oRsDoc = ICnx.Execute("Select Md.CodDoc, Md.NumDoc, Md.NumIte, M01.CodDocPad CodDoc01, M01.NumDocPad NumDoc01, M02.CodDocPad CodDoc02, M02.NumDocPad NumDoc02, M03.CodDocPad CodDoc03, M03.NumDocPad NumDoc03" & vbCrLf & _
                              "From " & Trim(IMovDet) & " Md" & vbCrLf & _
                              "Inner Join MovCtaArt M01 On M01.CodDoc = Md.CodDoc And M01.NumDoc = Md.NumDoc And M01.NumIteCta = Md.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Inner Join MovCtaArt M02 On M02.CodDoc = M01.CodDocPad And M02.NumDoc = M01.NumDocPad And M02.NumIteCta = M01.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Inner Join MovCtaArt M03 On M03.CodDoc = M02.CodDocPad And M03.NumDoc = M02.NumDocPad And M03.NumIteCta = M02.NumIte And M01.CodEmp = '" & Trim(ICodEmp) & "'" & vbCrLf & _
                              "Where Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NumDoc = '" & Trim(INumDoc) & "'" & vbCrLf, , adCmdText)
                                
                
    If Not rsCab.EOF Then
    
    
        'Call GetDatos("REI;REI2", ICodDoc, INumDoc, dMonEnt, dVueltoMn, dVueltoMe)
        'Call Get_RecordSet_Can(oRsC, "REI;REI2", ICodDoc, INumDoc)
            
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
    Dim sCad As String, lPosI As Long, lPosF As Long, lLen As Long, lYPos As Long, dTotal As Double, CurY As Long
    Dim dY As Integer, sSeries As String, sCurY As Long, sTmp() As String, lTmp As Long, lCnt As Long, dImp001 As Double
    Dim dX As Double, dyP As Double, xTmp As Long, xTotArt As Double, xCanTot As Double, xPreVta As Double, xSubTot As Double, xTotDsc As Double, xTotImp As Double, dImp01 As Double
    Dim bAnulado As Boolean
    Dim dCanRea As Double
    Dim dTotArt As Double
    Dim bSubTot As Boolean
    With Obj
        dX = 0: dY = 0
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Draft 17cpi"
        Obj.Font.Size = 8
        
        Obj.Height = 6000 + (RsDet.RecordCount * 470)
        Obj.Width = 4251.968503937
        
        .CurrentX = 37 - ((.TextWidth(Trim$("" & rsEmp.Fields("NomEmp").Value))) / 2) + dX
        .CurrentY = 3 + dY
        Obj.Print Trim$("" & rsEmp.Fields("NomEmp").Value)
        
        .CurrentX = 37 - ((.TextWidth(Trim$("" & rsEmp.Fields("DesEmp").Value))) / 2) + dX
        .CurrentY = 8 + dY
        Obj.Print Trim$("" & rsEmp.Fields("DesEmp").Value)
        
        CurY = 13 + dY
        
        .CurrentX = 37 - ((.TextWidth(Trim$("" & rsCab.Fields("FecDoc").Value & "   " & Format(Trim$("" & rsCab.Fields("FecEve").Value), "hh:mm:ss")))) / 2) + dX
        .CurrentY = 13 + dY
        Obj.Print Trim$("" & rsCab.Fields("FecDoc").Value) & "   " & Format(Trim$("" & rsCab.Fields("FecEve").Value), "hh:mm:ss")
        
        dY = dY + 3
        
        .CurrentX = 4 + dX
        .CurrentY = 18 + dY
        Obj.Print "Mesa : " & Trim$("" & rsCab.Fields("IdMesa").Value)
        
        .CurrentX = 20 + dX
        .CurrentY = 18 + dY
        Obj.Print "Mozo : " & Mid(Trim$("" & rsCab.Fields("NomVen").Value), 1, 22)
        
        .CurrentX = 37 + dX - (Obj.TextWidth("*** PRECUENTA ***") / 2)
        .CurrentY = 26 + dY
        Obj.Print "*** PRECUENTA ***"
        
        .CurrentX = 4 + dX
        .CurrentY = 34 + dY
        Obj.Print "Cant."
        
        .CurrentX = 12 + dX
        .CurrentY = 34 + dY
        Obj.Print "Producto"
        
        .CurrentX = 58 + dX
        .CurrentY = 34 + dY
        Obj.Print "SubTotal"
        
        .CurrentX = 4 + dX
        .CurrentY = 38 + dY
        Obj.Print "--------------------------------------------------------------"
        
                
        If RsDet.RecordCount > 0 Then RsDet.MoveFirst
        .CurrentY = 43 + dY
        CurY = .CurrentY
        Do Until RsDet.EOF
            
            If dCanRea < 0 Then dCanRea = RsDet.Fields("CanRea").Value
            
            If RsDet.Fields("CanRea").Value <> dCanRea Then
            
                If RsDet.AbsolutePosition > 1 Then
                    
                    .CurrentX = 52 + dX
                    .CurrentY = CurY
                    Obj.Print "----------------"
                    
                    CurY = CurY + 4
                    
                    .CurrentX = 69 - .TextWidth(Format(dTotArt, "#,0.00")) + dX
                    .CurrentY = CurY
                    Obj.Print Format(dTotArt, "#,0.00")
                    
                    CurY = CurY + 4
                    
                    dTotArt = 0
                    
                End If
                
                dCanRea = RsDet.Fields("CanRea").Value
                .CurrentX = 9 - .TextWidth(Format("" & Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)) + dX
                .CurrentY = CurY
                Obj.Print "CUENTA PARA LA SILLA " & dCanRea & ":"
                
                CurY = CurY + 4
                
            End If
            
            .CurrentX = 1 + dX
            .CurrentY = CurY

            .CurrentX = 9 - .TextWidth(Format("" & Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)) + dX
            .CurrentY = CurY
            Obj.Print Format("" & Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)
            
            dImp01 = Trim("" & RsDet.Fields("Imp001").Value)
            
            xCanTot = Format(Trim("" & RsDet.Fields("CanTot").Value), sRndRegNeg001)
            xPreVta = Format(Trim("" & RsDet.Fields("PreVta").Value), sRndRegNeg002)
            xSubTot = Format(xCanTot * xPreVta, sRndRegNeg003)
            
            Dim imp2 As Double
            imp2 = Val(Trim("" & RsDet.Fields("Imp002").Value)) + Val(Trim("" & RsDet.Fields("Imp001").Value))
            
            xTotDsc = Val(Format(xSubTot * (Val(Trim("" & RsDet.Fields("Dsc001").Value)) / 100), sRndRegNeg004))
            xSubTot = xSubTot - xTotDsc
            
            xTotImp = Val(Format(xSubTot * (imp2 / 100), sRndRegNeg005))
            xSubTot = xSubTot + xTotImp

            .CurrentX = 69 - .TextWidth(Format(xSubTot, "#,0.00")) + dX
            .CurrentY = CurY
            Obj.Print Format(xSubTot, "#,0.00")
            
            dTotArt = dTotArt + xSubTot

            sCad = Left(Trim("" & RsDet.Fields("DesArt").Value), 23)
            Call Cortar(sCad, 40, True, False, sTmp(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 12 + dX '17
                Obj.CurrentY = CurY
                Obj.Print sTmp(lCnt)
                If lTmp > 1 And lCnt < lTmp Then CurY = CurY + 4
            Next lCnt

            dTotal = dTotal + Format(xSubTot, sRndRegNeg003)

            CurY = CurY + 4

            RsDet.MoveNext
        Loop
        
        .CurrentX = 52 + dX
        .CurrentY = CurY
        Obj.Print "----------------"
        
        CurY = CurY + 4
        
        .CurrentX = 69 - .TextWidth(Format(dTotArt, "#,0.00")) + dX
        .CurrentY = CurY
        Obj.Print Format(dTotArt, "#,0.00")
        
        CurY = CurY + 4
        
        dTotArt = 0
    
        .CurrentX = 4 + dX
        .CurrentY = CurY
        Obj.Print "--------------------------------------------------------------"
        
        CurY = CurY - 4
        
        .CurrentX = 4 + dX
        .CurrentY = CurY + 6 + dY
        Obj.Print "TOTAL:    S/."
        
        .CurrentX = 69 - .TextWidth(Format(Trim("" & rsCab.Fields("TOTDOC").Value), "#,0.00")) + dX
        .CurrentY = CurY + 6 + dY
        Obj.Print "" & Format(Trim("" & rsCab.Fields("TOTDOC").Value), "#,0.00")
        
        .CurrentX = 4 + dX
        .CurrentY = CurY + 15 + dY
        Obj.Print "T.C. S/. " & Trim$("" & rsCab.Fields("TIPCAM").Value)
        
        .CurrentX = 4 + dX
        .CurrentY = CurY + 24 + dY
        Obj.Print "RUC:"
        
        .CurrentX = 13 + dX
        .CurrentY = CurY + 25 + dY
        Obj.Print "-----------------------------------------------------"
        
        
        .CurrentX = 4 + dX
        .CurrentY = CurY + 33 + dY
        Obj.Print "Raz.Soc.:"
        
        .CurrentX = 19 + dX
        .CurrentY = CurY + 34 + dY
        Obj.Print "-----------------------------------------------"
        
        Obj.CurrentX = 37 - ((Obj.TextWidth("Este documento es solo referencial.")) / 2)
        Obj.CurrentY = CurY + 43 + dY
        Obj.Print "Este documento es solo referencial."
        
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
               "Select Mc.MonEnt, Mc.VueltoMn, Mc.VueltoMe, Mc.xTipMov, Mtd.DesIte xTipMov_D" & vbCrLf & _
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





