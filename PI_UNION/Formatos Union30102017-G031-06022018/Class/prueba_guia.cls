VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "prueba_guia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ICnx As ADODB.Connection

Private rsCab As ADODB.Recordset
Private RsDet As ADODB.Recordset
Private RsSer As ADODB.Recordset
Private rsDocRef As ADODB.Recordset
Private rsDocRef2 As ADODB.Recordset
Private rsDirSal As ADODB.Recordset
Private rsDirLle As ADODB.Recordset
Private rsDocRefOP As ADODB.Recordset
Private RsEmp  As ADODB.Recordset


Private rsubidep  As ADODB.Recordset  '--jorge
Private rsubipro  As ADODB.Recordset   '--jorge
Private rsubidis  As ADODB.Recordset   '--jorge


Private departamento, provincia, distrito, codigoubigeo As String  ' --jorge


Private sCodGru As String
Private sCodEmp As String

Private sCodDoc As String
Private sNumDoc As String
Private sMovCab As String
Private sMovDet As String
Private sMovEdi As String
Private sProImp As String
Private sParam1 As String

Private sDevice As String

Private sDocPad1 As String 'Codigo Doc Pedido
Private sNomPAD1 As String 'DOc Pedido
Private sDocPed As String 'DOC Pedido
Private sNumPed As String 'Numero Pedido
Private pvWindow As Object
Private bConAde As Boolean
Private bVisDirecto As Boolean 'Imp sin Vista preliminar
Private bImpReemplazarCodigo As Boolean 'Reemplazar Descripcion Articulos por su Descripcion Especial
Public SCodDocG As String
'ULTIMA MODIFICACION 06/01/2006

Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set ICnx = Cnn
End Property
Public Property Let CodGru(ByVal Cad As String)
    sCodGru = Cad
End Property
Public Property Let CodEmp(ByVal Cad As String)
    sCodEmp = Cad
End Property
Public Property Let CodDoc(ByVal Cad As String)
    sCodDoc = Cad
End Property
Public Property Let NumDoc(ByVal Cad As String)
    sNumDoc = Cad
End Property
Public Property Let MovCab(ByVal Cad As String)
    sMovCab = Cad
End Property
Public Property Let MovDet(ByVal Cad As String)
    sMovDet = Cad
End Property
Public Property Let MovEdi(ByVal Valor As String)
    sMovEdi = Valor
End Property
Public Property Let Param1(ByVal Cad As String)
    sParam1 = Cad
End Property
Public Property Let ProImp(ByVal Valor As String)
    sProImp = Valor
End Property
Public Property Let RutImp(ByVal Cad As String)
    sDevice = Cad
End Property

Public Function PrintShow(Optional PreView As Boolean = False) As Boolean
    PrintDoc True
End Function
Public Function PrintDoc(Optional PreView As Boolean = False) As Boolean
    Dim bPrint As Boolean
    Dim Cadsql As String 'cadena para el WHERE
    Dim sCodDoc As String
    Dim sTmpTab1 As String
    
    'Dim rsDocRef As New ADODB.Recordset
    Dim sWheEdi As String 'condicion de editadas
    
    Dim sCadPed As String 'cadena de Guia para encontrar los Pedidos
    
    bPrint = True
    bVisDirecto = False
    SCodDocG = ""
    Screen.MousePointer = 11
    
    Set rsCab = ICnx.Execute("Select Top 1 CodDoc From CfgDoc00 Where CodEmp='" & sCodEmp & "' And ARCCAB='" & sMovCab & "'", , adCmdText)
    If rsCab.EOF = False Then
        sCodDoc = Trim$(rsCab!CodDoc)
    End If
    rsCab.Close
    Set rsCab = Nothing
    SCodDocG = sCodDoc
'    sWheEdi = " And Md.NumIte1 IS NULL "
    If UCase(Trim$("" & sParam1)) = "EDI" Then
        If Trim("" & sMovEdi) <> "" Then
            sMovDet = sMovEdi
            sWheEdi = " And Md.VISBUE=1" 'solo los articulos que esten chequeados se imprimira
        Else
            MsgBox "Este documento no tiene editadas", vbExclamation, "Error"
            Exit Function
        End If
    End If

    Cadsql = "Select Top 1  M.codane02 as CodAneTra,M.NomAne02 AS NomAneTra,M.ubigeo, M.DirAne02 DirAneTra,an3.DIRANE as dirane2,UB.descripcion,Ub2.Descripcion AS UbigeoTra, An3.IdeAne1 RucTra, Con.NomCon, Con.TelCon AS Marca, Con.CelCon AS Placa, Con.DniCon  AS Licencia, " & _
            " M.CodDoc, M.NumDoc, M.RefInt, M.FecEnt02, M.FecDoc, M.FecEnt, M.NumRf2,M.NumRf1, An2.NomAne AS Vendedor, M.CodVen, M.NumRea, M.FecRea, M.NomAne,M.COMMEM,M.CODMEM,M.DirAne,AN.XUbigeo, M.IdeAne, M.xTipMov, M.MonRf1, M.MonRf2, M1.Desite AS ConPago, Ub.Descripcion AS UbigeoCliente, M.TipTr1, M.NumRf1, M.TipTr2, M.NumRf2 " & _
            " From " & sMovCab & " M " & _
            " Left Join Anexo AN On  M.CodAne = AN.CodAne " & _
            " Left Join Anexo AN2 On  M.CodVen = AN2.CodAne " & _
            " Left Join Anexo AN3 On  M.CodAne02 = AN3.CodAne " & _
            " Left Join Locales LC On  M.CodAne = LC.CodAne And M.NumLoc = LC.NumLoc " & _
            " Left Join Ubigeo UB On CASE IsNull(M.NumLoc,'') WHEN '' THEN AN.xUbigeo ELSE LC.xUbigeo END=UB.Ubigeo " & _
            " Left Join MaeTabDet M1 On  M1.CodTab = 'CONPAG' And M1.CodFil= 'A'       And M.XCONPAG = M1.CodIte " & _
            " Left Join MaeTabDet M3 On  M3.CodTab = 'FORENV' And M.ForEnv = M3.CodIte And M3.CodFil ='A' " & _
            " Left Join Contactos CON On Con.CodAne = An3.CodAne And M.NumCon02 = Con.NumCon " & _
            " Left Join Ubigeo UB2 On An3.xUbigeo = Ub2.Ubigeo " & _
            " Where M.CodDoc='" & sCodDoc & "' And M.NumDoc='" & sNumDoc & "'"
      
    Set rsCab = ICnx.Execute(Cadsql, , adCmdText)
    Cadsql = ""
    
    
    '----------------ubigeo documento---------------------
      If IsNull(rsCab!UBIGEO) = False Then
      codigoubigeo = rsCab!UBIGEO
   
    
      Set rsubidep = ICnx.Execute("Select dep.ubigeo,dep.descripcion" & vbCrLf & _
                             "From ubigeo dep" & vbCrLf & _
                             "Where ubigeo='" & Left((Trim(codigoubigeo)), 5) & "'", , adCmdText)
      Set rsubipro = ICnx.Execute("Select pro.ubigeo,pro.descripcion" & vbCrLf & _
                             "From ubigeo pro" & vbCrLf & _
                             "Where ubigeo='" & Left((Trim(codigoubigeo)), 7) & "'", , adCmdText)
      Set rsubidis = ICnx.Execute("Select dis.ubigeo,dis.descripcion" & vbCrLf & _
                             "From ubigeo dis" & vbCrLf & _
                             "Where ubigeo='" & Left((Trim(codigoubigeo)), 9) & "'", , adCmdText)
                             
        departamento = Trim(rsubidep!descripcion)
        provincia = Trim(rsubipro!descripcion)
        distrito = Trim(rsubidis!descripcion)
        End If
        
    '--------------------------------------------------------
    
    If rsCab.EOF Or rsCab.BOF Then
        Screen.MousePointer = 0
        MsgBox "No se pudo encontrar el documento N° " & sNumDoc, vbExclamation, "Error"
        Exit Function
    End If
    
    'Crear Tabla temporal
    sTmpTab1 = GetTempTable(1)
    
     On Error GoTo deltabla:
    ICnx.Execute "DROP TABLE " & sTmpTab1
    
deltabla:
    On Error Resume Next
    
    Dim wsql11
    
    wsql11 = "Select MIN(Md.NUMORD)NUMORD,Md.CodDoc,Md.NumDoc,Md.CodSubAlm,Md.DesArt1,Md.CodArt,sum(Md.CanTot) CanTot,  CASE  WHEN IsNull(max(X04.NOMSUBART),'')=''  THEN Md.DesArt ELSE max(X04.NOMSUBART) + ' - ' + Md.DesArt END DesArt , MIN(Md.NumIte) AS NumIte,IsNull(Md.NumIte1,'') AS NumIte1,Md.INGSAL,IsNull(Md.TIEGAREmp,0) TIEGAREMP ,A.XTIPUNI, A.DETART" & _
                 " INTO " & sTmpTab1 & " " & _
                 " From " & sMovDet & " MD " & _
                 " Left Join ARTICULO A On A.CodEmp = '" & Trim(sCodEmp) & "' And Md.CodSubAlm = A.CodSubAlm And Md.CodArt=A.CodArt" & _
                 " Left Join ARTDET AD On AD.CodEmp = A.CodEmp And A.CodSubAlm = AD.CodSubAlm And A.CodArt = AD.CodArt" & _
                 " Left Join XSUBARTDET X04 On AD.XTIPALM = X04.XTIPALM And AD.XSUBART04 = X04.CODSUB And X04.XSUBART = 'XSUBART04'" & _
                 " Where Md.CodDoc='" & sCodDoc & "' And Md.NumDoc='" & sNumDoc & "' And Md.CodArt NOT IN('ADE','DSC','CDR')" & _
                 " Group By Md.CodDoc,Md.NumDoc,Md.CodSubAlm,Md.CodArt,Md.DesArt,Md.DesArt1,Md.NumIte1,Md.INGSAL,Md.TIEGAREMP,A.XTIPUNI, A.DETART "
                 
    ICnx.Execute (wsql11)
     
    Set RsDet = ICnx.Execute("Select Md.CodArt,Md.CanTot,  Md.DesArt, Md.DesArt1 , Md.NumIte,IsNull(Md.NumIte1,'') AS NumIte1,Md.INGSAL,IsNull(Md.TIEGAREmp,0) TIEGAREMP,MD.XTIPUNI, md.DETART" & _
                             " From " & sTmpTab1 & " MD" & _
                             " Where Md.CodDoc='" & sCodDoc & "' And Md.NumDoc='" & sNumDoc & "' And Md.CodArt NOT IN('ADE','DSC','CDR') " & sWheEdi & " Order By NUMORD,Md.NumIte", , adCmdText)
    If RsDet.EOF Or RsDet.BOF Then
        Screen.MousePointer = 0
        If UCase(Trim$("" & sParam1)) = "EDI" Then
            MsgBox "No se pudo encontrar el documento N° " & sNumDoc & " de Editadas", vbExclamation, "Error"
            Exit Function
        End If
        MsgBox "No se pudo encontrar el documento N° " & sNumDoc, vbExclamation, "Error"
        Exit Function
    End If
   
    sWheEdi = ""
    Cadsql = "Select  Md.NumIte,Md.NumIte1,Md.CodArt,se.numserext" & _
             " From " & sMovDet & " MD" & _
             " Left Join Series Se On Se.CodSubAlm = Md.CodSubAlm And Se.CodArt = Md.CodArt And Se.CodDocSal = Md.CodDoc And Se.NumDocSal = Md.NumDoc And Se.NumIteSal = Md.NumIte And Se.CodEmp ='" & sCodEmp & "'" & _
             " Where Md.CodDoc='" & sCodDoc & "' And Md.NumDoc='" & sNumDoc & "' And Md.CodArt NOT IN('ADE','DSC','CDR') And Md.INGSAL !='N' "
            
    Set RsSer = ICnx.Execute(Cadsql, , adCmdText)
    Cadsql = ""
    
    Set rsDocRef = ICnx.Execute("Select CodDoc,NumDoc,FecDoc " & _
                                " From movcabe1fc " & _
                                " Where CodDoc = '" & Trim$("" & rsCab!TIPTR2) & "' And NumDoc = '" & Trim$("" & rsCab!NUMRF2) & "'", , adCmdText)
    
    '=============DOC DE PAGO ===================
    Dim wsql1
    wsql1 = "Select NumCor, Max(DirLoc) DirSal, NumLoc " & _
                                " From DOCLOC " & _
                                " Where CodDoc='" & sCodDoc & "' And NumDoc='" & sNumDoc & "' and NumLoc = '' " & _
                                " Group By  NumCor, NumLoc  "
    Set rsDirSal = ICnx.Execute(wsql1)
    
    wsql1 = "Select NumCor, Max(DirLoc) DirLle, NumLoc " & _
                                " From DocLoc " & _
                                " Where CodDoc='" & sCodDoc & "' And NumDoc='" & sNumDoc & "' and NumLoc <> '' " & _
                                " Group By  NumCor, NumLoc  "
    Set rsDirLle = ICnx.Execute(wsql1)
    
    '=========================================================
    '=============DOC PADRE ORDEN PEDIDO (Padre de guia)======
    sNumPed = ""
    Dim wsql111
    wsql111 = " Select  CF.SERRIE,MC.CodDocPad,MC.NumDocPad From MOVCTAART  MC" & _
                                " Left Join CfgDoc00 CF On MC.CodEmp=CF.CodEmp And MC.CodDoc=CF.CodDoc" & _
                                " Where MC.CodEmp='" & sCodEmp & "' And MC.CodDoc ='" & sCodDoc & "'  And MC.NumDoc IN('" & sNumDoc & "') Group By CF.SERRIE,MC.CodDocPad,MC.NumDocPad"
    Set rsDocRef2 = ICnx.Execute(wsql111, , adCmdText)
        
    If Not rsDocRef2.EOF Or Not rsDocRef2.BOF Then
        
        Set rsDocRefOP = ICnx.Execute("select NUMREA from movcabe1OP Where CodDoc ='" & rsDocRef2!CodDocPad & "'  And NumDoc IN('" & rsDocRef2!NumDocPad & "') ", , adCmdText)
        Set RsEmp = ICnx.Execute("SELECT TOP 1 NOMEMP,DIREMP,TELEMP,FAXEMP,RUCEMP,RUtLOG FROM EMPRESAS WHERE CODEMP='" & sCodEmp & "'", , adCmdText)

''        Do While Not rsDocRef.EOF
''            sNumPed = sNumPed & Trim(rsDocRef("SERRIE")) & "-" & Trim(rsDocRef("NumDocPAD")) & ","
''            sDocPed = Trim(rsDocRef("CodDocPAD"))
''            sCadPed = sCadPed & "'" & Trim(rsDocRef("SERRIE")) & Trim(rsDocRef("CodDocPAD")) & Trim(rsDocRef("NumDocPAD")) & "',"
''            rsDocRef.MoveNext
''        Loop
''        sNumPed = Left(sNumPed, Len(sNumPed) - 1)
''        sCadPed = Left(sCadPed, Len(sCadPed) - 1)
    End If
''    rsDocRef.Close: Set rsDocRef = Nothing
    
    '=============== FIN DOC PADRE ORDEN PEDIDO ==============
    '=========================================================
       

    Set pvWindow = CreateObject("PrtTools.cPreView")
    
    Dim prt As Printer
    Set prt = GetPrinter(sDevice)
    
    If Not prt Is Nothing Then
        Set Printer = prt
    End If
    Set prt = Nothing
        
    Set pvWindow.ActivePrinter = Printer
    pvWindow.ScaleMode = vbMillimeters
    
    Screen.MousePointer = 0
'    Dim frm As frmProImpFCN
'    Set frm = New frmProImpFCN
'    If frm.ShowForm = False Then
'        Call Unload(frm): Set frm = Nothing
'        Exit Function
'    Else
'        bImpReemplazarCodigo = frm.ImpReemplazarCodigo
'        Call Unload(frm): Set frm = Nothing
'    End If
    
    If PreView Then
        PrintData pvWindow.ActivePage
        
        Screen.MousePointer = 0
        bPrint = pvWindow.ShowPreView()
    End If
    If bPrint = True Then
        PrintData pvWindow.ActivePrinter
        pvWindow.ActivePrinter.EndDoc
        Screen.MousePointer = 0
    End If
    
    Set pvWindow = Nothing
    
    ICnx.Execute "DROP TABLE " & sTmpTab1
    
    If Not RsDet Is Nothing Then
        RsDet.Close: Set RsDet = Nothing
    End If
    If Not RsSer Is Nothing Then
        RsSer.Close: Set RsSer = Nothing
    End If
    If Not rsDocRef2 Is Nothing Then
        rsDocRef2.Close: Set rsDocRef2 = Nothing
    End If
    If Not rsDocRef Is Nothing Then
        rsDocRef.Close: Set rsDocRef = Nothing
    End If
    If Not rsDirSal Is Nothing Then
        rsDirSal.Close: Set rsDirSal = Nothing
    End If
    If Not rsDirLle Is Nothing Then
        rsDirLle.Close: Set rsDirLle = Nothing
    End If
    If Not RsEmp Is Nothing Then
        RsEmp.Close: Set RsEmp = Nothing
    End If
    If Not rsDocRefOP Is Nothing Then
        rsDocRefOP.Close: Set rsDocRefOP = Nothing
    End If
    rsCab.Close: Set rsCab = Nothing
    PrintDoc = True
    
    Exit Function
Solucion:
    If Err.Number <> 0 Then
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "¿Desea ver mas información?", vbExclamation + vbYesNo, "Error") = vbYes Then
            MsgBox Err.Description, vbCritical, "Error"
            Err.Clear
        End If
    End If
    Screen.MousePointer = 0
End Function
Private Sub PrintData(Obj As Object)
    Dim campoB As String 'cadena a buscar
    Dim campoBB As String
    Dim Pos As Integer 'posición de la cadena a buscar
    Dim poss As Integer
    Dim Valor As String 'PARA BUSCAR CADENA
    Dim CurY As Long
    Dim CuryX As Long
    Dim I As Integer
    Dim II As Integer
    Dim xDerecha As Integer
    Dim xTraslado As Integer
    xDerecha = 5
    
    Dim BK As Variant 'Marcador del item padre de articulo
    Dim Nite As String 'item del articulo padre
    Dim iLonDet As Integer 'Longitud permitido por fila de detalle articulo Padre
    Dim iLonRec As Integer 'Longitud permitido por fila de detalle Receta
    Dim ilonRec1 As Integer
    Dim iLonDet1 As String
    Dim Cad As String
   
    
    Dim sCad As String
    Dim lPosI As Long
    Dim lPosF As Long
    Dim lLen As Long
    Dim sLn() As String, lLn As Long, lCnt As Long
    Dim sTotSinRAdel As Double 'total items que suman
    Dim sTmp() As String, lTmp As Long
    Dim Sw1 As String
    
    Dim CCY1, CCY2, CCY3, CCY4, CCY5, CCY6, CCY7, CCY8, CCY9, ccy10, CCY11, CCY12, CCY13, ccy14, CCY16 As Double
    Dim CFX1, CFX2, CFX3, CFX4, CFX5, CFX6, CFX7, CFX8, CFX9, CFX10, CFX11, CFX12, CFX13, CFX14, CFX15, CFX16, CFX17 As Double
    Dim DCY1, DCY2, DCY3, DCY4, DCY5, DCY6, DCY7, DCY8, DCY9, DCY10 As Double
    Dim DFX1, DFX2, DFX3, DFX4, DFX5, DFX6, DFX7, DFX8, DFX9 As Double
    Dim PCY1, PCY2, PCY3, PCY4, PCY5, PCY6, PCY7, PCY8, PCY9 As Double
    Dim PFX1, PFX2, PFX3, PFX4, PFX5, PFX6, PFX7, PFX8, PFX9 As Double
    
    CCY1 = 28: CCY2 = 32.5: CCY3 = 36.5: CCY4 = 45: CCY5 = 52.5: CCY6 = 56.5: CCY7 = 61: CCY8 = 47.5: CCY9 = 65: ccy10 = 150: CCY11 = 65: CCY12 = 31: ccy14 = 130: CCY16 = 135
    CFX1 = 6: CFX2 = 96: CFX3 = 40: CFX4 = 175: CFX5 = 30: CFX6 = 41: CFX7 = 26: CFX8 = 32: CFX9 = 101: CFX10 = 27: CFX11 = 149: CFX12 = 25: CFX13 = 50: CFX14 = 90: CFX16 = 34: CFX17 = 40
    DCY1 = 70: DCY2 = 0
    DCY10 = 80: DCY2 = 0
    
    DFX1 = 6: DFX2 = 30: DFX3 = 160: DFX4 = 186: DFX5 = 0: DFX6 = 0
    PCY1 = 183: PCY2 = 0: PCY3 = 0
    PFX1 = 168: PFX2 = 200: PFX3 = 0: PFX4 = 0: PFX5 = 0: PFX6 = 0
    
    iLonDet = 60 '60 jorge - longitud de descripcion detalle
    iLonRec = 60
  
    ilonRec1 = 60
    
    Obj.Font.Name = "Draft 17cpi"
    Obj.Height = 12547
    Obj.Width = 13257
    Obj.ScaleMode = 6
   ' Call Busca_SerNumDocumento(sCodEmp, sCodDoc, Trim(rsCab!NumDoc), ICnx)
    
    Dim sw
    sw = 1
    Dim xvalor As String
    On Error GoTo ver3
    On Error GoTo ver4
    Obj.CurrentX = CFX1: Obj.CurrentY = CCY1
  '  Obj.Print " ORDEN PEDIDO : " & nSerNumOP & "  ORDEN COMPRA: " & Trim$("" & rsDocRefOP!NUMREA)
ver4:
    Obj.CurrentX = CFX16 + 10: Obj.CurrentY = CCY12 + 3 ' jorge
    If SCodDocG = "GR" Then
       'Obj.Print Trim$("" & rsEmp!DirEmp)      '"Av. Minerales Nº 310 Lima-Perú"
       Obj.Print "Av. Minerales Nº 310 Cercado de Lima – Lima"
    Else
        If SCodDocG = "GR17" Then
        'rsEmp!DirEmp = "Av. Brasil Mz L Lote 01 –Villa Maria – Nuevo Chimbote"
        Obj.Print "Av. Brasil Mz L Lote 01 –Villa Maria – Nuevo Chimbote-Santa-Ancash"
        Else
            If SCodDocG = "GR06" Then
                'Obj.Print Trim$("" & rsEmp!DirEmp)
                Obj.Print "Av. Los Pescadores C12B Sector Ind  27 de Octubre – Chimbote-Santa-Ancash"
            Else
             If SCodDocG = "GR07" Then
                'Obj.Print Trim$("" & rsEmp!DirEmp)
                Obj.Print "Parcela Rustica 15176 – La encalada  - LAREDO – TRUJILLO –LA LIBERTAD"
             Else
                If SCodDocG = "GR21" Then
                'Obj.Print Trim$("" & rsEmp!DirEmp)
                Obj.Print "Mz Q Lt 2 ASOC UMAPALCA –SABANDIA –AREQUIPA"
               
            Else
                If SCodDocG = "GR22" Then
                'Obj.Print Trim$("" & rsEmp!DirEmp)
                Obj.Print "Av.Pacifico Mza.E Lote 3 Urb.buenos Aires - Nuevo Chimbote - Santa - Ancash"
            Else
                If SCodDocG = "GR23" Then
                'Obj.Print Trim$("" & rsEmp!DirEmp)
                Obj.Print "Nro S/N Predio Rustico La encalada La LIBERTAD - TRUJILLO - VICTOR LARCO HERRERA"
                End If
               End If
             End If
            End If
        End If
    End If
   End If
   
    sw = -1
ver3:
    'If sw <> -1 Then
      '  Obj.Print "Av. Minerales Nº 310 Lima-Perú"
    'End If
    'If rsEmp!DirEmp Then
    'End If
   '' MsgBox "Traslado"
   '' Obj.Print Trim$("" & rsEmp!DirEmp)      '"Av. Minerales Nº 310 Lima-Perú"

    
    'FECHA DE EMISION
    Obj.CurrentX = CFX2: Obj.CurrentY = CCY1
    Obj.Print Format$(rsCab!FecDoc, "dd/MM/YYYY")
    'DIRECCION DE PUNTO DE SALIDA
    If Not rsDirSal.EOF Or Not rsDirSal.BOF Then
        Obj.CurrentX = CFX3: Obj.CurrentY = CCY2
        Obj.Print Trim$("" & rsDirSal!DirSal)
    End If
    'CODIGO DEL DOCUMENTO
    Obj.CurrentX = CFX4: Obj.CurrentY = CCY2
   ' Obj.Print nserNumGr     'Trim$("" & rsCab!NumDoc)
    'DESTINATARIO
    Obj.CurrentX = CFX5 + 2: Obj.CurrentY = CCY3 + 3 'jorge
    Obj.Print Trim$("" & rsCab!NomAne)
    'DIRECCION DE PUNTO DE LLEGADA
    If Not rsDirLle.EOF Or Not rsDirLle.BOF Then
        Obj.CurrentX = CFX6: Obj.CurrentY = CCY4
        Obj.Print Trim$("" & rsDirLle!DirLLe)
    Else
        Obj.CurrentX = CFX6: Obj.CurrentY = CCY4
           If IsNull(rsCab!UBIGEO) = False Then
        Obj.Print Trim$("" & rsCab!DIRANE)
        Obj.CurrentX = CFX6: Obj.CurrentY = CCY4 + 3
        Obj.Print departamento + "-" + provincia + "-" + distrito
         Else
         Obj.Print Trim$("" & rsCab!DIRANE)
      End If

    End If
    
    'NUMERO DE GUIA
    Obj.CurrentX = CFX5 + 138: Obj.CurrentY = CCY3 - 2
    Obj.Print Trim$("Nº GR: " & rsCab!NumDoc)
    'RUC DEL DESTINATARIO
    Obj.CurrentX = CFX7 - 2: Obj.CurrentY = CCY5 + 3 'jorge
    Obj.Print Trim$("" & rsCab!IDEANE)
    'TRANSPORTISTA + xDerecha
    Obj.CurrentX = CFX8: Obj.CurrentY = CCY6 + 3
    Obj.Print Trim$("" & rsCab!NomAneTra)
    'RUC DEL TRANSPORTISTA
    Obj.CurrentX = CFX9: Obj.CurrentY = CCY6 + 3
    Obj.Print Trim$("" & rsCab!RucTra)
    'DIRECCION DEL TRANSPORTISTA
    Obj.CurrentX = CFX10: Obj.CurrentY = CCY7 + 1.5 'jorge
    If sw = 1 Then
       ' Call buscaAnexo(IIf(IsNull(rsCab!codAneTra), "", rsCab!codAneTra), xvalor, ICnx)
    If Len(xvalor) > 0 Then Obj.Print Trim$("" & xvalor)
    Else
    Obj.Print Trim$("" & rsCab!DirAne2) & " - " & Trim$("" & rsCab!UbigeoTra) ' jorge
    End If
    'CHOFER
    Obj.CurrentX = CFX12: Obj.CurrentY = CCY9 + 3 'jorge
    Obj.Print Trim$("" & rsCab!NOMCON)
    Dim CurLinY As Integer
    CurLinY = 0
    lTmp = 0
    If Trim$("" & rsCab!COMMEM) <> "" Then
        Call Cortar(Trim$("" & rsCab.Fields("ComMem").Value), 150, True, False, sTmp(), lTmp, True)
        For lCnt = 1 To lTmp
            Obj.CurrentX = CFX5
            Obj.CurrentY = ccy14 + CurLinY + 10
            Obj.Print sTmp(lCnt)
            CurLinY = CurLinY + 3
        Next
    End If
       
    If Trim$("" & rsCab!CodMem) <> "" Then
        Call Cortar(Trim$("" & rsCab.Fields("CoDMem").Value), 150, True, False, sTmp(), lTmp, True)
        For lCnt = 1 To lTmp
            Obj.CurrentX = CFX5
            Obj.CurrentY = CCY16 + CurLinY + 10
            Obj.Print sTmp(lCnt)
            CurLinY = CurLinY + 3
        Next
    End If
    
    'LICENCIA
    If bImpReemplazarCodigo = True Then
    Obj.CurrentX = CFX13: Obj.CurrentY = ccy10 + 10
    Obj.Print Trim$("BREVETE :" & rsCab!LICENCIA)
    
    'NRO. PLACA / MARCA VEHICULO.
    Obj.CurrentX = CFX14: Obj.CurrentY = CCY11
    Obj.Print Trim$("" & rsCab!PLACA)
    '/ " & Trim$("" & rsCab!marca)
    CurY = 15
    End If
    
    
    
    Select Case UCase(Trim$("" & rsCab!XTIPMOV)) 'Trim$("" & RsCab!XTIPMOV)
        Case "BI" 'VENTA    1
            Obj.CurrentX = CFX11
            Obj.CurrentY = CCY8 + 3
            Obj.Print "X"
        Case "CP" 'VTA SUJETA A CONFIRMACION    2
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 13.7
            'Obj.Print "x"
        Case "TF" 'compra   3
            Obj.CurrentX = 140
            Obj.CurrentY = CurY + 4.7
            'Obj.Print "x"
        Case "CO" 'CONSIGNACION  4
            Obj.CurrentX = 140
            Obj.CurrentY = CurY + 13.2
            'Obj.Print "x"
        Case "DP", "DS", "DC", "DR", "DR" 'DEVOLUCION  5
            Obj.CurrentX = 140
            Obj.CurrentY = CurY + 17.45
            'Obj.Print "x"
        Case "TR" 'Traslado entre establecimientos de la misma emp. 6
            Obj.CurrentX = 197
            Obj.CurrentY = CurY + 0.45
            'Obj.Print "x"
        Case "TS" 'Traslado de bienes para transformacion  7
            Obj.CurrentX = 140
            Obj.CurrentY = CurY + 8.25
            'Obj.Print "x"
        Case "TT" 'Traslado de bienes transformados  8
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 13.7
            'Obj.Print "x"
        Case "T1" 'TRASLADO POR EMISOR ITINERANTE  9
            Obj.CurrentX = 197
            Obj.CurrentY = CurY + 9.15
            'Obj.Print "x"
         Case "TZ" 'TRASLADO zona primaria 10
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 8
            'Obj.Print "x"
        Case "IM" 'IMPORTACION   11
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 13.7
            'Obj.Print "x"
        Case "EX" 'EXPORTACION   12
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 13.7
            'Obj.Print "x"
         Case Else 'ninguno de las anteriores  13
            Obj.CurrentX = 206
            Obj.CurrentY = CurY + 13.7
            'Obj.Print "x"
    End Select
    CurY = CurY + 20
    
    
'    'FECHA DE INICIO DE TRASLADO
'    Obj.CurrentX = 130 + xDerecha + 2
'    Obj.CurrentY = 78 + CurY + 1
'    Obj.Print Format$(rsCab!FECENT, "dd/MM/yyyy")
    
'    'HORA DE ENTREGA
'    Obj.CurrentX = 150
'    Obj.CurrentY = 64 + CurY
'    Obj.Print Trim$("" & RsCab!HorEnt)
    
'    'VENDEDOR
'    Obj.CurrentX = 150 + xDerecha
'    Obj.CurrentY = 74 + CurY + 1
'    Obj.Print Trim$("VENDEDOR  :  " & rsCab!CODVEN)
    
'    'NUMERO DE BULTOS
'    Obj.CurrentX = 170 + xDerecha + 4
'    Obj.CurrentY = 82 + CurY + 1
'    Obj.Print Trim$("" & rsCab!MonRf1)
'
'    'PESO
'    Obj.CurrentX = 100 + xDerecha - 3
'    Obj.CurrentY = 82 + CurY + 1
'    Obj.Print Trim$("" & rsCab!MonRf2)
    
    'IMPRIME LA FACTURA COMO REFERENCIA
''''    If rsDocRef2.EOF = False And rsDocRef2.BOF = False Then
''''        Obj.CurrentX = 83: Obj.CurrentY = CurY
''''        Obj.Print "ref" & Trim$("" & rsDocRef2!CodDoc) & "  -  " & "factura" & Trim$("" & rsDocRef2!factura)
''''    End If
    Dim sFechaEmiFc As String
    If rsDocRef.EOF = False And rsDocRef.BOF = False Then
       '  Obj.CurrentX = 40 + xDerecha
        ' Obj.CurrentY = 100 + CurY
'        Obj.Print Trim$("" & rsDocRef!CodDoc) & "  -  " & Trim$("" & rsDocRef!NumDoc)
       ' Obj.CurrentX = 45 + xDerecha
       'Obj.CurrentY = 78 + CurY + 1
'        Obj.Print Trim$("" & rsDocRef!FECDOC)
        If bImpReemplazarCodigo = True Then
         sFechaEmiFc = Trim$("" & rsDocRef!FecDoc)
        End If
   End If
'    Else
        If bImpReemplazarCodigo = True Then
        Obj.CurrentX = 40 + xDerecha
        Obj.CurrentY = 125 + CurY + 10
        Obj.Print Trim$("" & rsCab!REFINT)
        
        Obj.CurrentX = 80 + xDerecha + 50
        Obj.CurrentY = 125 + CurY + 3 + 10
        'Obj.Print Trim$("" & rsCab!FecEnt02)
        If sFechaEmiFc <> "" Then
            Obj.Print Trim$("" & sFechaEmiFc)
        Else
            Obj.Print Trim$("" & rsCab!FecEnt02)
        End If
        End If
'    End If
      iLonDet1 = 49
     
     'COMENTARIO
   '  If Len(Trim$("" & rsCab!commem)) > iLonDet1 Then
    ' campoBB = Trim("" & rsCab!commem) & ""
    ' poss = 0
    ' For II = 1 To Len(campoBB)
    ' poss = poss + iLonDet1
    ' Valor = Trim(Buscar_cadena(campoBB, "", II, poss))
    ' If II > 1 Then
    ' CurY = CurY + 3
    ' End If
    ' If poss < Len(campoBB) Then
     
   ' Obj.CurrentX = 35 + xDerecha + 9
   ' Obj.CurrentY = 90 + CurY + 1
  '  Obj.Print Mid$(Trim$("" & campoBB), II, (poss - 1) - II)
   ' Else
    ' Obj.CurrentX = 35 + xDerecha + 9
    'Obj.CurrentY = 90 + CurY + 1
  '  Obj.Print Mid$(Trim$("" & campoBB), II)
   
    'End If
    'II = poss - 1
    'Next II
     '   End If
        
        
        
    CurY = DCY10
    
    If Not RsSer.EOF Or Not RsSer.BOF Then
        RsSer.MoveFirst
    End If
    
    RsDet.MoveFirst
    Do Until RsDet.EOF
        If (Trim$("" & RsDet.Fields("NumIte1").Value) = "") Then
            CurY = CurY - 1 + 3
            Obj.CurrentY = CurY: Obj.CurrentX = DFX1 - Obj.TextWidth(RsDet!NumIte) + 12 + 3
            Obj.Print Trim$(RsDet!Codart)
            Obj.CurrentX = DFX3: Obj.CurrentY = CurY
            Obj.Print Trim$("" & RsDet!XTIPUNI)
            Obj.CurrentX = DFX4 - Obj.TextWidth(RsDet!CANTOT) + 10: Obj.CurrentY = CurY
            Obj.Print Format(RsDet!CANTOT, "#0.00")
            'Si elijo que reemplaze la descripcion
            If bImpReemplazarCodigo Then
                If Trim$("" & RsDet!DETART) <> "" Then
                    If Len(Trim$("" & RsDet!DETART)) > iLonDet Then
                        campoB = Trim$("" & RsDet!DETART) & ""
                        Pos = 0
                        For I = 1 To Len(campoB)
                            Pos = Pos + iLonDet
                            Valor = Trim(Buscar_cadena(campoB, " ", I, Pos))
                            If I > 1 Then
                                CurY = CurY + 3
                            End If
                            If Pos < Len(campoB) Then
                                Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                                Obj.Print Mid$(Trim$("" & campoB), I, (Pos - 1) - I)
                            Else
                                Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                                Obj.Print Mid$(Trim$("" & campoB), I)
                            End If
                            I = Pos - 1
                        Next I
                    Else
                        Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                        Obj.Print Trim$("" & RsDet!DETART)
                    End If
                Else
                    If Len(Trim$("" & RsDet!DESART)) > iLonDet Then
                        campoB = Trim$("" & RsDet!DESART) & ""
                        Pos = 0
                        For I = 1 To Len(campoB)
                            Pos = Pos + iLonDet
                            Valor = Trim(Buscar_cadena(campoB, " ", I, Pos))
                            If I > 1 Then
                                CurY = CurY + 3
                            End If
                            If Pos < Len(campoB) Then
                                Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                                Obj.Print Mid$(Trim$("" & campoB), I, (Pos - 1) - I)
                            Else
                                Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                                Obj.Print Mid$(Trim$("" & campoB), I)
                            End If
                            I = Pos - 1
                        Next I
                    Else
                        Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                        Obj.Print Trim$("" & RsDet!DESART)
                    End If
                End If
            Else
                If Len(Trim$("" & RsDet!DESART)) > iLonDet Then
                    campoB = Trim$("" & RsDet!DESART) & ""
                    Pos = 0
                    For I = 1 To Len(campoB)
                        Pos = Pos + iLonDet
                        Valor = Trim(Buscar_cadena(campoB, " ", I, Pos))
                        If I > 1 Then
                            CurY = CurY + 3
                        End If
                        If Pos < Len(campoB) Then
                            Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                            Obj.Print Mid$(Trim$("" & campoB), I, (Pos - 1) - I)
                        Else
                            Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                            Obj.Print Mid$(Trim$("" & campoB), I)
                        End If
                        I = Pos - 1
                    Next I
                Else
                    Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                    Obj.Print Trim$("" & RsDet!DESART)
                End If
            End If
            '******* DESART1 *********
            If Len(Trim$("" & RsDet!DesArt1)) > iLonDet Then
                CurY = CurY + 3
                campoB = Trim$("" & RsDet!DesArt1) & ""
                Pos = 0
                For I = 1 To Len(campoB)
                    Pos = Pos + iLonDet
                    Valor = Trim(Buscar_cadena(campoB, " ", I, Pos))
                    If I > 1 Then
                        CurY = CurY + 3
                    End If
                    If Pos < Len(campoB) Then
                        Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                        Obj.Print Mid$(Trim$("" & campoB), I, (Pos - 1) - I)
                    Else
                        Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                        Obj.Print Mid$(Trim$("" & campoB), I)
                    End If
                    I = Pos - 1
                Next I
            Else
                CurY = CurY + 3
                Obj.CurrentX = DFX2: Obj.CurrentY = CurY
                Obj.Print Trim$("" & RsDet!DesArt1)
            End If
        End If
        RsDet.MoveNext
    Loop
    
End Sub


Function GetTempTable(numero As Byte) As String
    On Error Resume Next
    GetTempTable = "##TMP" & Round(Rnd() * 10000, 0) & numero
    ICnx.CnEDB.Execute "Select Top 1 * From " & GetTempTable, , adCmdText
    If Err.Number = 0 Then GetTempTable = GetTempTable(numero) Else Err.Clear
End Function






