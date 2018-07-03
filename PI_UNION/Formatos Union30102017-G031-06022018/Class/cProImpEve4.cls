VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpEve4"
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
Private sNomDoc As String
Private bDetallado As Boolean
Dim rsCab As ADODB.Recordset, rsEmp As ADODB.Recordset
Dim pvWindow As Object
Dim nPag As Long
Private sNumItems As String

Private iNCopiasEti As Integer
Private iNumIniCor As Integer

'Ultima Modificacion 11/12/2004
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
Public Property Let NumIte(ByVal Cad As String)
    sNumItems = Trim$(Cad)
End Property
Public Property Let NCopiasEti(ByVal Cad As Integer)
    iNCopiasEti = Trim$(Cad)
End Property
Public Property Let NumIniCor(ByVal Cad As Integer)
    iNumIniCor = Trim$(Cad)
End Property
Public Property Let Detallado(ByVal NewValue As Boolean)
    bDetallado = NewValue
End Property
Public Sub PrintShow()
    Call PrintDoc(True)
End Sub
Public Function PrintDoc(Optional PreView As Boolean = False)
    Dim RsDetLar As ADODB.Recordset
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
    Dim frm As FrmProImpCo
    Dim bPrint As Boolean
    Dim bImpRec As Boolean
    Dim rsTmp As ADODB.Recordset
    bPrint = True
    'Set rsTmp = ICnx.Execute("Select Top 1 ArcCab, ArcDet, Descri From CfgDoc00 Where CodEmp='" & sCodEmp & "' And CodDoc='" & sCodDoc & "'")
    'If Not rsTmp.EOF Or Not rsTmp.BOF Then sNomDoc = Trim$("" & rsTmp!DESCRI): sMovCab = Trim$("" & rsTmp.Fields("ArcCab").Value): sMovDet = Trim$("" & rsTmp.Fields("ArcDet").Value)
'    rsTmp.Close: Set rsTmp = Nothing
'    Set rsTmp = ICnx.Execute("Select Top 1 * From MarOpe " & _
'                            "Where CodDoc = '" & sCodDoc & "' And NumDoc = '" & sNumDoc & "' And NumIte In ('" & Replace(sNumItems, ";", "','", , , vbTextCompare) & "') And EstTik = 'N'", , adCmdText)
'    If rsTmp.EOF Or rsTmp.BOF Then
'        Screen.MousePointer = 11
'        MsgBox "No se encontro ningun ticket", vbExclamation
'        If Not frm Is Nothing Then Call Unload(frm)
'        Set frm = Nothing: rsTmp.Close: Set rsTmp = Nothing
'        Exit Function
'    End If
    Screen.MousePointer = 11
    '=============================
    Set rsEmp = ICnx.Execute("Select Top 1 CodEmp,NomEmp,DirEmp,TelEmp,FaxEmp,BkupEmp,RucEmp,RutLog,eMail,Ciudad From Empresas Where CodEmp='" & sCodEmp & "'", , adCmdText)
    If rsEmp.EOF And rsEmp.BOF Then
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
        Call PrintData(pvWindow.ActivePage)
        Screen.MousePointer = 0
        bPrint = pvWindow.ShowPreView()
    End If
    If bPrint = True Then
        Call PrintData(pvWindow.ActivePrinter)
        pvWindow.ActivePrinter.EndDoc
        Screen.MousePointer = 0
    End If
    Set pvWindow = Nothing
    Call Close_RecordSet(rsTmp)
    Call Close_RecordSet(rsEmp)
    PrintDoc = True
End Function
Private Sub PrintData(Obj As Object)
    Dim sTmp() As String, lTmp As Long, lCnt As Long, sLn() As String, lLn As Long
    Dim rsTmp As ADODB.Recordset, sSql As String, lTop As Long, sCodBar As String
    Dim sWidth(1 To 8) As Single
    Dim sCurY As Single, lNumReg As Long
    Dim xCopy As Integer, sCurX As Single
    sSql = "Select rTrim(Mc.CodDoc) + '|' + rTrim(Mc.NumDoc) + '|' + rTrim(Md.NumIte) + '|' NumTik, Mc.NomAne Cliente, Mc.NumDoc Orden, IsNull(Md.FecIte,Md.FecDoc) Fecha, Case When rTrim(IsNull(Pmt.NomTra1,'')) = '' Then rTrim(An.NomAne) Else rTrim(Pmt.NomTra1) + ' ' + rTrim(Pmt.ApellPat) + ' ' + rTrim(Pmt.ApellMat) End Operario, Mc.CodDoc, Mc.NumDoc" & vbCrLf & _
            ", Md.CodArt Maq, Md.DesArt1 Descripcion, Md.DesArt Trabajo, Md.CanRea, Md.CodAne3 CodSup, Md.Bol04, CTA.CodDocPad,CTA.NumDocPad" & vbCrLf & _
            " From " & sMovCab & " Mc " & vbCrLf & _
            " Inner Join " & sMovDet & " Md On Md.CodDoc = Mc.CodDoc And Md.NumDoc = Mc.NumDoc" & vbCrLf & _
            " Left Join PlMaeTra Pmt On Pmt.CodAne = Md.AneRef And Pmt.CodEmp = '" & sCodEmp & "'" & vbCrLf & _
            " Left Join Anexo An On An.CodAne = Md.AneRef" & vbCrLf & _
            " LEFT JOIN MOVCTAART CTA ON Cta.CodEmp = '" & sCodEmp & "' And Cta.CodDoc=Mc.CodDoc And CTA.NUMDOC = MC.NUMDOC AND CTA.NUMITECTA = '01'" & vbCrLf & _
            " Where Md.CodDoc = '" & sCodDoc & "' And Md.NumDoc = '" & sNumDoc & "' " & vbCrLf & _
            " And Md.NumIte In ('" & Replace(sNumItems, ";", "','", , , vbTextCompare) & "')"
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = ICnx.Execute(sSql, , adCmdText)
    If rsTmp.EOF Or rsTmp.BOF Then
        Screen.MousePointer = 11
        MsgBox "No se encontro ningun item", vbExclamation
    End If
    If iNCopiasEti = 0 Then iNCopiasEti = 1
    Obj.Font.Name = "Arial"
    If TypeName(Obj) <> "PictureBox" Then
        Call Change_PaperSize(Obj)
'        Obj.Orientation = 1
        Obj.PaperSize = 9
        Obj.Orientation = 2
        'Obj.Orientation = vbPRORLandscape
        Obj.Height = 16520
        Obj.Width = 11900 '11900
    Else
        Obj.Height = 16520
        Obj.Width = 17170
        'Obj.Align =
    End If
'    If TypeName(Obj) <> "PictureBox" Then
'        Call Change_PaperSize(Obj)
''        Obj.Orientation = 1
'        Obj.Height = 11900
'        Obj.Width = 5100
'    Else
'        Obj.Height = 11900
'        Obj.Width = 5100
'    End If
    Obj.ScaleMode = 6
    Printer.Font.Name = "Arial": Printer.ScaleMode = 6: Printer.Font.Size = 9: Printer.Font.Bold = True
    Do Until rsTmp.EOF
        For xCopy = 1 To iNCopiasEti
            lNumReg = lNumReg + 1
            Dim superX As Integer
            superX = 205
            sWidth(1) = 2 + superX: sWidth(2) = 5 + superX: sWidth(3) = 29 + superX: sWidth(4) = 39 + superX
            sWidth(5) = 42 + superX: sWidth(6) = 52 + superX: sWidth(7) = 80 + superX: sWidth(8) = 82 + superX
            
            Printer.Font.Name = "BarCode 128": Printer.Font.Size = 26: Printer.Font.Bold = False
            Obj.Font.Name = "BarCode 128": Obj.Font.Size = 26: Obj.Font.Bold = False
            sCodBar = Trim$("" & rsTmp.Fields("NumTik").Value)
            sCodBar = PatronCodigoBarra128B(sCodBar)
            sCurX = (sWidth(8) - Printer.TextWidth(sCodBar))
            Obj.CurrentX = sCurX: Obj.CurrentY = sCurY + 7
            Obj.Print sCodBar
            
            sCurY = sCurY + 10
            Printer.Font.Name = "Arial": Printer.Font.Size = 9
            Obj.Font.Name = "Arial": Obj.Font.Size = 14: Obj.Font.Bold = True
            Obj.CurrentX = sCurX: Obj.CurrentY = sCurY + 7
            Obj.Print Trim$("" & rsTmp.Fields("NumTik").Value)
            sCurY = sCurY + 10
            Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 5
            Obj.Font.Size = 12
            Obj.Font.Bold = True
            Obj.Print UCase$(Trim$("" & rsEmp.Fields("NomEmp").Value))
            
''''            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 18
''''            Obj.Print "CLIENTE"
''''            Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + Printer.TextWidth("CLIENTE"): Obj.CurrentY = sCurY + 18
''''            Obj.Print Trim$("" & rsTmp.Fields("Cliente").Value)
''''            Obj.DrawWidth = 1
''''            Obj.DrawStyle = vbDot: Obj.Line ((sWidth(2) + 13), (sCurY + 22))-(sWidth(8), (sCurY + 22))

            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 13
            Obj.Print "ORDEN DE PEDIDO No "
            Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + Printer.TextWidth("NumDoc") + 19: Obj.CurrentY = sCurY + 13
            Obj.Print Trim$("" & rsTmp.Fields("NumDoc").Value)
            Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("NumDoc") + 19, sCurY + 17)-(sWidth(8), sCurY + 17)
            
            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 18
            Obj.Print "ORDEN"
            Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + Printer.TextWidth("ORDEN"): Obj.CurrentY = sCurY + 18
            Obj.Print Trim$("" & rsTmp.Fields("Orden").Value)
            Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("ORDEN"), sCurY + 22)-(sWidth(8) - 33, sCurY + 22)
            
            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(8) - 30: Obj.CurrentY = sCurY + 18
            Obj.Print "FECHA"
            Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(8) - 27 + Printer.TextWidth("FECHA"): Obj.CurrentY = sCurY + 18
            Obj.Print Format(Trim$("" & rsTmp.Fields("Fecha").Value), "dd/MM/yyyy")
            Obj.DrawStyle = vbDot: Obj.Line (sWidth(8) - 27 + Printer.TextWidth("FECHA"), sCurY + 22)-(sWidth(8), sCurY + 22)
            
''            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 30
''            Obj.Print "OPERARIO"
''            Call Cortar(Trim$("" & rsTmp.Fields("Operario").Value), sWidth(8) - sWidth(1), False, False, sLn(), lLn, True, sWidth(8) - Printer.TextWidth("OPERARIO") - sWidth(1))
''
''            For lCnt = 1 To 1
''                Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + IIf((lCnt = 1), 15, 0): Obj.CurrentY = sCurY + 24 + (lCnt * 6)
''                If lCnt <= lLn Then Obj.Print sLn(lCnt)
''                Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + IIf((lCnt = 1), 15, 0), sCurY + 28 + (lCnt * 6))-(sWidth(8), sCurY + 28 + (lCnt * 6))
''            Next lCnt
            
''            sCurY = sCurY + 5
''            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 30
''            Obj.Print "MAQ"
''            Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + Printer.TextWidth("MAQ"): Obj.CurrentY = sCurY + 30
''            Obj.Print Trim$("" & rsTmp.Fields("Maq").Value)
''            Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("MAQ"), sCurY + 34)-(sWidth(8), sCurY + 34)
            
            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 23
            Obj.Print "TRABAJO"
            Call Cortar(Trim$("" & rsTmp.Fields("TRABAJO").Value), sWidth(8) - sWidth(1), False, True, sLn(), lLn, True, sWidth(8) - Printer.TextWidth("TRABAJO") - sWidth(1))
                        
            For lCnt = 1 To 3
                Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + IIf((lCnt = 1), 15, 0): Obj.CurrentY = sCurY + 17 + (lCnt * 6)
                If lCnt <= lLn Then Obj.Print sLn(lCnt)
                'Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + IIf((lCnt = 1), 15, 0), sCurY + 21 + (lCnt * 6))-(sWidth(8), sCurY + 21 + (lCnt * 6))
            Next lCnt
            
            Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("TRABAJO") - 15, sCurY + 28)-(sWidth(8), sCurY + 28)
            
            sCurY = sCurY + 6

           Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("TRABAJO") - 15, sCurY + 28)-(sWidth(8), sCurY + 28)

            sCurY = sCurY + 6

          Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + Printer.TextWidth("TRABAJO") - 15, sCurY + 28)-(sWidth(8), sCurY + 28)
'
        
            
            Obj.Font.Size = 9: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1): Obj.CurrentY = sCurY + 29
            Obj.Print "DESCRIPCION"
            Call Cortar(IIf(CBool(Trim$("" & rsTmp.Fields("Bol04").Value)), "REPROCESO DE TRABAJO" & vbCrLf, "") & Trim$("" & rsTmp.Fields("Descripcion").Value), sWidth(8) - sWidth(1), True, True, sLn(), lLn, True, sWidth(8) - Printer.TextWidth("DESCRIPCION") - sWidth(1))
            For lCnt = 1 To 8 '10
                Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = sWidth(2) + IIf((lCnt = 1), Printer.TextWidth("DESCRIPCION"), 0): Obj.CurrentY = sCurY + 23 + (lCnt * 6)
                If lCnt <= lLn Then Obj.Print sLn(lCnt)
                Obj.DrawStyle = vbDot: Obj.Line (sWidth(2) + IIf((lCnt = 1), Printer.TextWidth("DESCRIPCION"), 0), sCurY + 27 + (lCnt * 6))-(sWidth(8), sCurY + 27 + (lCnt * 6))
            Next lCnt
            
            'Obj.Font.Size = 9: Obj.Font.Bold = False: Obj.CurrentX = (sWidth(8) - 5 - Printer.TextWidth("T. Sup :" & Trim$("" & rsTmp.Fields("CanRea").Value))): Obj.CurrentY = sCurY + 90
            'Obj.Print "T. Sup :" & Trim$("" & rsTmp.Fields("CanRea").Value)
            lCnt = 0
'''            If bDetallado Then
'''                Obj.DrawStyle = vbSolid
'''                For lCnt = 1 To 6
'''                    Obj.DrawWidth = 2: lTop = 60 + (lCnt * 12): Obj.Line (sWidth(1), sCurY + lTop)-(sWidth(4), sCurY + lTop): Obj.Line (sWidth(5), sCurY + lTop)-(sWidth(7), sCurY + lTop)
'''                    Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(2): Obj.CurrentY = sCurY + 62 + (lCnt * 12): Obj.Print "E"
'''                    Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(6) + 4: Obj.CurrentY = sCurY + 62 + (lCnt * 12): Obj.Print "E"
'''                    Obj.Line (sWidth(3), sCurY + lTop)-(sWidth(3), sCurY + 72 + (lCnt * 12)): Obj.Line (sWidth(4), sCurY + lTop)-(sWidth(4), sCurY + 72 + (lCnt * 12)): Obj.Line (sWidth(5), sCurY + lTop)-(sWidth(5), sCurY + 72 + (lCnt * 12)): Obj.Line (sWidth(6), sCurY + lTop)-(sWidth(6), sCurY + 72 + (lCnt * 12))
'''                    Obj.DrawWidth = 1: lTop = 66 + (lCnt * 12): Obj.Line (sWidth(1), sCurY + lTop)-(sWidth(3), sCurY + lTop): Obj.Line (sWidth(6), sCurY + lTop)-(sWidth(7), sCurY + lTop)
'''                    Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(2): Obj.CurrentY = sCurY + 68 + (lCnt * 12): Obj.Print "T"
'''                    Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(6) + 4: Obj.CurrentY = sCurY + 68 + (lCnt * 12): Obj.Print "T"
'''                Next lCnt
'''                Obj.DrawWidth = 2: lTop = 60 + (lCnt * 12): Obj.Line (sWidth(1), sCurY + lTop)-(sWidth(4), sCurY + lTop): Obj.Line (sWidth(5), sCurY + lTop)-(sWidth(7), sCurY + lTop)
'''                Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(2): Obj.CurrentY = sCurY + 62 + (lCnt * 12): Obj.Print "TOTAL HORAS"
'''                Obj.Line (sWidth(4), sCurY + lTop)-(sWidth(4), sCurY + 72 + (lCnt * 12)): Obj.Line (sWidth(5), sCurY + lTop)-(sWidth(5), sCurY + 72 + (lCnt * 12))
'''                Obj.DrawWidth = 1: lTop = 66 + (lCnt * 12): Obj.Line (sWidth(1), sCurY + lTop)-(sWidth(4), sCurY + lTop)
'''                Obj.DrawWidth = 2: lTop = 72 + (lCnt * 12): Obj.Line (sWidth(1), sCurY + lTop)-(sWidth(4), sCurY + lTop): Obj.Line (sWidth(5), sCurY + lTop)-(sWidth(7), sCurY + lTop)
'''            End If
'''                Printer.Font.Name = "BarCode 128": Printer.Font.Size = 26: Printer.Font.Bold = False
'''                Obj.Font.Name = "BarCode 128": Obj.Font.Size = 26: Obj.Font.Bold = False
'''                sCodBar = Trim$("" & rsTmp.Fields("CodSup").Value)
'''                sCodBar = PatronCodigoBarra128B(sCodBar)
'''                Obj.CurrentX = sWidth(6): Obj.CurrentY = sCurY + 74 + (lCnt * 12)
'''                Obj.Print sCodBar
'''            Printer.Font.Name = "Arial": Obj.Font.Name = "Arial"
'''            Obj.Line (sWidth(1), sCurY + 85 + (lCnt * 12))-(sWidth(1) + 30, sCurY + 85 + (lCnt * 12)): Obj.Line (sWidth(3), sCurY + 85 + (lCnt * 12))-(sWidth(3) + 30, sCurY + 85 + (lCnt * 12))
'''            Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(1) + 13: Obj.CurrentY = sCurY + 85 + (lCnt * 12): Obj.Print "VB1"
'''            Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(3) + 13: Obj.CurrentY = sCurY + 85 + (lCnt * 12): Obj.Print "VB2"
'''            Obj.Font.Size = 8: Obj.Font.Bold = True: Obj.CurrentX = sWidth(6) + 2: Obj.CurrentY = sCurY + 85 + (lCnt * 12): Obj.Print "COD RESP."
            If (lNumReg = 1) And (xCopy <> iNCopiasEti) Then
                If TypeName(Obj) = "PictureBox" Then pvWindow.NewPage: Set Obj = pvWindow.ActivePage Else Obj.NewPage
                lNumReg = 0
            End If
        Next xCopy 'numero de copias de Impresion del Item actual
        rsTmp.MoveNext
        If ((Not rsTmp.EOF) And (lNumReg = 1)) Then
            If TypeName(Obj) = "PictureBox" Then pvWindow.NewPage: Set Obj = pvWindow.ActivePage Else Obj.NewPage
            lNumReg = 0
        End If
    Loop
    rsTmp.Close: Set rsTmp = Nothing
End Sub
Private Sub Change_PaperSize(ByRef Obj As Object)
    On Error GoTo Solucion
    Obj.PaperSize = 9
    Exit Sub
Solucion:
    Err.Clear
End Sub
