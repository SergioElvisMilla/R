VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpTipPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ICodDoc As String, INumDoc As String, IProImp As String
Private ICodEmp As String, ICodGru As String, ITitulo As String
Private ICopCab As String, ICopMsg As String
Private ISubTitulo As String, IDevice As String
Private INivel As Byte, ICnx As ADODB.Connection
Private pvWindow As Object, oSysAdmin As Object
Private iPaperSize As Integer, IHeight As Long, IWidth As Long, IFontSize As Integer
Private IFontName As String
Private sParam01 As String, sParam02 As String, sParam03 As String
Private sFormatCol() As String, sWidthCol() As Single, bImprimirCol() As Boolean, lColsCol() As Long
Private lCntPag As Long
Private rsImp As ADODB.Recordset
Private bGroupBy As Boolean, bCheckGroupBy As Boolean
Private IResponsable As String
Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set ICnx = Cnn
End Property
Public Property Set SysAdmin(ByVal oSys As Object)
    Set oSysAdmin = oSys
End Property
Public Property Let CodDoc(ByVal Data As String)
    ICodDoc = Trim$(Data)
End Property
Public Property Let NumDoc(ByVal Data As String)
    INumDoc = Trim$(Data)
End Property
Public Property Let ProImp(ByVal Data As String)
    IProImp = Trim$(Data)
End Property
Public Property Let CodEmp(ByVal Data As String)
    ICodEmp = Trim$(Data)
End Property
Public Property Let CodGru(ByVal Data As String)
    ICodGru = Trim$(Data)
End Property
Public Property Let MovCab(ByVal Data As String)
End Property
Public Property Let MovDet(ByVal Data As String)
End Property
Public Property Let RutImp(ByVal Data As String)
    IDevice = Data
End Property
Public Property Let Param01(ByVal Data As String)
    sParam01 = Data
End Property
Public Property Let Param02(ByVal Data As String)
    sParam02 = Data
End Property
Public Property Let Param03(ByVal Data As String)
    sParam03 = Data
End Property
Public Property Let Nivel(ByVal Data As Byte)
    INivel = Data
End Property
Public Property Let PaperSize(ByVal Data As Integer)
    iPaperSize = Data
End Property
Public Property Let Width(ByVal Data As Long)
    IWidth = Data
End Property
Public Property Let Height(ByVal Data As Long)
    IHeight = Data
End Property
Public Property Let FontSize(ByVal Data As Integer)
    IFontSize = Data
End Property
Public Property Let FontName(ByVal Data As String)
    IFontName = Data
End Property
Public Property Let Titulo(ByVal Data As String)
    ITitulo = Data
End Property
Public Property Let SubTitulo(ByVal Data As String)
    ISubTitulo = Data
End Property

Public Sub PrintShow(ByRef Rs As ADODB.Recordset, ByRef sFormat() As String, ByRef sWidth() As Single, ByRef bRound() As Byte, ByRef bSum() As Byte, ByRef bImprimir() As Boolean, ByRef sFiltro() As String, ByRef lCols() As Long, ByVal Titulo As String, ByVal Detalle01 As String, ByVal Detalle02 As String)
    If Rs Is Nothing Then Call MsgBox("No hay datos para imprimir", vbInformation, "Mensaje al Usuario"): Exit Sub
    If Rs.State = adStateClosed Then Call MsgBox("No hay datos para imprimir", vbInformation, "Mensaje al Usuario"): Exit Sub
    If Rs.EOF Or Rs.BOF Then Call MsgBox("No hay datos para imprimir", vbInformation, "Mensaje al Usuario"): Exit Sub
    Set rsImp = Rs.Clone
    sFormatCol() = sFormat: sWidthCol() = sWidth: bImprimirCol() = bImprimir: lColsCol() = lCols
    ITitulo = Titulo: ICopCab = Detalle01: ICopMsg = Detalle02: ISubTitulo = "del " & sFiltro(3, 2, 0) & " Al " & sFiltro(4, 2, 0)
    bCheckGroupBy = CheckGroupBy(sFiltro(), 4, 1, 4)
    bGroupBy = IIf(lCols(2) <> 0, True, False)
    
    If Trim(sFiltro(3, 0, 1)) = "" Then
        IResponsable = "Resp.: Todos"
    Else
        IResponsable = "Resp.: " & sFiltro(3, 0, 1)
    End If
    
    Call PrintDoc(rsImp, True)
    If rsImp.State <> adStateClosed Then rsImp.Close
    Set rsImp = Nothing
End Sub
Public Function PrintDoc(ByRef rsMain As ADODB.Recordset, Optional PreView As Boolean = False)
    Dim bPrint As Boolean, iCntCop As Integer
    Screen.MousePointer = 11
    'On Error GoTo Solucion
    bPrint = True
    If rsMain.State = adStateClosed Then Call MsgBox("No hay datos para imprimir", vbInformation, "Mensaje al Usuario"): Exit Function
    If oSysAdmin Is Nothing Then Call MsgBox("Objecto SysAdmin no encontrado. No se puede continuar.", vbExclamation, "Mensaje al Usuario"): Exit Function
    Call SetDefault
    If rsMain.EOF = False And rsMain.BOF = False Then
        Set pvWindow = CreateObject("PrtTools.cPreView")
        
        Dim prt As Printer
        Set prt = GetPrinter(IDevice)
        If Not prt Is Nothing Then
            Set Printer = prt
        End If
        Set prt = Nothing
        Set pvWindow.ActivePrinter = Printer
        Screen.MousePointer = 0
       
        If PreView Then
            lCntPag = 0
            Call PreviewData(rsMain, pvWindow.ActivePage)
            Screen.MousePointer = 0
            bPrint = pvWindow.ShowPreView()
        End If
        If bPrint = True Then
            For iCntCop = 1 To pvWindow.NroCop
                lCntPag = 0
                Call PrintData(rsMain, pvWindow.ActivePrinter)
                If (iCntCop = pvWindow.NroCop) Then
                    pvWindow.ActivePrinter.EndDoc
                Else
                    pvWindow.ActivePrinter.NewPage
                End If
                Screen.MousePointer = 0
            Next iCntCop
        End If
        Set pvWindow = Nothing
        PrintDoc = True
    Else
        Screen.MousePointer = 0
        Call MsgBox("No hay datos para imprimir", vbInformation, "Mensaje al Usuario")
    End If
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

Private Function ImprimirCol(ByVal oRsTmp As ADODB.Recordset, ByVal sColName As String, ByVal sValue As String) As Boolean
    Dim lPos As Long, lCol As Integer, oRsD As ADODB.Recordset, lCol02 As Integer, oRs As ADODB.Recordset
    Set oRs = oRsTmp.Clone
    oRs.Filter = "[" & Trim(sColName) & "] = '" & Trim(sValue) & "'"
    If Not oRs.EOF Then
        For lCol = 0 To oRs.Fields.Count - 1
            If oRs.Fields(lCol).Type = adChapter Then
                Set oRsD = oRs.Fields(lCol).Value
                If Not oRsD.EOF Then
                    For lCol02 = 0 To oRsD.Fields.Count - 1
                        If UCase(Trim(oRsD.Fields(lCol02).Name)) = UCase(Trim(sColName)) Then
                            ImprimirCol = bImprimirCol(lCol02 + 1)
                            Call Close_RecordSet(oRs)
                            Call Close_RecordSet(oRsD)
                            Exit Function
                        End If
                    Next
                End If
            End If
        Next
    End If
    Call Close_RecordSet(oRs)
    Call Close_RecordSet(oRsD)
End Function

Private Sub GetProperties(ByVal oRsTmp As ADODB.Recordset, ByVal sColName As String, ByVal sValue As String, ByRef bImprimir As Boolean, ByRef IWidth As Single)
    Dim lPos As Long, lCol As Integer, oRsD As ADODB.Recordset, lCol02 As Integer, oRs As ADODB.Recordset
    Set oRs = oRsTmp.Clone
    If Not oRs.EOF Then
        For lCol = 0 To oRs.Fields.Count - 1
            If oRs.Fields(lCol).Type = adChapter Then
                Set oRsD = oRs.Fields(lCol).Value
                If Not oRsD.EOF Then
                    For lCol02 = 0 To oRsD.Fields.Count - 1
                        If UCase(Trim(oRsD.Fields(lCol02).Name)) = UCase(Trim(sColName)) Then
                            bImprimir = bImprimirCol(lCol02 + 1)
                            Call Close_RecordSet(oRs)
                            Call Close_RecordSet(oRsD)
                            Exit Sub
                        End If
                    Next
                End If
            Else
                If UCase(Trim(oRs.Fields(lCol).Name)) = UCase(Trim(sColName)) Then
                    IWidth = sWidthCol(1, lCol + 2)
                End If
            End If
        Next
    End If
    Call Close_RecordSet(oRs)
    Call Close_RecordSet(oRsD)
End Sub

Private Function GetTotalRegistros(ByVal oRsTmp As ADODB.Recordset, ByVal sColName As String, ByVal sValue As String) As Long
    Dim lPos As Long, lCol As Integer, oRsD As ADODB.Recordset, lCol02 As Integer, oRs As ADODB.Recordset
    Set oRs = oRsTmp.Clone
    oRs.Filter = "[" & Trim(sColName) & "] = '" & Trim(sValue) & "'"
    If Not oRs.EOF Then
        For lCol = 0 To oRs.Fields.Count - 1
            If oRs.Fields(lCol).Type = adChapter Then
                Set oRsD = oRs.Fields(lCol).Value
                If Not oRsD.EOF Then GetTotalRegistros = oRsD.RecordCount
                Call Close_RecordSet(oRs)
                Call Close_RecordSet(oRsD)
                Exit Function
            End If
        Next
    End If
    Call Close_RecordSet(oRs)
    Call Close_RecordSet(oRsD)
End Function

Private Function GetIndiceU(ByRef lInd As Long, ByRef sMat() As String, ByVal sValor As String) As Long
    Dim lCnt As Long
    For lCnt = 1 To lInd
        If UCase(Trim(sMat(lCnt))) = UCase(Trim(sValor)) Then GetIndiceU = lCnt: Exit Function
    Next
End Function

Private Sub PreviewData(ByRef rsMain As ADODB.Recordset, ByRef Obj As Object)
    'On Error GoTo MalImpreso
    Dim lDx As Long, lDy As Long, CurY As Long, bExit As Boolean, iCnt As Integer, CurX As Long, lCnt As Long, iCnt2 As Integer
    Dim lCab As Long, bImprimir As Boolean, sWidth As Single
    Dim RsDet As ADODB.Recordset, lInd As Long, dTotal() As Double, sNomColTot() As String, lFnd As Long, dTotGru As Double, dTotGra As Double
    With Obj
        
        If rsMain.RecordCount > 0 Then rsMain.MoveFirst
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Draft 17cpi"
        Obj.Font.Size = 9
        
        .Height = 3700 + ((rsMain.RecordCount * 100) + 2100)
        .Width = 4251.968503937
        
        Call PrintCabecera(Obj, True)
        
        .CurrentY = 26: CurY = .CurrentY + lDy
        
        'Si esta agrupado el tablero
        If bGroupBy Then
            'Si esta resumido
            If bCheckGroupBy Then
                'REGISTRO = Cuando se crea agrupados con el tablero te crea 2 recordsets, uno padre y uno hijo
                'El hijo lo reconoce como REGISTRO
                Do Until rsMain.EOF
                    lInd = 0: ReDim dTotal(1 To 1) As Double: ReDim sNomColTot(1 To 1) As String
                    dTotGru = 0
                    For iCnt = 0 To rsMain.Fields.Count - 1
                        If rsMain.Fields(iCnt).Type <> adChapter Then
                            Call GetProperties(rsMain, rsMain.Fields(iCnt).Name, Trim("" & rsMain.Fields(iCnt).Value), bImprimir, sWidth)
                            If bImprimir Then
                                Select Case rsMain.Fields(iCnt).Type
                                    Case adSingle, adNumeric, adDouble
                                    Case Else
                                        .CurrentX = 5 + lDx
                                        .CurrentY = CurY
                                        Obj.Print Mid(Trim$("" & rsMain.Fields(iCnt).Value), 1, 30)
                                End Select
                            End If
                        Else
                            Set RsDet = rsMain.Fields(iCnt).Value
                            Do Until RsDet.EOF
                                For iCnt2 = 0 To RsDet.Fields.Count - 1
                                    If bImprimirCol(iCnt2 + 1) Then
                                        Select Case RsDet.Fields(iCnt2).Type
                                            Case adSingle, adNumeric, adDouble
                                                lFnd = GetIndiceU(lInd, sNomColTot(), Trim("" & RsDet.Fields(iCnt2).Name))
                                                If lFnd > 0 Then
                                                    dTotal(lFnd) = dTotal(lFnd) + Format(Val(Trim$("" & RsDet.Fields(iCnt2).Value)), "#.00")
                                                Else
                                                    lInd = lInd + 1: ReDim Preserve dTotal(1 To lInd) As Double: ReDim Preserve sNomColTot(1 To lInd) As String
                                                    dTotal(lInd) = dTotal(lInd) + Format(Val(Trim$("" & RsDet.Fields(iCnt2).Value)), "#.00")
                                                    sNomColTot(lInd) = Trim$("" & RsDet.Fields(iCnt2).Name)
                                                End If
                                        End Select
                                    End If
                                Next
                                RsDet.MoveNext
                            Loop
                            If lInd > 0 Then
                                CurY = CurY + 3
                                For iCnt2 = 1 To lInd
                                    .CurrentX = 10 + lDx
                                    .CurrentY = CurY
                                    Obj.Print sNomColTot(iCnt2)
                                    
                                    .CurrentX = 60 + lDx - .TextWidth(Format(dTotal(iCnt2), "#.00")) ' + 30
                                    .CurrentY = CurY
                                    Obj.Print Format(dTotal(iCnt2), "#.00")
                                    
                                    If InStr(1, sNomColTot(iCnt2), "S/", vbTextCompare) > 0 Then dTotGru = dTotGru + Format(dTotal(iCnt2), "#.00")
                                    
                                    CurY = CurY + 3
                                Next
                                
                                CurY = CurY + 3
                                Obj.Line (45, CurY)-(68, CurY)
                                CurY = CurY + 1
                                
                                .CurrentX = 5 + lDx
                                .CurrentY = CurY
                                Obj.Print "Total en S/."
                                
                                .CurrentX = 60 + lDx - .TextWidth(Format(dTotGru, "#.00")) '+ 30
                                .CurrentY = CurY
                                Obj.Print Format(dTotGru, "#.00")
                                
                                dTotGra = dTotGra + dTotGru
                                
                                CurY = CurY + 3
                                
                            End If
                        End If
                    Next iCnt
                    CurY = CurY + 3
                    rsMain.MoveNext
                Loop
                
                If dTotGra > 0 Then
                    .CurrentX = 5 + lDx
                    .CurrentY = CurY
                    Obj.Print "Total Gral en S/."
                    
                    .CurrentX = 60 + lDx - .TextWidth(Format(dTotGra, "#.00")) '+ 30
                    .CurrentY = CurY
                    Obj.Print Format(dTotGra, "#.00")
                End If
                
            Else
                Call MsgBox("El formato solo se puede imprimir si el reporte esta agrupado y resumido", vbInformation, "Mensaje al Usuario"): Exit Sub
            End If
        Else
            Call MsgBox("El formato solo se puede imprimir si el reporte esta agrupado y resumido", vbInformation, "Mensaje al Usuario"): Exit Sub
        End If
        
    End With
    
MalImpreso:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "¿Desea ver mas información?", vbExclamation + vbYesNo, "Error") = vbYes Then
            Call MsgBox(Err.Description, vbCritical, "Error")
            Err.Clear
        End If
    End If
End Sub

Private Sub PrintData(ByRef rsMain As ADODB.Recordset, ByRef Obj As Object)
    'On Error GoTo MalImpreso
    Dim lDx As Long, lDy As Long, CurY As Long, bExit As Boolean, iCnt As Integer, CurX As Long, lCnt As Long, iCnt2 As Integer
    Dim lCab As Long, bImprimir As Boolean, sWidth As Single
    Dim RsDet As ADODB.Recordset, lInd As Long, dTotal() As Double, sNomColTot() As String, lFnd As Long, dTotGru As Double, dTotGra As Double
    With Obj
        
        If rsMain.RecordCount > 0 Then rsMain.MoveFirst
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Draft 17cpi"
        Obj.Font.Size = 8
        
        .Height = 3700 + ((rsMain.RecordCount * 100) + 2100)
        .Width = 6400
        
        Call PrintCabecera(Obj, False)
        
        .CurrentY = 26: CurY = .CurrentY + lDy
        
        'Si esta agrupado el tablero
        If bGroupBy Then
            'Si esta resumido
            If bCheckGroupBy Then
                'REGISTRO = Cuando se crea agrupados con el tablero te crea 2 recordsets, uno padre y uno hijo
                'El hijo lo reconoce como REGISTRO
                Do Until rsMain.EOF
                    lInd = 0: ReDim dTotal(1 To 1) As Double: ReDim sNomColTot(1 To 1) As String
                    dTotGru = 0
                    For iCnt = 0 To rsMain.Fields.Count - 1
                        If rsMain.Fields(iCnt).Type <> adChapter Then
                            Call GetProperties(rsMain, rsMain.Fields(iCnt).Name, Trim("" & rsMain.Fields(iCnt).Value), bImprimir, sWidth)
                            If bImprimir Then
                                Select Case rsMain.Fields(iCnt).Type
                                    Case adSingle, adNumeric, adDouble
                                    Case Else
                                        .CurrentX = 5 + lDx
                                        .CurrentY = CurY
                                        Obj.Print Mid(Trim$("" & rsMain.Fields(iCnt).Value), 1, 30)
                                End Select
                            End If
                        Else
                            Set RsDet = rsMain.Fields(iCnt).Value
                            Do Until RsDet.EOF
                                For iCnt2 = 0 To RsDet.Fields.Count - 1
                                    If bImprimirCol(iCnt2 + 1) Then
                                        Select Case RsDet.Fields(iCnt2).Type
                                            Case adSingle, adNumeric, adDouble
                                                lFnd = GetIndiceU(lInd, sNomColTot(), Trim("" & RsDet.Fields(iCnt2).Name))
                                                If lFnd > 0 Then
                                                    dTotal(lFnd) = dTotal(lFnd) + Format(Val(Trim$("" & RsDet.Fields(iCnt2).Value)), "#.00")
                                                Else
                                                    lInd = lInd + 1: ReDim Preserve dTotal(1 To lInd) As Double: ReDim Preserve sNomColTot(1 To lInd) As String
                                                    dTotal(lInd) = dTotal(lInd) + Format(Val(Trim$("" & RsDet.Fields(iCnt2).Value)), "#.00")
                                                    sNomColTot(lInd) = Trim$("" & RsDet.Fields(iCnt2).Name)
                                                End If
                                        End Select
                                    End If
                                Next
                                RsDet.MoveNext
                            Loop
                            If lInd > 0 Then
                                CurY = CurY + 6
                                For iCnt2 = 1 To lInd
                                    .CurrentX = 10 + lDx
                                    .CurrentY = CurY
                                    Obj.Print sNomColTot(iCnt2)
                                    
                                    .CurrentX = 60 + lDx - .TextWidth(Format(dTotal(iCnt2), "#.00")) + 30
                                    .CurrentY = CurY
                                    Obj.Print Format(dTotal(iCnt2), "#.00")
                                    
                                    If InStr(1, sNomColTot(iCnt2), "S/", vbTextCompare) > 0 Then dTotGru = dTotGru + Format(dTotal(iCnt2), "#.00")
                                    
                                    CurY = CurY + 6
                                Next
                                
                                CurY = CurY + 3
                                
                                .CurrentX = 59 + lDx
                                .CurrentY = CurY - 3
                                Obj.Print "----------------"
                                
                                CurY = CurY + 1
                                
                                .CurrentX = 5 + lDx
                                .CurrentY = CurY
                                Obj.Print "Total en S/."
                                
                                .CurrentX = 60 + lDx - .TextWidth(Format(dTotGru, "#.00")) + 30
                                .CurrentY = CurY
                                Obj.Print Format(dTotGru, "#.00")
                                
                                dTotGra = dTotGra + dTotGru
                                
                                CurY = CurY + 3
                                
                            End If
                        End If
                    Next iCnt
                    CurY = CurY + 3
                    rsMain.MoveNext
                Loop
                
                If dTotGra > 0 Then
                    .CurrentX = 5 + lDx
                    .CurrentY = CurY
                    Obj.Print "Total Gral en S/."
                    
                    .CurrentX = 60 + lDx - .TextWidth(Format(dTotGra, "#.00")) + 30
                    .CurrentY = CurY
                    Obj.Print Format(dTotGra, "#.00")
                End If
                
            Else
                Call MsgBox("El formato solo se puede imprimir si el reporte esta agrupado y resumido", vbInformation, "Mensaje al Usuario"): Exit Sub
            End If
        Else
            Call MsgBox("El formato solo se puede imprimir si el reporte esta agrupado y resumido", vbInformation, "Mensaje al Usuario"): Exit Sub
        End If
        
    End With
    
MalImpreso:
    Screen.MousePointer = 0
    If Err.Number <> 0 Then
        If MsgBox("Se genero un error durante la operacion por lo que no pudo ser completada" & vbCrLf & "¿Desea ver mas información?", vbExclamation + vbYesNo, "Error") = vbYes Then
            Call MsgBox(Err.Description, vbCritical, "Error")
            Err.Clear
        End If
    End If
End Sub
Private Sub SetDefault()
    If iPaperSize = 0 Then iPaperSize = 9
    If IWidth = 0 Then IWidth = 4200
    If IHeight = 0 Then IHeight = 16838
    If IFontSize = 0 Then IFontSize = 9
    If Trim$("" & IFontName) = "" Then IFontName = "Draft 17cpi"
End Sub

Private Sub PrintCabecera(ByRef Obj As Object, ByVal bVisPre As Boolean)
    With Obj
        If bVisPre Then
            .CurrentX = 37 - (.TextWidth(Trim$("" & oSysAdmin.NomEmp)) / 2)
            .CurrentY = 5
            Obj.Print Trim$("" & oSysAdmin.NomEmp)
            
            ITitulo = "Ventas por Tipo de Pago"
            
            .CurrentX = 37 - (.TextWidth(Trim$("" & ITitulo)) / 2)
            .CurrentY = 9
            Obj.Print Trim$("" & ITitulo)
            
            .CurrentX = 37 - (.TextWidth(Trim$("" & ISubTitulo)) / 2)
            .CurrentY = 13
            Obj.Print Trim$("" & ISubTitulo)
            
            .CurrentX = 37 - (Obj.TextWidth(Trim$("" & IResponsable)) / 2)
            .CurrentY = 17
            Obj.Print Trim$("" & IResponsable)
            
            .CurrentX = 5
            .CurrentY = 21
            Obj.Print "--------------------------------------------------------------"
            
            'Obj.Line (5, 20)-(68, 20)
        Else
            .CurrentX = 50 - (.TextWidth(Trim$("" & oSysAdmin.NomEmp)) / 2)
            .CurrentY = 5
            Obj.Print Trim$("" & oSysAdmin.NomEmp)
            
            ITitulo = "Ventas por Tipo de Pago"
            
            .CurrentX = 50 - (.TextWidth(Trim$("" & ITitulo)) / 2)
            .CurrentY = 9
            Obj.Print Trim$("" & ITitulo)
            
            .CurrentX = 50 - (.TextWidth(Trim$("" & ISubTitulo)) / 2)
            .CurrentY = 13
            Obj.Print Trim$("" & ISubTitulo)
            
            .CurrentX = 50 - (Obj.TextWidth(Trim$("" & IResponsable)) / 2)
            .CurrentY = 17
            Obj.Print Trim$("" & IResponsable)
            
            .CurrentX = 5
            .CurrentY = 21
            Obj.Print "---------------------------------------"
        End If
    End With
End Sub


