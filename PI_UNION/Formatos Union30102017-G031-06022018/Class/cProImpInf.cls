VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ICodDoc As String
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
Private rsCab As ADODB.Recordset
Private rsEmp As ADODB.Recordset
Private rsHis As ADODB.Recordset
Private pvWindow As Object
Private sFecIni As String, sFecFin As String

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
    Dim sCurDevice As String, bPrint As Boolean, sSql As String, oFrm As frmImpInf
    Screen.MousePointer = 11: bPrint = True
    
    Set rsEmp = ICnx.Execute("Select Top 1 NomEmp,DirEmp,TelEmp,FaxEmp,RucEmp,RutLog From Empresas Where CodEmp = '" & ICodEmp & "'", , adCmdText)
    
    Set rsCab = ICnx.Execute("Select Mc.CodDoc, Mc.NumDoc, Mc.NomAne, Mc.FecDoc " & vbCrLf & _
                             "From " & IMovCab & " Mc " & vbCrLf & _
                             "Where Mc.CodDoc = '" & Trim(ICodDoc) & "' And Mc.NumDoc='" & INumDoc & "'", , adCmdText)
                             
    Screen.MousePointer = 0
    Set oFrm = New frmImpInf
    Call oFrm.ShowForm
    If Not oFrm.Ok Then GoTo Finalizar
    
    sFecIni = oFrm.FecIni
    sFecFin = oFrm.FecFin
    
    Screen.MousePointer = 11
     
    Set rsHis = ICnx.Execute("Select AnR.CodAne, AnR.NomAne NomRes, Mh.FecCrea FecIni, Mh.Mod2 FecFin, MemHis Detalle, Sum(Convert(Numeric(18,4),(DateDiff(n,Mh.FecCrea,Mh.Mod2)/60.00))) TieHor" & vbCrLf & _
                             "From MovHis00 Mh" & vbCrLf & _
                             "Inner Join" & vbCrLf & _
                             "(" & vbCrLf & _
                             "  Select Mca.CodDoc, Mca.NumDoc" & vbCrLf & _
                             "  From MovCtaArt Mca" & vbCrLf & _
                             "  Where Mca.CodDocPad = '" & Trim(ICodDoc) & "' And Mca.NumDocPad = '" & INumDoc & "' And Mca.CodEmp = '" & ICodEmp & "' And Mca.CodDoc In ('Os01')" & vbCrLf & _
                             "  Group By Mca.CodDoc,Mca.NumDoc" & vbCrLf & _
                             ") Tmp On Tmp.CodDoc = Mh.CodDoc And Tmp.NumDoc = Mh.NumDoc" & vbCrLf & _
                             "Inner Join Anexo AnR On AnR.CodAne = Mh.CodRes" & vbCrLf & _
                             "Where Mh.CodEmp = '" & ICodEmp & "' And Mh.TipEve = Mh.CodDoc + 'CAD'" & " And Mh.FecCrea >= '" & sFecIni & "' And Mh.FecCrea <= '" & sFecFin & "'" & vbCrLf & _
                             "Group By AnR.CodAne, AnR.NomAne, Mh.FecCrea, Mh.Mod2, MemHis" & vbCrLf & _
                             "Order By Mh.FecCrea", , adCmdText)
     
    If Not rsCab.EOF Then
    
        If rsHis.EOF Then Call MsgBox("No existen marcas en el periodo seleccionado. Seleccione otro perido.", vbInformation, "Mensaje al Usuario"): GoTo Finalizar
            
        Dim prt As Printer
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
Finalizar:
    Screen.MousePointer = 0
    Call Close_RecordSet(rsCab)
    Call Close_RecordSet(rsHis)
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
Private Sub PrintData(Obj As Object)
    Dim dY As Integer, sTmp() As String, lTmp As Long, lCnt As Long, dX As Double, sCad As String, xCurY As Long, CurY As Long, lNxtLineA As Long, lNxtLineB As Long
    With Obj
        dX = 0: dY = 0
        
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Arial"
        Obj.Font.Size = 10
                            
        If TypeName(Obj) = "PictureBox" Then
            Obj.Height = 16837.79
            Obj.Width = 11905.5
        Else
            Obj.PaperSize = 9
        End If
        
        Call Caratula(Obj, rsCab, rsEmp)
        Call PrintPieCabecera(Obj)
        Call NewPage(Obj)
        Call PrintPieCabecera(Obj)
        
        Obj.Font.Bold = True
        Obj.Font.Size = 12
        
        Obj.CurrentX = 15
        Obj.CurrentY = 15
        Obj.Print "Resumen Informativo"
        
        Obj.Font.Bold = False
        Obj.Font.Size = 10
        
        sCad = "El presente informe detallaremos las horas trabajadas a " & Trim("" & rsCab.Fields("NomAne").Value) & " por el contrato celebrado correspondiente al periodo de " & Format(Trim("" & rsCab.Fields("FecDoc").Value), "MMMM") & " " & Format(Trim("" & rsCab.Fields("FecDoc").Value), "YYYY")
        xCurY = 25
        Call Cortar(sCad, 120, True, True, sTmp(), lTmp, False)
        For lCnt = 1 To lTmp
            Obj.CurrentX = 15
            Obj.CurrentY = xCurY
            Obj.Print sTmp(lCnt)
            xCurY = xCurY + 4
        Next
        
        xCurY = xCurY + 4
        
        Obj.Font.Bold = True
        Obj.Font.Size = 12
        
        Obj.CurrentX = 15
        Obj.CurrentY = xCurY
        Obj.Print "Resumen de horas por mes"
        
        xCurY = xCurY + 10
        
        Obj.Font.Bold = False
        Obj.Font.Size = 10
        
        Obj.CurrentX = 15
        Obj.CurrentY = xCurY
        Obj.Print UCase(Format(sFecIni, "MMMM")) & "            :          " & Format(CalcularTotal(rsHis), "0.00")
        
        Obj.Font.Bold = True
        Obj.Font.Size = 12
        
        xCurY = xCurY + 10
        
        Obj.CurrentX = 15
        Obj.CurrentY = xCurY
        Obj.Print "Detalle de horas informadas"
        
        Obj.Font.Bold = False
        Obj.Font.Size = 10
        
        If rsHis.RecordCount > 0 Then rsHis.MoveFirst
        xCurY = xCurY + 8
        
        Obj.Font.Bold = True
        Obj.Font.Size = 10
        
        Obj.CurrentX = 15
        Obj.CurrentY = xCurY
        Obj.Print "Trabajador"
        
        Obj.CurrentX = 55
        Obj.CurrentY = xCurY
        Obj.Print "Fecha inicio"
        
        Obj.CurrentX = 80
        Obj.CurrentY = xCurY
        Obj.Print "Descripción de la acción"
        
        Obj.CurrentX = 155
        Obj.CurrentY = xCurY
        Obj.Print "Horas"
        
        Obj.CurrentX = 170
        Obj.CurrentY = xCurY
        Obj.Print "Atención"
        
        xCurY = xCurY + 8
        
        Obj.Font.Bold = False
        Obj.Font.Size = 10
        
        CurY = xCurY
        Do Until rsHis.EOF
                        
            'Nombre del Trabajador
            
            Obj.CurrentX = 15
            Obj.CurrentY = CurY
            Obj.Print Mid(Trim("" & rsHis.Fields("NomRes").Value), 1, 20)
            
            'Fecha de Inicio
            Obj.CurrentX = 55
            Obj.CurrentY = CurY
            Obj.Print Format(Trim("" & rsHis.Fields("FecIni").Value), "dd/mm/yyyy")
            
            'Horas
            Obj.CurrentX = 162 - .TextWidth(Format(Trim("" & rsHis.Fields("TieHor").Value), "0.00"))
            Obj.CurrentY = CurY
            Obj.Print Format(Trim("" & rsHis.Fields("TieHor").Value), "0.00")
            
            'Atencion
            Obj.CurrentX = 170
            Obj.CurrentY = CurY
            Obj.Print Mid(Trim("" & rsCab.Fields("NomAne").Value), 1, 10)
            
            'Detalle
            sCad = Trim("" & rsHis.Fields("Detalle").Value)
            Call Cortar(sCad, 45, True, True, sTmp(), lTmp, False)
            For lCnt = 1 To lTmp
                Obj.CurrentX = 80
                Obj.CurrentY = CurY
                Obj.Print sTmp(lCnt)
                CurY = CurY + 4
                Call CheckPage(lNxtLineB, Obj)
            Next
            
            CurY = CurY + 4
            Call CheckPage(CurY, Obj)
            
            rsHis.MoveNext
        Loop
        
    End With
End Sub

Private Function CalcularTotal(ByRef oRs As ADODB.Recordset) As Double
    Dim lCnt As Long, dTotal As Double
    If oRs.RecordCount > 0 Then oRs.MoveFirst
    Do Until oRs.EOF
        dTotal = dTotal + Val(Trim("" & oRs.Fields("TieHor").Value))
        oRs.MoveNext
    Loop
    CalcularTotal = dTotal
End Function

Private Function GetValorScalar(ByVal sCampo As String, Optional ByVal sTabla As String = "", Optional ByVal sCondicion As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = ICnx.Execute("Select " & sCampo & IIf(Trim(sTabla) <> "", " From " & sTabla, "") & IIf(Trim(sCondicion) <> "", " Where " & sCondicion, ""), , adCmdText)
    If Not rsTmp.EOF Then
        GetValorScalar = Trim("" & rsTmp.Fields(0).Value)
    End If
    Call Close_RecordSet(rsTmp)
End Function

Private Sub NewPage(ByRef oObjImp As Object, Optional ByVal CurY As Long = 0)
    If TypeName(oObjImp) = "PictureBox" Then
        pvWindow.NewPage
        Set oObjImp = pvWindow.ActivePage
        If CurY > 0 Then oObjImp.CurrentY = CurY
    Else
        oObjImp.NewPage
        If CurY > 0 Then oObjImp.CurrentY = CurY
    End If
End Sub

Private Sub CheckPage(ByRef CurY As Long, ByRef Obj As Object)
    If CurY > 260 Then
        CurY = 15
        Call NewPage(Obj, CurY)
        Call PrintPieCabecera(Obj)
    End If
End Sub

Private Sub PrintPieCabecera(ByRef Obj As Object)
    Obj.FontSize = 10
    Obj.Font.Bold = False
    
    Obj.CurrentX = 10
    Obj.CurrentY = 5
    Obj.Print "INFORME DE HORAS - SIDIGE- ERP"
    
    Obj.Line (10, 10)-(200, 10)
    
    Obj.Line (10, 283)-(200, 283)
    
    Obj.CurrentX = 10
    Obj.CurrentY = 285
    Obj.Print "INTUITIVE SYSTEMS S.A.C"
    
    Obj.FontSize = 10
    Obj.Font.Bold = False
    
End Sub

Private Sub Caratula(ByRef oObjImp As Object, ByRef rsCab As ADODB.Recordset, ByRef rsEmp As ADODB.Recordset)
    Dim oStd As StdPicture
    With oObjImp
        
        Set oStd = GetLogo(Trim("" & rsEmp.Fields("RutLog").Value))
        If Not oStd Is Nothing Then
            Call oObjImp.PaintPicture(oStd, 85, 30, 30, 25)
        End If
        
        oObjImp.FontSize = 16
        
        oObjImp.CurrentX = 100 - (.TextWidth(Trim("" & rsEmp.Fields("NomEmp").Value)) / 2)
        oObjImp.CurrentY = 70
        oObjImp.Print Trim("" & rsEmp.Fields("NomEmp").Value)
        
        oObjImp.FontSize = 20
        oObjImp.Font.Bold = True
        
        oObjImp.CurrentX = 100 - (.TextWidth("INFORME DE HORAS EJECUTADAS") / 2)
        oObjImp.CurrentY = 120
        oObjImp.Print "INFORME DE HORAS EJECUTADAS"

        oObjImp.CurrentX = 100 - (.TextWidth("(SIDIGE ERP)") / 2)
        oObjImp.CurrentY = 130
        oObjImp.Print "(SIDIGE ERP)"

        oObjImp.CurrentX = 100 - (.TextWidth("CLIENTE") / 2)
        oObjImp.CurrentY = 150
        oObjImp.Print "CLIENTE"

        oObjImp.CurrentX = 100 - (.TextWidth(Trim("" & rsCab.Fields("NomAne").Value)) / 2)
        oObjImp.CurrentY = 160
        oObjImp.Print Trim("" & rsCab.Fields("NomAne").Value)

        oObjImp.CurrentX = 100 - (.TextWidth("MES") / 2)
        oObjImp.CurrentY = 180
        oObjImp.Print "MES"

        oObjImp.CurrentX = 100 - (.TextWidth(Format(sFecIni, "MMMM") & " " & Format(sFecIni, "YYYY")) / 2)
        oObjImp.CurrentY = 190
        oObjImp.Print Format(sFecIni, "MMMM") & " " & Format(sFecIni, "YYYY")
        
        oObjImp.Font.Size = 9
        oObjImp.Font.Bold = False

    End With
End Sub
