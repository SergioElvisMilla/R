VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpEti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bPrint As Boolean
Private ICodDoc As String
Private INumDoc As String
Private IProImp As String
Private ICodEmp As String
Private ICodGru As String
Private ITitFrm As String
Private IMovCab As String
Private IMovEdi As String
Private IMovDet As String
Private sDevice As String
Private sParam1 As String
Private INivel As Byte
Private ICnx As ADODB.Connection
Private rsCab As ADODB.Recordset
Private RsDet As ADODB.Recordset
Private sRndRegNeg001 As String 'CanTot
Private sRndRegNeg002 As String 'PreVta
Private sRndRegNeg003 As String 'CanTot * PreVta
Private sRndRegNeg004 As String 'Descuento
Private sRndRegNeg005 As String 'Impuestos
Private sRndTot As String 'Impuestos
Private oSysAdmin As Object
Private pvWindow As Object
Private Const sConst As Double = 0.01455
Property Get Printed() As Boolean
    Printed = bPrint
End Property
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
    IMovEdi = Valor
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
    Dim sCurDevice As String
    Screen.MousePointer = 11: bPrint = True
    Set rsCab = ICnx.Execute("Select Mc.CodDoc, Mc.NumDoc, Mc.FecDoc" & vbCrLf & _
                             "From " & IMovCab & " Mc" & vbCrLf & _
                             "Where Mc.CodDoc = '" & Trim(ICodDoc) & "' And Mc.NumDoc='" & INumDoc & "'", , adCmdText)
    
    Call Default_Reglas_Negocio
    If Not rsCab.EOF Then Call Reglas_Negocio(Trim$("" & rsCab.Fields("FecDoc").Value))
     
    Set RsDet = ICnx.Execute("Select Md.CodDoc, Md.NumDoc, Md.NumIte, Md.CanTot, Case When Ar.Imp01 <> 0 Then (1 + (Ar.Imp01/100)) Else 1 End * Case When Mc.xTipMon='MN' Then Ar.PrecMn Else Ar.PrecMe End PreVta," & vbCrLf & _
                             "Mdp.CodSub01 CodCat,Mdp.CodSub12 CodTemp,Mdp.CodSub13 CodLote,(Rtrim(Mdp.CodSub01)+' '+Rtrim(Mdp.CodSub12)+' '+Rtrim(Mdp.CodSub13))as CombX1  ,Mdp.NomSubArt13 Lote, Mdp.NomSubArt02 SubCategoria, Mdp.NomSubArt05 Modelo, Mdp.NomSubArt07 Color, Mdp.NomSubArt06 Talla, Mdp.NomSubArt20 Id, Mtd.DesAux SimMon" & vbCrLf & _
                             "From " & IMovDet & " Md" & vbCrLf & _
                             "Inner Join Articulo Ar On Ar.CodEmp='" & Trim(ICodEmp) & "' And Ar.CodSubAlm=Md.CodSubAlm And Ar.CodArt=Md.CodArt" & vbCrLf & _
                             "Inner Join " & IMovCab & " Mc On Mc.CodDoc=Md.CodDoc And Mc.NumDoc=Md.NumDoc" & vbCrLf & _
                             "Inner Join MovDetPro Mdp On Mdp.CodDoc=Md.CodDoc And Mdp.NumDoc=Md.NumDoc And Mdp.NumIte=Md.NumIte" & vbCrLf & _
                             "Inner Join MaeTabDet Mtd On Mtd.CodTab='xTipMon' And Mtd.CodIte=Ar.xTipMon And Mtd.CodFil='A'" & vbCrLf & _
                             "Where Md.CodDoc = '" & Trim(ICodDoc) & "' And Md.NumDoc='" & INumDoc & "' And IsNull(Md.NumIte1,'')='' And Md.CodSubAlm <> 'OT'" & vbCrLf & _
                             "Order By Convert(Int,Md.NumIte) Asc", , adCmdText)
    If Not rsCab.EOF Then
        Dim prt As Printer
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
        'Call pvWindow.Cerrar
        Set pvWindow = Nothing
        PrintDoc = True
    Else
        Screen.MousePointer = 0
        Call MsgBox("No se pudo encontrar el documento " & Trim(ICodDoc) & "-" & Trim(INumDoc) & ".", vbExclamation, "Mensaje al Usuario")
    End If
    Call Close_RecordSet(rsCab)
    Call Close_RecordSet(RsDet)
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
Private Sub PreviewData(ByRef Obj As Object)
   Dim dX As Long, dY As Long, dCurX As Long, dCurY As Long, lCnt As Long, dPreVta As Double, lNumReg As Long, lNumAct As Long, sCodBar As String
   Dim dxAj As Long
   Dim lMaxCol As Integer
   Dim lMaxRow As Integer
   Dim lColAct As Long
   'With Obj
        dX = 0: dY = 0: lMaxCol = 3: lMaxRow = 1: dxAj = 0
        Obj.Font.Size = 6
        Obj.ScaleMode = vbMillimeters
        Obj.Font.Name = "Arial"
        Obj.Font.Bold = True
        Obj.Height = 2400
        Obj.Width = 5400
        If (TypeName(Obj) = "PictureBox") Then
            Printer.Font.Size = 6
            Printer.ScaleMode = vbMillimeters
            Printer.Font.Name = "Arial"
            Printer.Font.Bold = True
            Obj.Height = 2400
            Obj.Width = 5400
        End If


        If RsDet.EOF Then RsDet.MoveFirst
        Do Until RsDet.EOF
        
            For lCnt = 1 To CLng(RsDet.Fields("CanTot").Value)
                
                
                lNumAct = lNumAct + 1: lColAct = lColAct + 1
                
                If (TypeName(Obj) = "PictureBox") Then
                Obj.Font.Name = "Arial": Obj.Font.Size = 6: Obj.ScaleMode = vbMillimeters: Obj.Font.Bold = True
                Else
                Printer.Font.Name = "Arial": Printer.Font.Size = 6: Printer.ScaleMode = vbMillimeters: Printer.Font.Bold = True
                End If
                
                dCurY = 2 + dY
                dCurX = 2 + dX
                
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("CombX1").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("CombX1").Value)
                dCurY = dCurY + 4
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("SubCategoria").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("SubCategoria").Value)
                dCurY = dCurY + 4
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("Modelo").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("Modelo").Value)
                dCurY = dCurY + 4
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("Color").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("Color").Value)
                dCurY = dCurY + 4
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth("T/. " & Trim("" & RsDet.Fields("Talla").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print "T/. " & Trim("" & RsDet.Fields("Talla").Value)
                dCurY = dCurY + 4
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("SimMon").Value) & " " & Format(RsDet.Fields("PreVta").Value, sRndTot)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("SimMon").Value) & " " & Format(RsDet.Fields("PreVta").Value, sRndTot)
                dCurY = dCurY + 4
                
                Obj.Font.Name = "BarCode 128": Obj.Font.Size = 20: Obj.Font.Bold = False
                
                If (TypeName(Obj) = "PictureBox") Then
                Printer.Font.Name = "BarCode 128": Printer.Font.Size = 20: Printer.Font.Bold = False
                Else
                Obj.Font.Name = "BarCode 128": Obj.Font.Size = 20: Obj.Font.Bold = False
                End If
                
                sCodBar = Trim("" & RsDet.Fields("Id").Value)
                sCodBar = PatronCodigoBarra128B(sCodBar)
                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(sCodBar) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print sCodBar
                
                dCurY = dCurY + 8
                
                Obj.Font.Name = "Arial": Obj.Font.Size = 6: Obj.ScaleMode = vbMillimeters: Obj.Font.Bold = True
                If (TypeName(Obj) = "PictureBox") Then
                Obj.Font.Name = "Arial": Obj.Font.Size = 6: Obj.ScaleMode = vbMillimeters: Obj.Font.Bold = True
                Else
                Printer.Font.Name = "Arial": Printer.Font.Size = 6: Printer.ScaleMode = vbMillimeters: Printer.Font.Bold = True
                End If

                Obj.CurrentX = ((Obj.Width / (2 * lMaxCol)) * sConst) - (Obj.TextWidth(Trim("" & RsDet.Fields("Id").Value)) / 2) + dX + dxAj
                Obj.CurrentY = dCurY
                Obj.Print Trim("" & RsDet.Fields("Id").Value)
                
                dX = dX + 30
                
                If lNumAct = 1 Then
                dX = dX + 1.5
                End If
                If lNumAct = 2 Then
                dX = dX + 3.25
                End If

                If lNumAct >= (lMaxCol * lMaxRow) Then
                    lNumAct = 0: dX = 0: dY = 0: lColAct = 0
                    If TypeName(Obj) = "PictureBox" Then
                        Call pvWindow.NewPage
                        Set Obj = pvWindow.ActivePage
                    Else
                        Obj.NewPage
                    End If
                ElseIf lColAct >= lMaxCol Then
                    dX = 0: lColAct = 0: dY = dY + 40
                End If
            Next
            RsDet.MoveNext
        Loop
  'End With
End Sub

Private Sub Default_Reglas_Negocio()
    If Trim(sRndRegNeg001) = "" Then sRndRegNeg001 = "#,0.00"
    If Trim(sRndRegNeg002) = "" Then sRndRegNeg002 = "#,0.00"
    If Trim(sRndRegNeg003) = "" Then sRndRegNeg003 = "#,0.00"
    If Trim(sRndRegNeg004) = "" Then sRndRegNeg004 = "#,0.00"
    If Trim(sRndRegNeg005) = "" Then sRndRegNeg005 = "#,0.00"
    If Trim(sRndTot) = "" Then sRndTot = "#,0.00"
End Sub

Private Sub Reglas_Negocio(ByVal sFecDoc As String)
    Call oSysAdmin.Reglas_Negocios("Reg0000001", , sRndRegNeg001, , sFecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000002", , sRndRegNeg002, , sFecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000003", , sRndRegNeg003, , sFecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000004", , sRndRegNeg004, , sFecDoc)
    Call oSysAdmin.Reglas_Negocios("Reg0000005", , sRndRegNeg005, , sFecDoc)
End Sub



