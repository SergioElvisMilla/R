Attribute VB_Name = "modShared"
Option Explicit

Global Const gMaeTabIdImp As String = "Cfg_IdImp"

Public Enum EnumPaperSize
    DMPAPER_10X11 = 45                      '  10 x 11 pulg.
    DMPAPER_10X14 = 16                      '  10x14 pulg.
    DMPAPER_11X17 = 17                      '  11x17 pulg.
    DMPAPER_15X11 = 46                      '  15 x 11 pulg.
    DMPAPER_9X11 = 44                       '  9 x 11 pulg.
    DMPAPER_A_PLUS = 57                     '  SuperA/SuperA/A4 227 x 356 mm
    DMPAPER_A2 = 66                         '  A2 420 x 594 mm
    DMPAPER_A3 = 8                          '  A3 297 x 420 mm
    DMPAPER_A3_EXTRA = 63                   '  A3 Extra 322 x 445 mm
    DMPAPER_A3_EXTRA_TRANSVERSE = 68        '  A3 Extra Transversal 322 x 445 mm
    DMPAPER_A3_TRANSVERSE = 67              '  A3 Transversal 297 x 420 mm
    DMPAPER_A4 = 9                          '  A4 210 x 297 mm
    DMPAPER_A4_EXTRA = 53                   '  A4 Extra 9.27 x 12.69 pulg.
    DMPAPER_A4_PLUS = 60                    '  A4 Plus 210 x 330 mm
    DMPAPER_A4_TRANSVERSE = 55              '  A4 Transversal 210 x 297 mm
    DMPAPER_A4SMALL = 10                    '  A4 Pequeño 210 x 297 mm
    DMPAPER_A5 = 11                         '  A5 148 x 210 mm
    DMPAPER_A5_EXTRA = 64                   '  A5 Extra 174 x 235 mm
    DMPAPER_A5_TRANSVERSE = 61              '  A5 Transversal 148 x 210 mm
    DMPAPER_B_PLUS = 58                     '  SuperB/SuperB/A3 305 x 487 mm
    DMPAPER_B4 = 12                         '  B4 250 x 354
    DMPAPER_B5 = 13                         '  B5 182 x 257 mm
    DMPAPER_B5_EXTRA = 65                   '  B5 (ISO) Extra 201 x 276 mm
    DMPAPER_B5_TRANSVERSE = 62              '  B5 (JIS) Transversal 182 x 257 mm
    DMPAPER_CSHEET = 24                     '  Hoja tamaño C
    DMPAPER_DSHEET = 25                     '  Hoja tamaño D
    DMPAPER_ENV_10 = 20                     '  Sobre Nº 10 4 1/8 x 9 1/2
    DMPAPER_ENV_11 = 21                     '  Sobre Nº 11 4 1/2 x 10 3/8
    DMPAPER_ENV_12 = 22                     '  Sobre Nº 12 4 \276 x 11
    DMPAPER_ENV_14 = 23                     '  Sobre Nº 14 5 x 11 1/2
    DMPAPER_ENV_9 = 19                      '  Sobre Nº 9 3 7/8 x 8 7/8
    DMPAPER_ENV_B4 = 33                     '  Sobre B4  250 x 353 mm
    DMPAPER_ENV_B5 = 34                     '  Sobre B5  176 x 250 mm
    DMPAPER_ENV_B6 = 35                     '  Sobre B6  176 x 125 mm
    DMPAPER_ENV_C3 = 29                     '  Sobre C3  324 x 458 mm
    DMPAPER_ENV_C4 = 30                     '  Sobre C4  229 x 324 mm
    DMPAPER_ENV_C5 = 28                     '  Sobre C5 162 x 229 mm
    DMPAPER_ENV_C6 = 31                     '  Sobre C6  114 x 162 mm
    DMPAPER_ENV_C65 = 32                    '  Sobre C65 114 x 229 mm
    DMPAPER_ENV_DL = 27                     '  Sobre DL 110 x 220mm
    DMPAPER_ENV_INVITE = 47                 '  Sobre Invitación 220 x 220 mm
    DMPAPER_ENV_ITALY = 36                  '  Sobre 110 x 230 mm
    DMPAPER_ENV_MONARCH = 37                '  Sobre monarca 3.875 x 7.5 pulg.
    DMPAPER_ENV_PERSONAL = 38               '  Sobre 6 3/4  3 5/8 x 6 1/2 pulg.
    DMPAPER_ESHEET = 26                     '  Hoja tamaño E
    DMPAPER_EXECUTIVE = 7                   '  Ejecutivo 7 1/4 x 10 1/2 pulg.
    DMPAPER_FANFOLD_LGL_GERMAN = 41         '  Continuo alemán oficio 8 1/2 x 13 pulg.
    DMPAPER_FANFOLD_STD_GERMAN = 40         '  Continuo alemán estándar 8 1/2 x 12 pulg.
    DMPAPER_FANFOLD_US = 39                 '  Continuo USA estándar 14 7/8 x 11 pulg.
    DMPAPER_FIRST = 1                       '  Carta 8 1/2 x 11 pulg.
    DMPAPER_FOLIO = 14                      '  Folio 8 1/2 x 13 pulg.
    DMPAPER_ISO_B4 = 42                     '  B4 (ISO) 250 x 353 mm
    DMPAPER_JAPANESE_POSTCARD = 43          '  Tarjeta japonesa 100 x 148 mm
    DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN
    DMPAPER_LEDGER = 4                      '  Doble carta 17 x 11 pulg.
    DMPAPER_LEGAL = 5                       '  Oficio 8 1/2 x 14 pulg.
    DMPAPER_LEGAL_EXTRA = 51                '  Oficio Extra 9 \275 x 15 pulg.
    DMPAPER_LETTER = 1
    DMPAPER_LETTER_EXTRA = 50               '  Carta Extra 9 \275 x 12 pulg.
    DMPAPER_LETTER_EXTRA_TRANSVERSE = 56    '  Carta Extra Transversal 9\275 x 12 pulg.
    DMPAPER_LETTER_PLUS = 59                '  Carta Plus 8.5 x 12.69 pulg.
    DMPAPER_LETTER_TRANSVERSE = 54          '  Carta Transversal 8 \275 x 11 pulg.
    DMPAPER_LETTERSMALL = 2                 '  Carta pequeña 8 1/2 x 11 pulg.
    DMPAPER_NOTE = 18                       '  Nota 8 1/2 x 11 pulg.
    DMPAPER_QUARTO = 15                     '  Cuarto 215 x 275 mm
    DMPAPER_RESERVED_48 = 48                '  RESERVADO--NO LO USE
    DMPAPER_RESERVED_49 = 49                '  RESERVADO--NO LO USE
    DMPAPER_STATEMENT = 6                   '  Estamento 5 1/2 x 8 1/2 pulg.
    DMPAPER_TABLOID = 3                     '  Tabloide 11 x 17 pulg.
    DMPAPER_TABLOID_EXTRA = 52              '  Tabloide Extra 11.69 x 18 pulg.
End Enum
Function GetPrinter(Device As String) As Printer
    Dim prt As Printer
    For Each prt In Printers
        If prt.DeviceName = Device Then
            Set GetPrinter = prt
            Exit For
        End If
    Next
    Set prt = Nothing
End Function
Function GetLogo(Ruta As String) As StdPicture
    On Error Resume Next
    If Ruta <> "" Then
        Set GetLogo = LoadPicture(Ruta)
    End If
    Err.Clear
End Function
Public Sub Set_PaperSize(ByRef oObject As Object, ByVal iPaperSize As Integer)
On Error GoTo Solucion
    oObject.PaperSize = iPaperSize
Exit Sub
Solucion:
    Err.Clear
End Sub
Public Function Buscar_cadena(ByVal Cdn As String, ByVal bsqCdn As String, ByVal Pst1 As Integer, ByRef Pst2 As Integer) As String
    Pst1 = InStr(Pst2, Cdn, bsqCdn, 3): If Pst1 <> 0 Then Buscar_cadena = Mid(Cdn, Pst2, Pst1 - Pst2): Pst2 = Pst1 + Len(bsqCdn)
End Function
Public Sub Cortar(ByVal Cdn As String, ByVal Mx As Long, ByVal bCrLf As Boolean, ByVal bJust As Boolean, ByRef sLn() As String, ByRef lLn As Long, ByVal bPrinter As Boolean, Optional lFirstLine As Long = 0)
    Dim lIni As Long, StmP As String, sTmp2 As String, sTmp3 As String, lCnt As Long, lMaxWidth As Long
    Cdn = Trim$("" & Cdn): ReDim sLn(1 To 1) As String: sLn(1) = "": lLn = 0: lMaxWidth = IIf((lFirstLine > 0), lFirstLine, Mx)
    Do
        If bCrLf Then
            lIni = InStr(1, Cdn, vbCrLf)
            If lIni = 0 Then
                StmP = Trim$(Cdn): Cdn = ""
            Else
                StmP = Trim$(Mid(Cdn, 1, lIni - 1))
                Cdn = Mid(Cdn, lIni + 2)
            End If
        Else
            StmP = Replace(Trim$(Cdn), vbCrLf, " ", 1, , vbTextCompare): Cdn = ""
        End If
        Do
            If bPrinter Then
                If (Len(StmP) > 425) Then
                    GoTo MaxWidthLen
                ElseIf Printer.TextWidth(StmP) < lMaxWidth Then
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = StmP: StmP = "": lMaxWidth = Mx
                Else
MaxWidthLen:
                    lIni = 0
                    Do
                        lCnt = lIni: lIni = InStr((lIni + 1), Trim$(StmP) & " ", " ", vbTextCompare)
                        If lIni <= 0 Then lCnt = 0: Exit Do
                    Loop Until (Printer.TextWidth(Mid(Trim$(StmP) & " ", 1, lIni)) >= lMaxWidth)
                    If (lCnt > 0) Then
                        lIni = lCnt
                        sTmp2 = Trim$(Mid(Trim$(StmP) & " ", 1, lIni))
                        If bJust Then
                            For lCnt = 2 To lMaxWidth
                                If Printer.TextWidth(Replace(sTmp2, " ", Space(lCnt + 1), 1, , vbTextCompare)) > lMaxWidth Then
                                    sTmp2 = Replace(sTmp2, " ", Space(lCnt - 1), 1, , vbTextCompare): Exit For
                                End If
                            Next lCnt
                        End If
                        StmP = Trim$(Mid(StmP, lIni))
                    Else
                        sTmp2 = Trim$(StmP): StmP = ""
                    End If
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = sTmp2: lMaxWidth = Mx
                End If
            Else
                If Len(StmP) < lMaxWidth Then
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = StmP: StmP = "": lMaxWidth = Mx
                Else
                    sTmp2 = Mid(Trim$(StmP) & " ", 1, lMaxWidth)
                    lIni = InStrRev(sTmp2, " ", lMaxWidth, vbTextCompare)
                    If lIni = 0 Then
                        lIni = InStr(1, Trim$(StmP) & " ", " ", vbTextCompare)
                        sTmp2 = Mid(Trim$(Mid(Trim$(StmP) & " ", 1, lIni)), 1, lMaxWidth)
                        StmP = Trim$(Mid(StmP, Len(sTmp2)))
'                        sTmp = Trim$(Mid(sTmp, 1, lIni - 1))
                    Else
                        sTmp2 = Trim$(Mid(Trim$(StmP) & " ", 1, lIni - 1))
                        If bJust Then
                            For lCnt = 2 To lMaxWidth
                                If Len(Replace(sTmp2, " ", Space(lCnt), 1, , vbTextCompare)) > lMaxWidth Then sTmp2 = Replace(sTmp2, " ", Space(lCnt - 1), 1, , vbTextCompare): Exit For
                            Next lCnt
                        End If
                        StmP = Trim$(Mid(StmP, lIni + 1))
                    End If
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = sTmp2: lMaxWidth = Mx
                End If
            End If
        Loop Until StmP = ""
    Loop Until Cdn = ""
End Sub
Public Sub Cortar2(ByVal Cdn As String, ByVal Mx As Long, ByVal bCrLf As Boolean, ByVal bJust As Boolean, ByRef sLn() As String, ByRef lLn As Long, ByVal bPrinter As Boolean, Optional lFirstLine As Long = 0)
    Dim lIni As Long, StmP As String, sTmp2 As String, sTmp3 As String, lCnt As Long, lMaxWidth As Long
    Cdn = Trim$("" & Cdn): ReDim sLn(1 To 1) As String: sLn(1) = "": lLn = 0: lMaxWidth = IIf((lFirstLine > 0), lFirstLine, Mx)
    Do
        If bCrLf Then
            lIni = InStr(1, Cdn, vbCrLf)
            If lIni = 0 Then
                StmP = Trim$(Cdn): Cdn = ""
            Else
                StmP = Trim$(Mid(Cdn, 1, lIni - 1))
                Cdn = Mid(Cdn, lIni + 2)
            End If
        Else
            StmP = Replace(Trim$(Cdn), vbCrLf, " ", 1, , vbTextCompare): Cdn = ""
        End If
        Do
            If bPrinter Then
                If (Len(StmP) > 425) Then
                    GoTo MaxWidthLen
                ElseIf Printer.TextWidth(StmP) < lMaxWidth Then
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = StmP: StmP = "": lMaxWidth = Mx
                Else
MaxWidthLen:
                    lIni = 0
                    Do
                        lCnt = lIni: lIni = InStr((lIni + 1), Trim$(StmP) & " ", " ", vbTextCompare)
                        If lIni <= 0 Then lCnt = 0: Exit Do
                    Loop Until (Printer.TextWidth(Mid(Trim$(StmP) & " ", 1, lIni)) >= lMaxWidth)
                    If (lCnt > 0) Then
                        lIni = lCnt
                        sTmp2 = Trim$(Mid(Trim$(StmP) & " ", 1, lIni))
                        If bJust Then
                            For lCnt = 2 To lMaxWidth
                                If Printer.TextWidth(Replace(sTmp2, " ", Space(lCnt + 1), 1, , vbTextCompare)) > lMaxWidth Then
                                    sTmp2 = Replace(sTmp2, " ", Space(lCnt - 1), 1, , vbTextCompare): Exit For
                                End If
                            Next lCnt
                        End If
                        StmP = Trim$(Mid(StmP, lIni))
                    Else
                        sTmp2 = Trim$(StmP): StmP = ""
                    End If
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = sTmp2: lMaxWidth = Mx
                End If
            Else
                If Len(StmP) < lMaxWidth Then
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = StmP: StmP = "": lMaxWidth = Mx
                Else
                    sTmp2 = Mid(Trim$(StmP) & " ", 1, lMaxWidth)
                    lIni = InStrRev(sTmp2, " ", lMaxWidth, vbTextCompare)
                    If lIni = 0 Then
                        lIni = InStr(1, Trim$(StmP) & " ", " ", vbTextCompare)
                        sTmp2 = Mid(Trim$(Mid(Trim$(StmP) & " ", 1, lIni)), 1, lMaxWidth)
                        StmP = Trim$(Mid(StmP, Len(sTmp2)))
'                        sTmp = Trim$(Mid(sTmp, 1, lIni - 1))
                    Else
                        sTmp2 = Trim$(Mid(Trim$(StmP) & " ", 1, lIni - 1))
                        If bJust Then
                            For lCnt = 2 To lMaxWidth
                                If Len(Replace(sTmp2, " ", Space(lCnt), 1, , vbTextCompare)) > lMaxWidth Then sTmp2 = Replace(sTmp2, " ", Space(lCnt - 1), 1, , vbTextCompare): Exit For
                            Next lCnt
                        End If
                        StmP = Trim$(Mid(StmP, lIni)) '+1
                    End If
                    lLn = lLn + 1: ReDim Preserve sLn(1 To lLn) As String: sLn(lLn) = sTmp2: lMaxWidth = Mx
                End If
            End If
        Loop Until StmP = ""
    Loop Until Cdn = ""
End Sub
Public Function PatronCodigoBarra128B(ByVal Codigo As String) As String
    Dim sCarIni     As String
    Dim sCarFin     As String
    Dim sCodBar     As String
    Dim nCodBar     As Long
    Dim nLenCod     As Long
    Dim nCaracter   As Long
    Dim nAscii      As Byte
    Dim sDigVer     As String
    
    sCarIni = Chr(104 + 32)
    sCarFin = Chr(106 + 32)
    nCodBar = Asc(sCarIni) - 32
    sCodBar = Codigo
    nLenCod = Len(sCodBar)
    
    For nCaracter = 1 To nLenCod
        nAscii = Asc(Mid(sCodBar, nCaracter, 1)) - 32
    
        If Not nAscii >= 0 And nAscii <= 99 Then
            sCodBar = Replace(sCodBar, nCaracter, 1, Chr(32))
            nAscii = Asc(Mid(sCodBar, nCaracter, 1)) - 32
        End If
        nCodBar = nCodBar + (nAscii * nCaracter)
    Next
    
    sDigVer = Chr((nCodBar Mod 103) + 32)
    sCodBar = sCarIni + sCodBar + sDigVer + sCarFin

    sCodBar = Replace(sCodBar, Chr(32), Chr(232))
    sCodBar = Replace(sCodBar, Chr(127), Chr(192))
    sCodBar = Replace(sCodBar, Chr(128), Chr(193))
    
    PatronCodigoBarra128B = sCodBar
End Function
Function Imp_Comentario(Cx As Long, Cy As Long, Comentario As String, Ob As Object, Optional ByRef objFormato As Object)
    Dim ini As Integer
    Dim Cad As String
    Dim Pos As Long
    Cad = Comentario
    ini = 1
    Ob.CurrentY = Cy
    If InStr(1, Cad, vbCrLf) = 0 Then
        Pos = Len(Cad)
        Ob.Print Cad
    Else
        Pos = 0
    End If
    Do While ini < Len(Cad)
        Pos = InStr(ini + 1, Cad, vbCrLf) + 2
        If Pos = 2 Then
            Ob.CurrentX = Cx
            If Not objFormato Is Nothing Then
                objFormato.CheqPage Cx, Ob.TextHeight("0"), Ob
            End If
            Ob.Print Mid(Cad, ini)
            Exit Function
        End If
        Ob.CurrentX = Cx
        Ob.Print Mid(Cad, ini, Pos - ini - 2)
        ini = Pos '+ 1
    Loop
End Function

Public Function RngExc(iCol As Integer) As String
    Dim sCodExc As String, dRst As Double
    If iCol > 25 Then
        dRst = iCol / 25
        If Round(dRst, 0) <> dRst Then dRst = Round(dRst - 0.5, 0)
        sCodExc = Chr(64 + CInt(dRst)) & Chr(64 + CInt((iCol - (CInt(dRst) * 25))))
    Else
        sCodExc = Chr(65 + iCol)
    End If
    RngExc = sCodExc
End Function

Public Sub DropTable(ByRef oCnx As ADODB.Connection, ByVal sTab As String)
    On Error GoTo Solucion
    Call oCnx.Execute("Drop Table " & sTab, , adCmdText)
    Exit Sub
Solucion:
    Err.Clear
End Sub

Public Function GetTempTable(ByRef oCnx As ADODB.Connection, Numero As Long) As String
    Dim rsTmp As ADODB.Recordset, iSpid As Integer
    Set rsTmp = oCnx.Execute("Select @@Spid", , adCmdText)
    If Not rsTmp.EOF Then iSpid = rsTmp.Fields(0).Value
    GetTempTable = "##Tmp_Spid_" & iSpid & "_" & Numero & "_" & Round(Rnd() * 10000, 0)
    Set rsTmp = oCnx.Execute("Select * From TempDb..SysObjects Where Name = '" & GetTempTable & "';", , adCmdText)
    If Not rsTmp.EOF Then GetTempTable = GetTempTable(oCnx, Numero)
    rsTmp.Close: Set rsTmp = Nothing
End Function
Public Sub Close_RecordSet(ByRef oRs As ADODB.Recordset)
    If Not oRs Is Nothing Then
        If oRs.State <> adStateClosed Then oRs.Close
        Set oRs = Nothing
    End If
End Sub
Public Sub Close_RecordSet1(ByRef oRd As ADODB.Recordset)
    If Not oRd Is Nothing Then
        If oRd.State <> adStateClosed Then oRd.Close
        Set oRd = Nothing
    End If
End Sub
Public Function CheckGroupBy(ByRef sFilter() As String, ByVal iColA As Integer, ByVal iColB As Integer, ByVal iColC As Integer) As Boolean
    If Trim$("" & sFilter(iColA, iColB, iColC)) = "1" Then CheckGroupBy = True
End Function

Public Function GetNomPc() As String
    Dim oTool As Object
    Set oTool = CreateObject("vbXTools.xSystemTools")
    GetNomPc = Trim$("" & oTool.ComputerName)
    Set oTool = Nothing
End Function

Public Function GetIdImp(ByRef oCnx As ADODB.Connection) As String
    Dim oRs As ADODB.Recordset, sNomPc As String
    sNomPc = GetNomPc
    Set oRs = oCnx.Execute("Select DesAux From MaeTabDet Where CodTab = '" & Trim(gMaeTabIdImp) & "' And CodFil = 'A' And DesIte = '" & UCase(Trim(sNomPc)) & "'", , adCmdText)
    If Not oRs.EOF Then
        GetIdImp = Trim("" & oRs.Fields("DesAux").Value)
    End If
    Call Close_RecordSet(oRs)
End Function
