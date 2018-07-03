VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProImpOcn"
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
Private IMovDet As String
Private IDevice As String
Private IParam1 As String
Private INivel As Byte
Private MPcx As ADODB.Connection
Private ICnx As ADODB.Connection
Private dSubTotal As Double
Private dDesc1 As Double
Private dDesc2 As Double
Private dDesc3 As Double
Private dDesc4 As Double
Private BVERCODBAR As Boolean
Private BVEROCN As Boolean


'ULTIMA MODIFCACION 24/08/2006
Public Property Set Connection(ByVal Cnn As ADODB.Connection)
    Set MPcx = Cnn
End Property
Public Property Let NumDoc(ByVal Cad As String)
    INumDoc = Cad
End Property
Public Property Let CodDoc(ByVal Cad As String)
    ICodDoc = Cad
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
Public Property Let RutImp(ByVal Cad As String)
    IDevice = Cad
End Property
Public Property Let Param1(ByVal Cad As String)
End Property
Public Property Let Nivel(ByVal Cad As Byte)
    INivel = Cad
End Property
Public Sub PrintShow()
    PrintDoc True
End Sub
Public Sub PrintDoc(Optional PreView As Boolean = False)
    Dim RsEmp As ADODB.Recordset
    Dim rsDes As ADODB.Recordset
    Dim RsCab As ADODB.Recordset
    Dim rsDes2 As ADODB.Recordset
    Dim rsCab2 As ADODB.Recordset
    Dim RsDetLar As ADODB.Recordset
    Dim RsPadre As ADODB.Recordset

    Dim campoB As String
    Dim Pos As Integer
    Dim iLonDet As Integer 'Longitud maxima por linea de Articulo
    Dim Valor As String
    Dim sCadAux As String
    
    Dim lPosI As Integer
    Dim lPosF As Integer
    Dim lLen As Integer
    Dim sCad As String
    Dim Igv As String
    Dim sTmp() As String, lTmp As Long, lCnt As Long
    Dim bVerFrm As Boolean
    
    Dim rpt As Object
    
    BVERCODBAR = True
    BVEROCN = False
    
    
    
'    Set oRsCfg = ICnx.Execute("Select * From MaeTabDet Where CodFil = 'A' And CodTab = '" & Trim(sParam1) & "';", , adCmdText)
'    Do Until oRsCfg.EOF
'        Select Case UCase(Trim("" & oRsCfg.Fields("CodIte").Value))
'            Case "VERFRM": bVerFrm = Mid(UCase(Trim("" & oRsCfg.Fields("DesAux").Value)), 1, 1) = "S"
'        End Select
'        oRsCfg.MoveNext
'    Loop
    
'    Dim frm As New frmOpcImp
'    If frm.ShowForm() Then
'            BVERCODBAR = frm.VerSello
'            BVEROCN = frm.VerPeriodo
'    End If
'    Call Unload(frm): Set frm = Nothing
'
    Screen.MousePointer = 11
    
    '=============================
    iLonDet = 55 'Longitud del detalle por linea
        
    
    'bImpDesArt1 = True
    'bImpRec = True
    'bImpMarDes = False
        
    Set RsEmp = MPcx.Execute("Select Top 1 CODEMP,NOMEMP,DIREMP,TELEMP,FAXEMP,BKUPEMP,RUCEMP,RUtLOG,EMAIL FROM EMPRESAS WHERE CODEMP='" & ICodEmp & "'", , adCmdText)
    If RsEmp.EOF And RsEmp.BOF Then
        Screen.MousePointer = 0
        MsgBox "No se puede localizar la empresa", vbExclamation, "Error"
        Exit Sub
    End If
    Set RsCab = MPcx.Execute("Select Top 1 isnull(cab.fecent02,'') as fecent02,isnull(cab.glosa,'')as glosa,cab.REFINT,MAE4.DESITE AS FORENV,MAE3.DESITE AS TIPVTA,ENT.TELANE AS TELENT,CON2.NOMCON AS NOMCONENT ,CAB.CODANE02,CAB.NOMANE02,CAB.NUMRF1,ANE.CELANE AS CELRES,ANE.TELANE AS TELRES,ANE.MAILANE  AS MAILRES,CAB.NUMRF1,CAB.NUMREA,CON.XCODCAR,MAE2.DESITE AS XDESCAR,con.TelCon as TelfCont,Con.CelCon as CellCon,Con.MailCon as Mail, Con.MailCon,Con.NomCon MailCon,Cli.CelAne TelCon,CAB.CODDOC,CAB.SUBTOT,CAB.NUMCON,CON.NOMCON,UBIG.DESCRIPCION AS XUBIGEO,CAB.DIRANE,CAB.DIRANE02,CFG.DESCRI,CFG.DESCRI + ' Aceptada'  AS NomDocAce, ANE.NOMANE AS NOMRES,CLI.UBIZON,UBI.NOMBRE AS NOMUBI,CAB.FECENT,CAB.CODDOC,CAB.CODANE,CAB.NUMDOC,CAB.FECDOC,CAB.NOMANE,CAB.CODVEN,CAB.TELANE,CLI.FAXANE FaxAne,CAB.IDEANE,CAB.TOTDOC,CAB.TOTIMP1,MAE.DESITE AS CONPAG,MAE1.DESITE AS FORPAG,CAB.NUMDIA,CAB.CODMEM,CAB.COMMEM,CASE CAB.XTIPMON WHEN 'MN' THEN 'S/.' ELSE 'US$' END AS SimMon, " & _
                             " CAB.SUBTOT - (CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) AS IMPNET, " & _
                             " CAB.TOTIMP1+CAB.TOTIMP2+CAB.TOTIMP3+CAB.TOTIMP4 AS TOTIMP, " & _
                             " CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4 AS TOTDSC, " & _
                             " CAB.SUBTOT,CAB.FECAPR,CAB.TOTIMP1,CAB.TOTIMP2,CAB.TOTIMP3,CAB.TOTIMP4,CAB.TOTDSC1,CAB.TOTDSC2,CAB.TOTDSC3,CAB.TOTDSC4,CAB.TOTDOC,CLI.ideane4 AS CodBri , " & _
                             " Case When SUBTOT = 0 Then 0 Else ((CAB.TOTDSC1+CAB.TOTDSC2+CAB.TOTDSC3+CAB.TOTDSC4) *100)/SUBTOT End AS TOTPORDESC " & _
                             " FROM " & IMovCab & "  AS CAB" & _
                             " Left Join MAETABDET AS MAE On MAE.CODTAB='CONPAG' And XCONPAG=MAE.CODITE" & _
                             " Left Join MAETABDET AS MAE1 On MAE1.CODTAB='FORPAG' And XFORPAG=MAE1.CODITE" & _
                             " Left Join Anexo AS ANE On CAB.CODRES=ANE.CODANE" & _
                             " Left Join CFGDOC00 AS CFG On CAB.CODDOC=CFG.CODDOC " & _
                             " Left Join Anexo CLI On  CAB.CODANE=CLI.CODANE " & _
                             " Left Join Anexo ENT On  CAB.CODANE02=ENT.CODANE " & _
                             " Left Join UBIZON UBI On CLI.UBIZON = UBI.UBIZON " & _
                             " Left Join UBIGEO UBIG On CLI.XUBIGEO = UBIG.UBIGEO " & _
                             " Left Join CONTACTOS CON On CAB.CODANE = CON.CODANE And CON.NUMCON = CAB.NUMCON " & _
                             " Left Join CONTACTOS CON2 On CAB.CODANE02 = CON2.CODANE And CON2.NUMCON = CAB.NUMCON02 " & _
                             " Left Join MAETABDET AS MAE2 On MAE2.CODTAB='XCODCAR' And CON.XCODCAR=MAE2.CODITE " & _
                             " Left Join MAETABDET AS MAE3 On MAE3.CODTAB='MOV-OC' And CAB.XTIPMOV=MAE3.CODITE   " & _
                             " Left Join MAETABDET AS MAE4 On MAE4.CODTAB='FORENV' And CAB.FORENV=MAE4.CODITE   " & _
                             " WHERE NUMDOC='" & INumDoc & "'", , adCmdText)
    
    If RsCab.EOF = False And RsCab.BOF = False Then
        Dim cls As Cls02
        Set cls = New Cls02
        dDesc1 = RsCab!TotDsc1
        dDesc2 = RsCab!TotDsc2
        dDesc3 = RsCab!TotDsc3
        dDesc4 = RsCab!TotDsc4
        cls.SetPrinterOrientation OrientPortrait ', DMPAPER_LETTER
        Set cls = Nothing
    
    Set rsDes = MPcx.Execute("Select MD.xtipuni as uni,Md.IMP001,MD.INGSAL,MD.NumIte,ISNULL(MD.NumIte1,'') as NumIte1,MD.CODART,MD.DesArt AS DesArt,MD.DesArt1,A.DetArt1 as DetArt1,MD.XTIPUNI,MD.PreVta,MAE.DESITE AS DESUNI,X04.NOMSUBART + ' ' + Md.DesArt AS Marca" & _
                            " ,AD.XSUBART17 AS CODGAR,X17.NOMSUBART AS GARANTIA,MD.XTIPUNI, " & _
                            " Case When Md.PreVta = 0 Then 0 Else  (Md.PreVta*MD.CanTot) - (((Md.PreVta*MD.CanTot)  * (((CanTot*PreVta) -((CanTot*PreVta) *(1-DSC001/100)*(1-DSC002/100)*(1-DSC003/100)*(1-DSC004/100))) * 100)/(CanTot*PreVta))/100 ) End AS IMPTOT, " & _
                            " CASE WHEN Md.DSC001 <> 0 THEN CONVERT(VARCHAR(20),MD.DSC001) ELSE '' END as DSC001, Md.DSC002, Md.DSC003, Md.DSC004,  " & _
                            " Md.PreVta AS PreVta1,MD.CanTot,case WHEN Md.INGSAL = 'N' THEN 'Opt' ELSE '' END AS NEUTRO, Md.PreVta*MD.CanTot  AS PREITE, " & _
                            " (CanTot*PreVta) * " & _
                            " CASE WHEN DSC001 <> 0 THEN (DSC001/100) ELSE 1 END * " & _
                            " CASE WHEN DSC002 <> 0 THEN (DSC002/100) ELSE 1 END * " & _
                            " CASE WHEN DSC003 <> 0 THEN (DSC003/100) ELSE 1 END * " & _
                            " CASE WHEN DSC004 <> 0 THEN (DSC004/100) ELSE 1 END AS TOTDSC, " & _
                            " (MD.PreVta*MD.CanTot) - (CanTot*PreVta) * " & _
                            " CASE WHEN DSC001 <> 0 THEN (DSC001/100) ELSE 1 END * " & _
                            " CASE WHEN DSC002 <> 0 THEN (DSC002/100) ELSE 1 END * " & _
                            " CASE WHEN DSC003 <> 0 THEN (DSC003/100) ELSE 1 END * " & _
                            " CASE WHEN DSC004 <> 0 THEN (DSC004/100) ELSE 1 END AS PRETOT, " & _
                            " Case When Md.PreVta = 0 Then 0 Else  (((CanTot*PreVta) -((CanTot*PreVta) *(1-DSC001/100)*(1-DSC002/100)*(1-DSC003/100)*(1-DSC004/100))) * 100)/(CanTot*PreVta)   End as PorDesc " & _
                            " FROM " & IMovDet & " MD " & _
                            " Left Join ARTICULO A On A.CODEMP = '" & ICodEmp & "' And Md.CODSUBALM = A.CODSUBALM And Md.CODART=A.CODART " & _
                            " Left Join ARTDET AD On AD.CODEMP = A.CODEMP And A.CODSUBALM = AD.CODSUBALM And A.XTIPALM = AD.XTIPALM And A.CODART = AD.CODART " & _
                            " Left Join XSUBARTDET X17 On AD.XTIPALM = X17.XTIPALM And AD.XSUBART17 = X17.CODSUB And X17.XSUBART = 'XSUBART17' " & _
                            " Left Join XSUBARTDET X04 On AD.XTIPALM = X04.XTIPALM And AD.XSUBART04 = X04.CODSUB And X04.XSUBART = 'XSUBART04' " & _
                            " Left Join MAETABDET MAE On MAE.CODTAB = 'XTIPUNI' And Md.XTIPUNI = MAE.CODITE " & _
                            " WHERE NUMDOC='" & INumDoc & "' And ISNULL(MD.NumIte1,'')='' ORDER BY Md.NumIte", , adCmdText)
        
        Set RsDetLar = MPcx.Execute("Select CODDOC,NUMDOC,DETLAR AS COMENTARIO, " & _
                                    " ISNULL((Select DETLAR " & _
                                    " From MOVDETLAR " & _
                                    " WHERE CODEMP='" & Trim$("" & RsEmp!CodEmp) & "' And CODDOC = '" & Trim$("" & RsCab!CodDoc) & "' And NUMDOC='" & INumDoc & "' And " & _
                                    " NUMDETLAR = 1),'') AS FORPAG " & _
                                    " From MOVDETLAR " & _
                                    " WHERE CODEMP = '" & Trim$("" & ICodEmp) & "' And CODDOC = '" & Trim$("" & RsCab!CodDoc) & "' And NUMDOC='" & INumDoc & "' And NUMDETLAR = 1 ", adCmdText)
        
        Set RsPadre = MPcx.Execute("SELECT DISTINCT CODDOCPAD,NUMDOCPAD FROM MOVCTAART WHERE NUMDOC = '" & INumDoc & "' AND CODDOC = 'OCN' AND CODDOCPAD ='RQI'", adCmdText)
        
        '============CODIGO DE BARRAS=====================
        
        'Set rsCab2 = MPcx.Execute("select CODDOC,NUMDOC,CODSUBALM,CODART,DESART,CANTOT from " & IMovDet & " where NUMDOC= '" & INumDoc & "' AND CODDOC ='" & ICodDoc & "'", , adCmdText)
        
        
'        If bVerFrm Then
'        Dim oFrm As New frmOpcImp
'        If oFrm.ShowForm() Then
'            BVERSPOT = oFrm.VerSello
'            BVERPERIODO = oFrm.VerPeriodo
'            bForGloFc = oFrm.ForGloFc
'            bGloGra = oFrm.GloGra
'        End If
'
'
        
        
        '=================================================
        Dim rsTemp As ADODB.Recordset
        Set rsTemp = New ADODB.Recordset
           
        With rsTemp.Fields
            .Append "Nro", adDouble, , adFldIsNullable
            .Append "NumIte", adChar, 4, adFldIsNullable
            .Append "CODART", adChar, 30, adFldIsNullable
            .Append "CanTot", adDouble, , adFldIsNullable
            .Append "DesArt", adVarChar, 600, adFldIsNullable
            .Append "UNI", adVarChar, 50, adFldIsNullable
            .Append "PreVta1", adDouble, , adFldIsNullable
            .Append "PREITE", adDouble, , adFldIsNullable
            .Append "PorDesc", adDouble, , adFldIsNullable
           
            
        End With
        
        rsTemp.Open
        
        Dim nI As Integer 'Contador de Nro registros que tiene el detalle
        Dim nMax As Integer  'Numero maximo de registros a contener el detalle para luego insertar registros en blanco si hace falta
        Dim I As Integer
        Dim X As Integer 'item temporal
        Dim XTemp As Integer
        
        Dim BK As Variant
        Dim Nite As String
        
        nI = 0
        nMax = 21
        rsDes.MoveFirst
        
        
        
        X = 0
        Igv = Trim$("" & rsDes!Imp001)
      Do While rsDes.EOF = False
            
            If Trim$(rsDes!NumIte1) = "" Then
            
                X = X + 1
                
                XTemp = 0
                sCad = ""
                sCadAux = ""
                If Trim$("" & rsDes!DEtART1) = "" Then
                    sCadAux = Trim$("" & rsDes!DESART)
                Else
                    sCadAux = Trim$("" & rsDes!DEtART1)
                End If
            ' =========IMPRIME ARTICULO
                Call Cortar(sCadAux, CLng(iLonDet), True, False, sTmp(), lTmp, False)
                 
                 For lCnt = 1 To lTmp
                  If lCnt = 1 Then
                        nI = nI + 1
                        rsTemp.AddNew
                        rsTemp.Fields("Nro") = nI
                        rsTemp.Fields("NUMITE") = Format$(X, "00") ' Trim$("" & rsDes!NUMITE)
                        rsTemp.Fields("CODART") = Trim$("" & rsDes!CODART)
                        rsTemp.Fields("CANTOT") = rsDes!CANTOT
                        rsTemp.Fields("DesArt") = sTmp(lCnt) '' Trim$("" & rsDes!DESART)
                        rsTemp.Fields("UNI") = Trim$("" & rsDes!Uni)
                        rsTemp.Fields("PREVTA1") = rsDes!PreVta1
                        rsTemp.Fields("PREITE") = rsDes!IMPTOT
                        rsTemp.Fields("PorDesc") = rsDes!PorDesc
                        dSubTotal = dSubTotal + rsDes!IMPTOT
                        rsTemp.Update
                    Else
                        nI = nI + 1
                        rsTemp.AddNew
                        rsTemp.Fields("Nro") = nI
                        rsTemp.Fields("DesArt") = sTmp(lCnt)
                        rsTemp.Update
                    End If
                Next lCnt
                
                '==================================
                '========= Imprimir Detalle de Articulo =======
                sCad = ""
                sCad = Trim$("" & rsDes!DesArt1)
                 lLen = Len(sCad)
                 If lLen <> 0 Then
'''                I
                                Call Cortar(sCad, 50, True, False, sTmp(), lTmp, False)
                                For lCnt = 1 To lTmp
                                    If lCnt = 1 Then
                                          nI = nI + 1
                                      ''  If Pos < Len(campoB) Then
                                            
                                            rsTemp.AddNew
                                            rsTemp.Fields("Nro") = nI
                                            rsTemp.Fields("DesArt") = sTmp(lCnt) '' Mid$(Trim$(" " & campoB), i, (Pos - 1) - i)
                                            rsTemp.Update
                                          Else
                                            nI = nI + 1
                                            rsTemp.AddNew
                                            rsTemp.Fields("Nro") = nI
                                            rsTemp.Fields("DesArt") = sTmp(lCnt) '''Mid$(Trim$(" " & campoB), i - 1)
                                            rsTemp.Update
                                       End If
''
                                   
                                Next lCnt
                                    '==========================================
''
''
                End If
                '========= Fin Imprimir Detalle ===
                '==================================
                
                
                '===================================
                'IMPRIME RECETAS DEL ARTICULO ACTUAL
                '===================================
'                If bImpRec Then
'
'                    Nite = rsDes!NumIte
'
'                    BK = rsDes.Bookmark
'                    rsDes.Filter = "NUMITE1='" & Trim$(rsDes!NumIte) & "'"
'
'                    Do While rsDes.EOF = False
'                        nI = nI + 1
'                        rsTemp.AddNew
'                        rsTemp.Fields("Nro") = nI
'                        rsTemp.Fields("DesArt") = "     " & Mid$(Trim$("" & rsDes!DESART), 1, 46)
'                        rsTemp.Update
'                        rsDes.MoveNext
'                    Loop
'
'                    rsDes.Filter = ""
'                    rsDes.Bookmark = BK
'                End If
                
            End If
        
            rsDes.MoveNext
        Loop
        For I = nI + 1 To nMax
            rsTemp.AddNew
            rsTemp.Fields("Nro") = I
            rsTemp.Update
        Next
        '=============================
        
        rsTemp.MoveFirst
                
        Set rpt = New rptOrdenCompra
                      
        With rpt.Sections("CABPAG")
            
'            IMPRIMIA CON LOGO
'            rpt.Sections("CABPAG").Controls("imgGraf").Visible = frm.PrintGraf

             'Set .Controls("imgGraf").Picture = GetLogo(Trim$("" & RsEmp!RutLog))
    ' //Modificado x LRH .Controls("lblEmpresa").Caption = Trim$("" & rsEmp!NomEmp)
            .Controls("lbldiremp").Caption = Trim$("" & RsEmp!DirEmp)
           
            .Controls("lbltelemp").Caption = Trim$("" & RsEmp!TELEMP)
            .Controls("lblfaxemp").Caption = Trim$("" & RsEmp!FaxEmp)
            .Controls("lblwebemp").Caption = Trim$("" & RsEmp!BKUPEMP)
            
            
            .Controls("lblDesDoc").Caption = UCase(Trim$("" & RsCab!DESCRI))
'            .Controls("lblCodVen").Caption = Trim$("" & RsCab!CODVEN)
            .Controls("lblDocumento").Caption = Trim$("" & RsCab!NumDoc)
            .Controls("lblFechaent").Caption = Format$(RsCab!FecEnt02, "dd mmmm yyyy")
            .Controls("lblFecha").Caption = Format$(RsCab!FecDoc, "dd mmmm yyyy")
'            .Controls("lblFecha").Caption = Format$(RsCab!FECDOC, "dd/mm/yyyy")
            'Facturar A:
            
            
            .Controls("lblNomAne").Caption = Trim$("" & RsEmp!nomemp)
            .Controls("lblDirAne").Caption = Trim$("" & RsEmp!DirEmp)
            .Controls("lblIdeAne").Caption = Trim$("" & RsEmp!RUCEMP)
            .Controls("lblTipVta").Caption = Trim$("" & RsCab!TIPVTA)
            .Controls("lblGlosa").Caption = Trim$("" & RsCab!Glosa)
            .Controls("lblformapago").Caption = Trim$("" & RsCab!CONPAG)
            '================
            .Controls("lblNomAne1").Caption = Trim$("" & RsCab!NOMANE02)
            .Controls("lblNomAne2").Caption = Trim$("" & RsCab!NomAne)
            .Controls("lblDirAne2").Caption = Trim$("" & RsCab!DIRANE & " - " & Trim$("" & RsCab!XUBIGEO))
            '------------------------------------------------------
            '//Lrh.Controls("lblTelPro").Caption = Trim$("" & rsCab!TELANE)
            '.Controls("lblDirAne1").Caption = Trim$("" & rsCab!DIRANE02)
           
           '//Lrh  .Controls("lblFaxAne").Caption = Trim$("" & rsCab!FAXANE)
'            .Controls("lblTelAne1").Caption = Trim$("" & RsCab!TELENT)
          '  .Controls("lblMailContacto").Caption = Trim$("" & rsCab!MAILCON)
         
            
'            .Controls("lblFaxAne1").Caption = Trim$("" & RsCab!FAXANE)
            
''            .Controls("lblCodBri").Caption = Trim$("" & RsCab!CodBri) 'codigo de como lo concen a Bright el proveedor
            .Controls("lblCodPro").Caption = Trim$("" & RsCab!CODANE)
          '   .Controls("lblMailVen").Caption = Trim$("" & rsCab!MAILRES)
'            .Controls("lblCelVen").Caption = Trim$("" & RsCab!CELVEN)
           ' .Controls("lblNomVen").Caption = Trim$("" & rsCab!NomVen)
'            .Controls("lblReferencia").Caption = Trim$("" & rsCab!REFINT)
'            .Controls("lblCodCli").Caption = Trim$("" & rsCab!CODANE)
          .Controls("lblcotizacion").Caption = Trim$("" & RsCab!NumDoc)
            
           ' *************Datos Responsable ************************
            .Controls("lblAtencion").Caption = UCase(Trim$("" & RsCab!NomRes))
'            .Controls("lblCarCon").Caption = Trim$("" & RsCab!XDESCAR)
'              .Controls("lblNomVen").Caption = Trim$("" & rsCab!NomRes)
            .Controls("lbltlfno").Caption = Trim$("" & RsCab!TELRES)
            .Controls("lblTelAne").Caption = Trim$("" & RsEmp!TELEMP)
'             .Controls("lblMailVen").Caption = Trim$("" & rsCab!MailRes)
             '*********** Datos Contacto ************************
             
            .Controls("lblcontacto").Caption = Trim$("" & RsCab!NOMCON)
            .Controls("lbltelconta").Caption = Trim$("" & RsCab!CellCon)
             .Controls("lblMailContacto").Caption = Trim$("" & RsCab!Mail)

            If Trim$("" & RsCab!SIMMON) = "US$" Then
                .Controls("lblMoneda").Caption = Trim$("US DOLARES")
            ElseIf Trim$("" & RsCab!SIMMON) = "S/." Then
                .Controls("lblMoneda").Caption = Trim$("SOLES")
            End If

            If Trim$("" & RsCab!NUMRF1) <> "" Then
                '.Controls("lblNumRf1").Caption = Trim$("" & RsCab!NUMRF1)
                Dim sRef As String
                Dim iRef As String
                iRef = 0
                Do While RsPadre.EOF = False
                    sRef = sRef & Trim$("" & RsPadre!NumDocPad) & " "
                '    .Controls("lblNumRf1").Caption = sRef
                    RsPadre.MoveNext
                Loop
            
'                Call Cortar(sRef, 10, True, False, sTmp(), lTmp, False)
'                For lCnt = 1 To lTmp
'                    .Controls("lblNumRf1").Caption = sTmp(lCnt)
'                Next lCnt
            Else
                '.Controls("lblNumRf1").Caption = ""   ' PROBLEMA
                '.Controls("Etiqueta12").Caption = ""
            End If
            
           
'            .Controls("lblNomEnt").Caption = Trim$("" & RsCab!NOMCONENT)
            .Controls("lblPU").Caption = Trim$("P.Unitario ")
            .Controls("lblTot").Caption = Trim$("P.Total ")
       End With
        
        With rpt.Sections("PieInforme")
            .Controls("lblTipMon0").Caption = Trim$("" & RsCab!SIMMON)
            .Controls("lblTipMon1").Caption = Trim$("" & RsCab!SIMMON)
            .Controls("lblTipMon2").Caption = Trim$("" & RsCab!SIMMON)
            .Controls("lblETIGV").Caption = "IGV ( " & Igv & " % ) "
            dSubTotal = dSubTotal - dDesc1 - dDesc2 - dDesc3 - dDesc4
            .Controls("lblSubTotal").Caption = Format$(dSubTotal, "#,0.00")
            .Controls("lblIGV").Caption = Format$(RsCab!totimp1, "#,0.00")
            .Controls("lblTotal").Caption = Format$(RsCab!TotDoc, "#,0.00")
            .Controls("lblCodMem").Caption = Trim$("" & RsCab!CodMem)
            .Controls("lblComMem").Caption = Trim$("" & RsCab!ComMem)
'            .Controls("lblHORA").Caption = ("" & rsCab!FECPR)
            
        End With
        
                
        Set rpt.DataSource = rsTemp
        Screen.MousePointer = 0
        If PreView Then
            rpt.Show 1
        Else
            rpt.PrintReport
        End If
        Unload rpt
        Set rpt = Nothing
        rsDes.Close
        Set rsDes = Nothing
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Screen.MousePointer = 11
        MsgBox "No se encontro el documento", vbExclamation
    End If

    RsCab.Close
    Set RsCab = Nothing
    RsEmp.Close
    Set RsEmp = Nothing
End Sub
'' pasar de un excel a dbf ... con un proyecto de VB




