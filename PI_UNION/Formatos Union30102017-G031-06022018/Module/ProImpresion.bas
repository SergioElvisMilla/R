Attribute VB_Name = "ProImpresion"
'' Hecho en el 2010 a pedido de Tricar

Function GenSqlCtaArtPad(pCodDoc As String, pNumDoc As String, pSubAlm As String, pCodArt As String) As String
' Funcion que genera una cadena query para sacar el documento padre
Dim sCadSql As String
sCadSql = "Select coddocpad, numdocpad, codsubalm, codart from movctaart where coddoc = '" & pCodDoc & "' and numdoc = '" & pNumDoc & "' and codsubalm = '" & pSubAlm & "' and codart = '" & pCodArt & "'"
GenSqlCtaArtPad = sCadSql
End Function

Function GenSqlCtaPadCod(pCodDoc As String, pNumDoc As String) As String
' Funcion que genera una cadena query para sacar el documento padre
Dim sCadSql As String
sCadSql = "Select distinct coddocpad, numdocpad from movctaart where coddoc = '" & pCodDoc & "' and numdoc = '" & pNumDoc & "'"
GenSqlCtaPadCod = sCadSql
End Function
Function GenSqlCtaHijoOT(pCodDocPad As String, pNumDocPad As String) As String
' Funcion que genera una cadena query para sacar el documento hijo
Dim sCadSql As String
sCadSql = "Select distinct coddoc, numdoc from movctaart where coddocpad = '" & pCodDocPad & "' and numdocPad = '" & pNumDocPad & "' and coddoc like 'NI%'"
GenSqlCtaHijoOT = sCadSql
End Function
Function GenSqlRecetaDoc(pNomTab As String, pNumDoc As String) As String
' Funcion que genera una cadena query para sacar la receta de un documento
Dim sCadSql As String
sCadSql = "Select coddoc,numdoc,substring(codart,1,3)tipalm,codart,desart,xtipuni,sum(cantot)cant from " + pNomTab + " where numdoc = '" & pNumDoc & "' and ingsal='N' " & _
          "group by coddoc,numdoc,codart,desart,xtipuni "
GenSqlRecetaDoc = sCadSql
End Function
Function GenSqlRecetaLevel1(pNomTab As String, pNumDoc As String) As String
' Funcion que genera una cadena query para sacar la receta de un documento
Dim sCadSql As String
sCadSql = "Select r.codart,r.desart,a1.xtipuni,r.cantot*sum(m.cantot)cantreq ,ceiling(r.cantot*sum(m.cantot)) cantped from " + pNomTab + " m inner join articulo a on a.codart=m.codart and a.codsubalm=m.codsubalm inner join recartdet r on r.codrec=a.codrec join articulo a1 on a1.codart=r.codart and r.codsubalm=a1.codsubalm where m.numdoc = '" & pNumDoc & "'  and m.ingsal='N' and m.tipdoc_p is null group by r.codart,r.desart,a1.xtipuni,r.cantot order by r.desart"
GenSqlRecetaLevel1 = sCadSql
End Function
 Function GenSqlCab(pNomTab As String, pCodDoc As String, pNumDoc As String) As String

'PARAMETROS:
'pNomTab: Nombre de la tabla Cabecera a consultar
'pCodDoc: Tipo de Documento que agrupa los tipos de movimiento
'pNumDoc: Numero del documento

'VALOR DE RETORNO
'Retorna una cadena con la consulta hecha a una tabla cabecera (reforna solo un registro)
'MovCab[Empresa][Tipo Documento]. xEj. MovCabE1CO)
'[Empresa] = 2 bytes
'[Tipo de documento] = Maximo 4 bytes

 
 Dim sCadCampos As String, sCadFrom As String, sCadLeft As String, sCadWhere As String
 sCadCampos = "Select " & _
    "coddoc=isnull(cab.coddoc,''),  numdoc=isnull(cab.numdoc,''),  fecdoc=isnull(cab.fecdoc,''),  fecrea=isnull(cab.fecrea,''),  fecven=isnull(cab.fecven,''),  fecent=isnull(cab.fecent,''),  codane=isnull(cab.codane,''),  nomane=isnull(cab.nomane,''),  dirane=isnull(cab.dirane,''),  telane=isnull(ane.telane,''), " & _
    "faxane=isnull(ane.faxane,''),  celane=isnull(ane.celane,''),  maiane=isnull(ane.mailAne,''), webane=isnull(ane.webane,''),  ref1=Case When Substring(cab.tiptr1,0,2)='FC' Then isnull(cab.tiptr1+' - '+cab.numrf1,'') Else '' End,ref3=Case When Substring(cab.tiptr3,0,2)='FC' Then isnull(cab.tiptr3+' - '+cab.numrf3,'') Else '' End,  " & _
    "tipide=isnull(cab.xtipide,''), deside=isnull(mt1.desite,''),  ideane=isnull(cab.ideane,''),  tipmov=isnull(cab.xtipmov,''), desmov=isnull(mt2.desite,''),  estdoc=isnull(cab.xestdoc,''), conpag=isnull(cab.xconpag,''), desmon=isnull(mt4.desite,''),  abrmon=isnull(mt4.desaux,''),  tipcam=isnull(cab.tipcam,0), " & _
    "refint=isnull(cab.refint,''),  despag=isnull(mt3.desite,''),  serrea=isnull(cab.serrea,''),  numrea=isnull(cab.numrea,''),  tipmon=isnull(cab.xtipmon,''), forenv=isnull(cab.forenv,''),  desfor=isnull(mt6.desite,''),  codmem=isnull(cab.codmem,''),  commem=isnull(cab.commem,''),  ubigeo=isnull(ubi.descripcion,''),  " & _
    "subtot=isnull(cab.subtot,0),   totim1=isnull(cab.totimp1,0),  totim2=isnull(cab.totimp2,0),  totim3=isnull(cab.totimp3,0),  totim4=isnull(cab.totimp4,0),  totdoc=isnull(cab.totdoc,0),   totds1=isnull(cab.totdsc1,0),  totds2=isnull(cab.totdsc2,0),  totds3=isnull(cab.totdsc3,0),  totds4=isnull(cab.totdsc4,0), " & _
    "totimp=isnull(cab.totimp1,0)+  isnull(cab.totimp2,0)+         isnull(cab.totimp3,0)+         isnull(cab.totimp4,0),         totdes=isnull(cab.totdsc1,0)+  isnull(cab.totdsc2,0)+         isnull(cab.totdsc3,0)+         isnull(cab.totdsc4,0),         totltr=isnull(cab.totdocl,''), " & _
    "nomcon=isnull(cn1.nomcon,''),  maicon=isnull(cn1.mailcon,''), telcon=isnull(cn1.telcon,''),  celcon=isnull(cn1.celcon,''),  carcon=isnull(mt5.desite,''), " & _
    "codan2=isnull(cab.codane02,''),noman2=isnull(cab.nomane02,''),diran2=isnull(cab.dirane02,''),tipid2=isnull(an3.xtipide1,''), idean2=isnull(an3.ideane1,''), " & _
    "nomcn1=isnull(cn2.nomcon,''),  maicn1=isnull(cn2.mailcon,''), telcn1=isnull(cn2.telcon,''),  celcn1=isnull(cn2.celcon,''), " & _
    "codres=isnull(cab.codres,''),  nomres=isnull(an1.nomane,''), mailres=isnull(an1.mailane,''), ideaneres=isnull(an1.ideane1,''),  codven=isnull(cab.codven,''),  nomven=isnull(an2.nomane,''), " & _
    "tmposx=isnull(mt2.numentde,0), tmposy=isnull(mt2.numdecde,0), tmotro=isnull(mt2.desaux,''), ideane2=isnull(ane.ideane2,''), XTIPIDE2=isnull(ane.XTIPIDE2,''), fecimp=getdate() "
sCadFrom = "From " + pNomTab + " cab "
sCadLeft = "left join maetabdet mt1 on mt1.codtab = 'XTIPIDE' and mt1.codite=cab.xtipide and mt1.codfil = 'A' " & _
    "left join maetabdet mt2 on mt2.codtab = 'MOV-" + pCodDoc + "' and mt2.codite=cab.xtipmov " & _
    "left join anexo ane on ane.codane=cab.codane " & _
    "left join anexo an1 on an1.codane=cab.codres " & _
    "left join anexo an2 on an2.codane=cab.codven " & _
    "left join anexo an3 on an3.codane=cab.codane02 " & _
    "left join contactos cn1 on cab.codane=cn1.codane and cab.numcon=cn1.numcon " & _
    "left join contactos cn2 on cab.codane02=cn2.codane and cab.numcon02=cn1.numcon " & _
    "left join maetabdet mt3 on mt3.codtab='CONPAG'  and mt3.codite=cab.xconpag " & _
    "left join maetabdet mt4 on mt4.codtab='XTIPMON' and mt4.codite=cab.xtipmon " & _
    "left join maetabdet mt5 on mt5.codtab='XCODCAR' and mt5.codite=cn1.XCODCAR " & _
    "left join maetabdet mt6 on mt6.codtab='FORENV' and mt6.codite=cab.forenv " & _
    "left join ubigeo    ubi on ane.xubigeo = ubi.ubigeo "
    

sCadWhere = "where numdoc = '" + pNumDoc + "'"
    
' Descripción de cada campo
' coddoc: Codigo del Documento
' numdoc: Numero del Documento
' fecdoc: Fecha del Documento
' fecrea: Fecha Real
' fecven: Fecha de Vencimiento
' codane: codigo del Anexo
' nomane: Nombre del Anexo
' dirane: Dirección
' telane: Teléfono
' tipide: Tipo de Identificador
' deside: Descripción del tipo de Identificador
' ideane: Identificador
' tipmov: Tipo de Movimiento
' desmov: Descripción del Tipo de Movimiento
' estdoc: Estado del Documento
' codres: Codigo de Responsable
' desres: Nombre del Resposable
' codven: Codigo de Vendedor
' desven: Nombre del Vendedor
' conpag: Codigo de la Condición de Pago
' despag: Descripción de la Condición de Pago
' serrea: Serie Real
' numrea: Numero Real
' tipmon: Tipo de Moneda
' desmon: Decripción larga del Tipo de Moneda
' abrmon: Descripción corta del Tipo de Moneda
' tipcam: Tipo de Cambio
' subtot: Sub Total del Documento
' totim1: Total Impuesto 1
' totds1: Total descuento 1
 GenSqlCab = sCadCampos + sCadFrom + sCadLeft + sCadWhere

End Function
Function GenSqlDet(pNomTab As String, pNumDoc As String) As String
 
'PARAMETROS:
'pNomTab: Nombre de la tabla Cabecera a consultar
'pNumDoc: Numero del documento

'VALOR DE RETORNO
'Retorna una cadena con la consulta hecha a una tabla detalle (puede retornar varios registros)
'MovDet[Empresa][Tipo Documento]. xEj. MovDetE1CO)
'[Empresa] = 2 bytes
'[Tipo de documento] = Maximo 4 bytes Dim sCadenaSql As String
 Dim sCadCampos As String, sCadFrom As String, sCadLeft As String, sCadWhere As String
 
 
 sCadCampos = "Select " & _
    "coddoc=isnull(det.coddoc,''), numdoc=isnull(det.numdoc,''),    numitp=isnull(det.numite,''),    numith=isnull(det.numite1,''), " & _
    "numord=isnull(det.numord,''), subalm=isnull(det.codsubalm,''), desalm=isnull(alm.nomsubalm,''), codart=isnull(det.codart,''), " & _
    "desart=isnull(det.desart,''), detar1=isnull(det.desart1,''),   desar1=isnull(art.detart,''),    desar2=isnull(art.detart1,''),   tipuni=isnull(det.xtipuni,''), " & _
    "desuni=isnull(mt1.desite,''), cantot=isnull(det.cantot,0),     prevta=isnull(det.prevta,0),     totart=det.totart, " & _
    "imp001=isnull(det.imp001,0),  codrec=isnull(art.codrec,''), " & _
    "cla01=isnull(mdp.nomsubart01,''), cla02=isnull(mdp.nomsubart02,''), cla03=isnull(mdp.nomsubart03,''), cla04=isnull(mdp.nomsubart04,''), cla05=isnull(mdp.nomsubart05,''), cla06=isnull(mdp.nomsubart06,''), cla07=isnull(mdp.nomsubart07,''), cla08=isnull(mdp.nomsubart08,''), cla09=isnull(mdp.nomsubart09,''), cla10=isnull(mdp.nomsubart10,''), cla11=isnull(mdp.nomsubart11,''), cla12=isnull(mdp.nomsubart12,''), cla13=isnull(mdp.nomsubart13,''), cla14=isnull(mdp.nomsubart14,''), cla15=isnull(mdp.nomsubart15,''), " & _
    "cla16=isnull(mdp.nomsubart16,''), cla17=isnull(mdp.nomsubart17,''), cla18=isnull(mdp.nomsubart18,''), cla19=isnull(mdp.nomsubart19,''), cla20=isnull(mdp.nomsubart20,''), cla21=isnull(mdp.nomsubart21,''), cla22=isnull(mdp.nomsubart22,''), cla23=isnull(mdp.nomsubart23,''), cla24=isnull(mdp.nomsubart24,''), cla25=isnull(mdp.nomsubart25,''), cla26=isnull(mdp.nomsubart26,''), cla27=isnull(mdp.nomsubart27,''), cla28=isnull(mdp.nomsubart28,''), cla29=isnull(mdp.nomsubart29,''), cla30=isnull(mdp.nomsubart30,''), " & _
    "tit01=isnull(x01.nomsubart,''), tit02=isnull(x02.nomsubart,''), tit03=isnull(x03.nomsubart,''), tit04=isnull(x04.nomsubart,''), tit05=isnull(x05.nomsubart,''), tit06=isnull(x06.nomsubart,''), tit07=isnull(x07.nomsubart,''), tit08=isnull(x08.nomsubart,''), tit09=isnull(x09.nomsubart,''), tit10=isnull(x10.nomsubart,''), tit11=isnull(x11.nomsubart,''), tit12=isnull(x12.nomsubart,''), tit13=isnull(x13.nomsubart,''), tit14=isnull(x14.nomsubart,''), tit15=isnull(x15.nomsubart,''), " & _
    "tit16=isnull(x16.nomsubart,''), tit17=isnull(x17.nomsubart,''), tit18=isnull(x18.nomsubart,''), tit19=isnull(x19.nomsubart,''), tit20=isnull(x20.nomsubart,''), tit21=isnull(x21.nomsubart,''), tit22=isnull(x22.nomsubart,''), tit23=isnull(x23.nomsubart,''), tit24=isnull(x24.nomsubart,''), tit25=isnull(x25.nomsubart,''), tit26=isnull(x26.nomsubart,''), tit27=isnull(x27.nomsubart,''), tit28=isnull(x28.nomsubart,''), tit29=isnull(x29.nomsubart,''), tit30=isnull(x30.nomsubart,''), " & _
    "xsubartdet05=isnull(xd05.nomsubart,''),xsubartdet06=isnull(xd06.nomsubart,''), xsubartdet07=isnull(xd07.nomsubart,''),xsubartdet08=isnull(xd08.nomsubart,'') "

    sCadFrom = "From " + pNomTab + " det "
    
    sCadLeft = "" & _
    "left join articulo  art on art.codsubalm = det.codsubalm and art.codart  = det.codart " & _
    "left join artdet    artdet on artdet.codsubalm=art.codsubalm and artdet.codart=art.codart and artdet.xtipalm=art.xtipalm " & _
    "left join subalm    alm on alm.codsubalm = det.codsubalm " & _
    "left join maetabdet mt1 on mt1.codtab    = 'XTIPUNI'     and det.xtipuni = mt1.codite and mt1.codfil = 'A' " & _
    "left join movdetpro mdp on det.coddoc    = mdp.coddoc    and det.numdoc  = mdp.numdoc and det.numite = mdp.numite " & _
    "left join xsubart   x01 on art.xtipalm = x01.xtipalm     and x01.xsubart = 'XSUBART01' left join xsubart    x02 on art.xtipalm = x02.xtipalm     and x02.xsubart = 'XSUBART02' left join xsubart    x03 on art.xtipalm = x03.xtipalm     and x03.xsubart = 'XSUBART03' left join xsubart    x04 on art.xtipalm = x04.xtipalm     and x04.xsubart = 'XSUBART04' left join xsubart    x05 on art.xtipalm = x05.xtipalm     and x05.xsubart = 'XSUBART05' " & _
    "left join xsubart   x06 on art.xtipalm = x06.xtipalm     and x06.xsubart = 'XSUBART06' left join xsubart    x07 on art.xtipalm = x07.xtipalm     and x07.xsubart = 'XSUBART07' left join xsubart    x08 on art.xtipalm = x08.xtipalm     and x08.xsubart = 'XSUBART08' left join xsubart    x09 on art.xtipalm = x09.xtipalm     and x09.xsubart = 'XSUBART09' left join xsubart    x10 on art.xtipalm = x10.xtipalm     and x10.xsubart = 'XSUBART10' " & _
    "left join xsubart   x11 on art.xtipalm = x11.xtipalm     and x11.xsubart = 'XSUBART11' left join xsubart    x12 on art.xtipalm = x12.xtipalm     and x12.xsubart = 'XSUBART12' left join xsubart    x13 on art.xtipalm = x13.xtipalm     and x13.xsubart = 'XSUBART13' left join xsubart    x14 on art.xtipalm = x14.xtipalm     and x14.xsubart = 'XSUBART14' left join xsubart    x15 on art.xtipalm = x15.xtipalm     and x15.xsubart = 'XSUBART15' " & _
    "left join xsubart   x16 on art.xtipalm = x16.xtipalm     and x16.xsubart = 'XSUBART16' left join xsubart    x17 on art.xtipalm = x17.xtipalm     and x17.xsubart = 'XSUBART17' left join xsubart    x18 on art.xtipalm = x18.xtipalm     and x18.xsubart = 'XSUBART18' left join xsubart    x19 on art.xtipalm = x19.xtipalm     and x19.xsubart = 'XSUBART19' left join xsubart    x20 on art.xtipalm = x20.xtipalm     and x20.xsubart = 'XSUBART20' " & _
    "left join xsubart   x21 on art.xtipalm = x21.xtipalm     and x21.xsubart = 'XSUBART21' left join xsubart    x22 on art.xtipalm = x22.xtipalm     and x22.xsubart = 'XSUBART22' left join xsubart    x23 on art.xtipalm = x23.xtipalm     and x23.xsubart = 'XSUBART23' left join xsubart    x24 on art.xtipalm = x24.xtipalm     and x24.xsubart = 'XSUBART24' left join xsubart    x25 on art.xtipalm = x25.xtipalm     and x25.xsubart = 'XSUBART25' " & _
    "left join xsubart   x26 on art.xtipalm = x26.xtipalm     and x26.xsubart = 'XSUBART26' left join xsubart    x27 on art.xtipalm = x27.xtipalm     and x27.xsubart = 'XSUBART27' left join xsubart    x28 on art.xtipalm = x28.xtipalm     and x28.xsubart = 'XSUBART28' left join xsubart    x29 on art.xtipalm = x29.xtipalm     and x29.xsubart = 'XSUBART29' left join xsubart    x30 on art.xtipalm = x30.xtipalm     and x30.xsubart = 'XSUBART30' " & _
    "left join xsubartdet xd05 on xd05.xsubart=x05.xsubart and xd05.codsub=artdet.xsubart05 and xd05.xtipalm=x05.xtipalm left join xsubartdet xd06 on xd06.xsubart=x06.xsubart and xd06.codsub=artdet.xsubart06 and xd06.xtipalm=x06.xtipalm left join xsubartdet xd07 on xd07.xsubart=x07.xsubart and xd07.codsub=artdet.xsubart07 and xd07.xtipalm=x07.xtipalm left join xsubartdet xd08 on xd08.xsubart=x08.xsubart and xd08.codsub=artdet.xsubart08 and xd08.xtipalm=x08.xtipalm "
    
    sCadWhere = "where det.numdoc = '" + pNumDoc + "' and (isnull(det.numite1,0) = 0 or ltrim(rtrim(det.numite1)) = '') "
 
 
 'select det.codart, det.numite, det.prevta, det.totart,
'PreVtaIncInt = (det.prevta / (select sum(prevta) from movdete1fc09 det1 where numdoc = '0011136' and det1.codart <> 'INT-ND')) * (select sum(prevta) from movdete1fc09 det2 where det2.numdoc = '0011136' and det2.codart = 'INT-ND')
'from movdete1fc09 det where numdoc = '0011136' and codart <> 'INT-ND'



' Descripcion de cada campo
' coddoc:Codigo del Documento
' numdoc:Numero del Documento
' numitp:numero de Item (Padre : Null o Vacío)
' numith:numero de Item 1 (Hijo)
' numord:Nro. Orden (para impresión si se quiere)
' subalm:Codigo del Almacén
' desalm:Nombre del Almacén
' codart:Codigo del Artículo
' desart:Descripción del Artículo
' desar1:Descripción auxiliar 1 del Articulo
' desar2:Descripción Auxiliar 2 del Artículo
' tipuni:Tipo d Unidad
' desuni:Descripción del tipo de unidad
' cantot:Cantidad Total por item
' prevta:Precio de venta
' totart:Total Articulo.Prevta x Cantot
' imp001: Impuesto 1
' codrec: Codigo de Receta del articulo relacionado
 
 GenSqlDet = sCadCampos + sCadFrom + sCadLeft + sCadWhere + " order by det.numite"

End Function
Function GenSqlEmp(pEmpresa As String) As String
'PARAMETROS
'pEmpresa: Codigo de la Empresa

'VALOR DE RETORNO
' Retorna una cadena con la consulta hecha a la tabla Empresas

Dim sCadenaSql As String
 sCadenaSql = "Select top 1 " & _
    "CodEmp , NomEmp, DirEmp, TelEmp, Faxemp, Pagweb=BkupEmp, RucEmp, RutLog, EMail,RutLog2,RutLog3 " & _
    "From Empresas " & _
    "where codemp = '" + pEmpresa + "'"
        
    ' CodEmp: Codigo de la Empresa
    ' NomEmp: nombre de la Empresa
    ' DirEmp: Direccion de la Empresa
    ' TelEmp: telefono de la empresa
    ' Faxemp: Fax de la Empresa
    ' PagWeb: Portal Web
    ' RucEmp: Ruc de la empresa
    ' RutLog: Ruta y logo de la empresa
    ' Email:  Correo electronico
    
 GenSqlEmp = sCadenaSql
End Function


Function GenSqlDLa(pEmpresa As String, pCodDoc As String, pNumDoc As String) As String
'PARAMETROS
'pEmpresa: Codigo de la empmresa
'pCodDoc:  Codigo del Documento
'pNumDoc:  Numero del Documento
'VALOR DE RETORNO
'Retorna una cadeda con la consulta de la tabla Detalle Largo

Dim sCadenaSql As String
    sCadenaSql = "Select " & _
    "numdetlar, nomdetlar, detlar from movdetlar where " & _
    "codemp = '" + pEmpresa + "' and coddoc = '" + pCodDoc + "' AND NUMDOC ='" + pNumDoc + "'"
    GenSqlDLa = sCadenaSql
End Function
Function GenSqlDoc(pEmpresa As String, pCodDoc As String) As String
    Dim sCadenaSql As String
    sCadenaSql = "Select * from cfgdoc00 a left join maetabdet b on a.xlocal=b.codite and codtab='LOCAL'  " & _
    " where a.coddoc = '" + pCodDoc + "' and a.CodEmp = '" + pEmpresa + "'"
    GenSqlDoc = sCadenaSql
End Function
Function GenSqlLoc(pEmpresa As String, pCodDoc As String, pMaeTab As String) As String
    Dim sCadenaSql As String
    sCadenaSql = "Select Cfg.CodDoc, Cfg.xLocal, Mtd.DesIte,Mtd.DesAux " & _
    "From CfgDoc00 Cfg " & _
    "Left Join MaeTabDet Mtd On Mtd.CodIte = Cfg.xLocal And Mtd.CodTab = '" & pMaeTab & "' " & _
    "Where Cfg.CodEmp = '" + pEmpresa + "' And Cfg.CodDoc = '" + pCodDoc + "' "
    GenSqlLoc = sCadenaSql
End Function
Function GenSqlMaeTab(pCodTab As String) As String
    Dim sCadenaSql As String
    sCadenaSql = "Select * from MaeTabDet where codtab = '" & pCodTab & "' and codfil = 'I'"
    GenSqlMaeTab = sCadenaSql
End Function
Function Justifica_Texto(ByRef Obj As Object, ByVal pCadena As String, ByVal pCurY As Long) As Long

Dim lCuerpo As String
Dim lTamMax As Double, lTamCad As Double, lArrTam() As Double
Dim lPosAct As Integer, lPosAnt As Integer, lNumLin As Integer
Dim lNewLin As String, lArrLin() As String, lParraf() As String
Dim lPosEnter As Integer, lNumPar As Integer
Dim bExDo02 As Boolean

Dim lCadAux As String, lSubCad01 As String, lSubCad02 As String
Dim lContad As Integer, lSpace As Integer, lPosCd2 As Integer

lTamMax = 150           ' Tamaño maximo de una linea dependiendo de la fuente y tamaño de letra.
lCuerpo = pCadena       ' El cuerpo
lPosEnter = InStr(1, lCuerpo, vbCrLf)
lNumPar = 1

ReDim Preserve lParraf(1 To lNumPar)
Do While lPosEnter > 0
    lParraf(lNumPar) = Mid(lCuerpo, 1, lPosEnter + 1)
    If lParraf(lNumPar) = vbCrLf Then lParraf(lNumPar) = " "
    lCuerpo = Mid(lCuerpo, lPosEnter + 2)
    lPosEnter = InStr(1, lCuerpo, vbCrLf)
    lNumPar = lNumPar + 1
    ReDim Preserve lParraf(1 To lNumPar)
Loop
lParraf(lNumPar) = lCuerpo & vbCrLf

Dim lConPar  As Integer
For lConPar = 1 To lNumPar
    lCuerpo = lParraf(lConPar)
    lNumLin = 1
    Do While Len(lCuerpo) > 0
        lPosAnt = 1
        lPosAct = InStr(1, lCuerpo, " ", vbTextCompare)
        lNewLin = Mid(lCuerpo, lPosAnt, lPosAct - 1)
        lTamCad = Obj.TextWidth(lNewLin)
        Do While lTamCad < lTamMax
            lPosAnt = lPosAct
            lPosAct = InStr(lPosAnt + 1, lCuerpo, " ", vbTextCompare)
            If lPosAct = 0 Then
                lNewLin = lCuerpo
                lTamCad = 100000
                bExDo02 = True
            Else
                lNewLin = Mid(lCuerpo, 1, lPosAct - 1)
                lTamCad = Obj.TextWidth(lNewLin)
            End If
        Loop
        ReDim Preserve lArrLin(1 To lNumLin)
        ReDim Preserve lArrTam(1 To lNumLin)
        If bExDo02 Then
            lArrLin(lNumLin) = lCuerpo
            lArrTam(lNumLin) = Obj.TextWidth(lArrLin(lNumLin))
            lCuerpo = ""
            bExDo02 = False
        Else
            lArrLin(lNumLin) = Trim$("" & Mid(lCuerpo, 1, lPosAnt - 1))
            lArrTam(lNumLin) = Obj.TextWidth(lArrLin(lNumLin))
            lNumLin = lNumLin + 1
            lCuerpo = Mid(lCuerpo, lPosAnt + 1)
        End If
    Loop
    
    For lContad = 1 To lNumLin
        lSubCad01 = ""
        lSubCad02 = ""
        If lContad = lNumLin Then
            ' No espaciar xq es ultima linea del parrafo
            lSubCad01 = Trim$(lArrLin(lContad))
            lSubCad02 = ""
        Else
            lSpace = 2
            lCadAux = Trim$(lArrLin(lContad))
            lTamCad = Obj.TextWidth(lCadAux)
            lPosAnt = 1
            lPosAct = InStr(lPosAnt, lCadAux, " ", vbTextCompare)
            Do While lTamCad <= lTamMax
                lSubCad01 = Mid(lCadAux, 1, lPosAct) & Space(lSpace - 1)
                lSubCad02 = Mid(lCadAux, lPosAct + lSpace - 1)
                lCadAux = lSubCad01 & lSubCad02
                lPosAnt = lPosAct + lSpace
                lPosCd2 = InStr(1, lSubCad02, Space(lSpace - 1), vbTextCompare)
                lPosAct = Len(lSubCad01) + lPosCd2
                If lPosCd2 = 0 Then
                   ' Llego al final pero no llego al tamano
                   lPosAnt = 1
                   lPosAct = InStr(lPosAnt, lCadAux, Space(lSpace), vbTextCompare)
                   lSpace = lSpace + 1
                End If
                lTamCad = Obj.TextWidth(lCadAux)
            Loop
            lSubCad01 = Trim$(Left(lCadAux, lPosAnt - 1))
            lSubCad02 = Mid(lCadAux, lPosAnt)
        End If
        Obj.CurrentX = 30 + xPosScreen
        Obj.CurrentY = pCurY + yPosScreen
        Obj.Print lSubCad01
    
        Obj.CurrentX = 180 + xPosScreen - Obj.TextWidth(lSubCad02)
        Obj.CurrentY = pCurY + yPosScreen
        Obj.Print lSubCad02
        
        pCurY = pCurY + 5
    Next
    pCurY = pCurY
Next
Justifica_Texto = pCurY
End Function
Function GenSqlCabCuo(pNomTab As String, pCodDoc As String, pNumDoc As String) As String
Dim sCadCampos As String, sCadFrom As String, sCadLeft As String, sCadWhere As String
sCadCampos = "CodAne, NumPro, DesPro, NroCtas, FlaAct, "
sCadFrom = "from anexopro"
sCadLeft = ""
sCadWhere = ""
''cCab
End Function
Function GenSqlDetCuo(pNomTab As String, pCodDoc As String, pNumDoc As String) As String

End Function

Function ImprimeClasificadores(ByRef Obj As Object, pRsDet As Recordset, pPosX01 As Integer, pPosX02 As Integer, pPosX03 As Integer, ByRef pCurY As Long, pxPosScreen As Integer)
    ''' Clasificadores
    'Clasificador 03
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!tit03)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla03)
    pCurY = pCurY + 5
    'Clasificador 04
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!tit04)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla04)
    pCurY = pCurY + 5
    'Clasificador 05
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!tit05)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla05)
    pCurY = pCurY + 5
    'Clasificador 06
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Tit06)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla06)
    pCurY = pCurY + 5
    'Clasificador 07
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Tit07)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla07)
    pCurY = pCurY + 5
    'Clasificador 08
    Obj.CurrentX = pPosX01 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Tit08)
    Obj.CurrentX = pPosX02 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = pPosX03 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla08)
    pCurY = pCurY + 5
    'Clasificador 10
    Obj.CurrentX = 35 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Tit10)
    Obj.CurrentX = 58 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print ":"
    Obj.CurrentX = 60 + pxPosScreen: Obj.CurrentY = pCurY: Obj.Print Trim$("" & pRsDet!Cla10)
End Function


