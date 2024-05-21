Attribute VB_Name = "modFEv1"
Option Explicit

Public Function cae(comprobanteId As Integer, clienteId As Long, pTotal As Currency, pExento As Currency, pNeto As Currency, pIVA As Currency, pProduccion As Boolean, ByVal pAlicuotaIVA As Currency, pDB As clsDB, pNroComprob As Long, pCAEVenc As String, pBarras As String, Optional clientemovimientoIdasociado As Variant = Null) As String
Dim WSAA As Object
Dim WSFEv1 As Object

Dim strCert As String
Dim strClav As String
Dim strFecha As String
Dim strCAE As String
Dim strExcepcion As String
Dim strCache As String
Dim strProxy As String
Dim strWrapper As String
Dim strCACert As String
Dim strWSDL As String
Dim strDocu As String
Dim strToken As String
Dim strSign As String
Dim strDest As String
Dim strTA_xml As String
Dim fechaAsociado As String
Dim empresaCuit As String

Dim comprobante As New clsDAOTiposComprob
Dim cliente As New clsDAOClientes
Dim electronico As New clsDAORegistroCAE

Dim comprobanteAsociado As New clsMODComprobante
Dim electronicoAsociado As clsMODElectronico
Dim clienteMovimientoAsociado As clsMODClienteMovimiento

Dim comprobanteRep As New clsREPComprobante
Dim electronicoRep As New clsREPElectronico
Dim clienteMovimientoRep As New clsREPClienteMovimiento

Dim lngTTL As Long

Dim intFD As Integer
Dim intTipo As Integer

Dim tra
Dim cms
Dim ok

Dim varUltimo As Variant
Dim v As Variant
    
On Error GoTo ManejoError
    
    cae = ""
    
    comprobante.codigo = comprobanteId
    comprobante.findByPrimaryKey pDB
    
    If comprobante.factelect = 0 Then Exit Function
    
    With cliente
        .findByCodigo clienteId, pDB
        
        intTipo = 80
        strDocu = Trim(Replace(.cuit, "-", ""))
        
        If strDocu = "" Then
            intTipo = 96
            strDocu = Trim(.nrodoc)
        End If
    End With
    
    If Not IsNull(clientemovimientoIdasociado) Then
        Set clienteMovimientoAsociado = clienteMovimientoRep.findByClientemovimientoId(clientemovimientoIdasociado)
        Set electronicoAsociado = electronicoRep.findByComprobante(clienteMovimientoAsociado.comprobanteId, clienteMovimientoAsociado.puntoventa, clienteMovimientoAsociado.numerocomprobante)
        Set comprobanteAsociado = comprobanteRep.findByComprobanteId(clienteMovimientoAsociado.comprobanteId)
    End If
    
    Set WSAA = CreateObject("WSAA")
    
    Debug.Assert WSAA.Version >= "2.04a"
    ' deshabilito errores no manejados (version 2.04 o superior)
    WSAA.LanzarExcepciones = False
    
    ' datos de prueba del certificado (para depuración):
    If pProduccion Then
        empresaCuit = Replace(gEmpresa.cuit, "-", "")
        strDest = "C=AR, O=" & gEmpresa.rsocial & ", serialNumber=CUIT " & empresaCuit & ", CN=Facturacion"
    Else
        empresaCuit = "23236938409"
        strDest = "C=AR, O=Daniel Eusebio Quinteros, serialNumber=CUIT " & empresaCuit & ", CN=Facturacion"
    End If
    
    ' inicializo las variables:
    strToken = ""
    strSign = ""
        
    If Dir("ta.xml") <> "" Then
        ' leo el xml almacenado del archivo
        Open "ta.xml" For Input As #1
        Line Input #1, strTA_xml
        Close #1
        ' analizo el ticket de acceso previo:
        ok = WSAA.AnalizarXml(strTA_xml)
        ' verifico que el destino corresponda (CUIT)
        Debug.Assert WSAA.ObtenerTagXml("destination") = strDest
        ' verificar CUIT
        If Not WSAA.Expirado() Then
            ' puedo reusar el ticket de acceso:
            strToken = WSAA.ObtenerTagXml("token")
            strSign = WSAA.ObtenerTagXml("sign")
        End If
    End If
        
    If strToken = "" Or strSign = "" Then
        ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
        lngTTL = 43200 ' tiempo de vida hasta expiración
        tra = WSAA.CreateTRA("wsfe", lngTTL)
        controlarExcepcion WSAA
        Debug.Print tra
    
        ' Certificado: certificado es el firmado por la AFIP
        ' ClavePrivada: la clave privada usada para crear el certificado
        If pProduccion Then
            strCert = App.Path & "\" & modCert.cntCertificado & ".crt" ' certificado
            strClav = App.Path & "\" & modCert.cntCertificado & ".key"  ' clave privada
        Else
            strCert = App.Path & "\dqmdz.crt" ' certificado
            strClav = App.Path & "\dqmdz.key" ' clave privada
        End If

        cms = WSAA.SignTRA(tra, strCert, strClav)
        controlarExcepcion WSAA
        Debug.Print cms
        
        If cms <> "" Then
            ' Conectarse con el webservice de autenticación:
            strCache = ""
            strProxy = "" '"usuario:clave@localhost:8000"
            strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
            strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante
            
            If pProduccion Then
                strWSDL = "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl" ' Producción
            Else
                strWSDL = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl" ' Homologación
            End If
            
            ok = WSAA.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert)
            controlarExcepcion WSAA
        
            ' Llamar al web service para autenticar:
            strTA_xml = WSAA.LoginCMS(cms)
            controlarExcepcion WSAA

            ' Imprimir el ticket de acceso, ToKen y Sign de autorización
            Debug.Print strTA_xml
            Debug.Print "Token:", WSAA.Token
            Debug.Print "Sign:", WSAA.Sign
            
            If strTA_xml <> "" Then
                ' guardo el ticket de acceso en el archivo
                Open "ta.xml" For Output As #1
                Print #1, strTA_xml
                Close #1
            End If
            
            strToken = WSAA.Token
            strSign = WSAA.Sign
        End If
        ' reviso que no haya errores:
        Debug.Print "excepcion", WSAA.Excepcion
        If WSAA.Excepcion <> "" Then
            Debug.Print WSAA.Traceback
            MsgBox WSAA.Excepcion, vbCritical, "Excepción"
        End If
    End If

    ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    
    ' Setear tocken y sign de autorización (pasos previos)
    WSFEv1.Token = strToken
    WSFEv1.Sign = strSign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    If pProduccion Then
        WSFEv1.cuit = Replace(gEmpresa.cuit, "-", "")
    Else
        WSFEv1.cuit = "23236938409"
    End If
    
    ' deshabilito errores no manejados
    WSFEv1.LanzarExcepciones = False

    ' Conectar al Servicio Web de Facturación
    strProxy = "" ' "usuario:clave@localhost:8000"
    strCache = "" 'Path
    strWrapper = "" ' libreria http (httplib2, urllib2, pycurl)
    strCACert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)
    
    If pProduccion Then
        strWSDL = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL" ' producción
    Else
        strWSDL = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" ' homologación
    End If
    
    ok = WSFEv1.Conectar(strCache, strWSDL, strProxy, strWrapper, strCACert) ' homologación
    Debug.Print WSFEv1.Version
    controlarExcepcion WSFEv1
    
    ' mostrar bitácora de depuración:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    controlarExcepcion WSFEv1
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus

    ' Establezco los valores de la factura o lote a autorizar:
    ' Establezco los valores de la factura a autorizar:
    varUltimo = WSFEv1.CompUltimoAutorizado(comprobante.tipoafip, comprobante.puntovta)
    controlarExcepcion WSFEv1
    For Each v In WSFEv1.errores
        Debug.Print v
    Next
    Debug.Print WSFEv1.errmsg
    Debug.Print WSFEv1.errcode
    If varUltimo = "" Then
        varUltimo = 0                ' no hay comprobantes emitidos
    Else
        varUltimo = CLng(varUltimo)   ' convertir a entero largo
    End If
    
    strFecha = Format(Date, "yyyymmdd")
    
    ok = WSFEv1.crearFactura(1, intTipo, strDocu, comprobante.tipoafip, comprobante.puntovta, _
        varUltimo + 1, varUltimo + 1, Format(Abs(pTotal), "0.00"), Format(Abs(pExento), "0.00"), Format(Abs(pNeto), "0.00"), Format(Abs(pIVA), "0.00"), _
        "0.00", "0.00", strFecha, , , , "PES", "1.000")
        
    ' Agrego tasas de IVA
    If pIVA <> 0 Or pAlicuotaIVA = 21 Then ok = WSFEv1.AgregarIva(5, Format(Abs(pNeto), "0.00"), Format(Abs(pIVA), "0.00"))
    If pAlicuotaIVA = 10.5 Then ok = WSFEv1.AgregarIva(4, Format(Abs(pNeto), "0.00"), Format(Abs(pIVA), "0.00"))
    ' Agrego comprobante asociado si es necesario
    If comprobante.asociado = 1 Then
        If Not IsNull(clienteMovimientoAsociado.clientemovimientoId) Then
            fechaAsociado = Format(clienteMovimientoAsociado.fechacomprobante, "yyyymmdd")
            ok = WSFEv1.agregarCmpAsoc(CInt(comprobanteAsociado.comprobanteafipId), clienteMovimientoAsociado.puntoventa, clienteMovimientoAsociado.numerocomprobante, empresaCuit, fechaAsociado)
        End If
    End If
    
    ' Habilito reprocesamiento automático (predeterminado):
    WSFEv1.Reprocesar = True

    ' Solicito CAE:
    strCAE = WSFEv1.CAESolicitar()
    controlarExcepcion WSFEv1
    
    Debug.Print "Resultado", WSFEv1.Resultado
    Debug.Print "CAE", WSFEv1.cae

    Debug.Print "Numero de comprobante:", WSFEv1.cbtenro
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo
    
    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & WSFEv1.cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    pNroComprob = WSFEv1.cbtenro
    pCAEVenc = modFecha.YYYYMMDD2DDMMYYYY(WSFEv1.Vencimiento)
    pBarras = generateI2of5(Replace(gEmpresa.cuit, "-", ""), comprobante.tipoafip, comprobante.puntovta, strCAE, pCAEVenc)
    
    electronico.tcoID = comprobanteId
    electronico.prefijo = comprobante.puntovta
    electronico.nroComprob = WSFEv1.cbtenro
    electronico.cliID = clienteId
    electronico.total = pTotal
    electronico.neto = pNeto
    electronico.iva = pIVA
    electronico.cae = strCAE
    electronico.fecha = Format(Date, "ddmmyyyy")
    electronico.caevenc = pCAEVenc
    electronico.barras = pBarras
    electronico.tipodocumento = intTipo
    electronico.numerodocumento = strDocu
    electronico.clientemovimientoIdasociado = clientemovimientoIdasociado

    electronico.save pDB
        
    cae = strCAE
    
    Exit Function
    
ManejoError:
    ' Si hubo error (tradicional, no controlado):
    
    ' Depuración (grabar a un archivo los detalles del error)
    intFD = FreeFile
    Open "c:\error.txt" For Append As intFD
    If Not WSAA Is Nothing Then
        If WSAA.Version >= "1.02a" Then
            Print #intFD, WSAA.Excepcion
            Print #intFD, WSAA.Traceback
            Print #intFD, WSAA.XmlRequest
            Print #intFD, WSAA.XmlResponse
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSAA.Excepcion
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.Version >= "1.10a" Then
            Print #intFD, WSFEv1.Excepcion
            Print #intFD, WSFEv1.Traceback
            Print #intFD, WSFEv1.XmlRequest
            Print #intFD, WSFEv1.XmlResponse
            Print #intFD, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            strExcepcion = WSFEv1.Excepcion
        End If
    End If
    Close intFD
    
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    If strExcepcion = "" Then                 ' si no tengo mensaje de excepcion
        strExcepcion = Err.Description        ' uso el error de VB
    End If
    
    ' Mostrar el mensaje de error
    Select Case MsgBox(strExcepcion, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    
End Function

Private Sub controlarExcepcion(obj As Object)
Dim intFD As Integer

    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        intFD = FreeFile
        Open "c:\excepcion.txt" For Append As intFD
        Print #intFD, obj.Excepcion
        Print #intFD, obj.Traceback
        Print #intFD, obj.XmlRequest
        Print #intFD, obj.XmlResponse
        Close intFD
        MsgBox obj.Excepcion, vbExclamation, "Excepción"
        End
    End If
    
End Sub

Public Function generateI2of5(pCUIT As String, comprobanteId As Integer, pPuntoVta As Integer, pCAE As String, pCAEVenc As String) As String
Dim strCadena As String

Dim intVerif As Integer

    strCadena = pCUIT & Trim(Format(comprobanteId, "00")) & Trim(Format(pPuntoVta, "0000")) & pCAE & pCAEVenc
    
    intVerif = calculateVerif(strCadena)
    
    strCadena = strCadena & Trim(Str(intVerif))
    
    generateI2of5 = i2of5(strCadena)
    
End Function

Private Function i2of5(ByVal pCadena As String) As String
Dim strCadena As String
Dim strCaracter As String
Dim intCiclo As Integer
Dim intValor As Integer
    
    If (Len(pCadena) And 1) = 1 Then Exit Function
    strCadena = Chr(33)
    For intCiclo = 1 To Len(pCadena) Step 2
        intValor = Val(Mid(pCadena, intCiclo, 2))
        Select Case intValor
            Case 0 To 3
                strCaracter = Chr(intValor + 35)
            Case 4
                strCaracter = "'"
            Case 5 To 91
                strCaracter = Chr(intValor + 35)
            Case 92
                strCaracter = Chr(196)
            Case 93
                strCaracter = Chr(197)
            Case 94
                strCaracter = Chr(199)
            Case 95
                strCaracter = Chr(201)
            Case 96
                strCaracter = Chr(209)
            Case 97
                strCaracter = Chr(214)
            Case 98
                strCaracter = Chr(220)
            Case 99
                strCaracter = Chr(225)
        End Select
        strCadena = strCadena & strCaracter
    Next
    strCadena = strCadena & Chr(34)
    i2of5 = strCadena

End Function

Private Function calculateVerif(pCadena As String) As Integer
Dim intPares As Integer
Dim intImpares As Integer
Dim intCiclo As Integer
Dim intResultado As Integer
Dim intMultiplo As Integer
Dim intIzquierda As Integer
Dim intDerecha As Integer
    
    intPares = 0
    intImpares = 0
    
    For intCiclo = 1 To Len(pCadena)
        If intCiclo Mod 2 = 0 Then
            intPares = intPares + CInt(Mid(pCadena, intCiclo, 1))
        Else
            intImpares = intImpares + CInt(Mid(pCadena, intCiclo, 1))
        End If
    Next intCiclo
    
    intImpares = 3 * intImpares
    
    intResultado = intPares + intImpares
    
    intMultiplo = Int(intResultado / 10) * 10
    
    intIzquierda = intResultado - intMultiplo
    
    intMultiplo = intMultiplo + 10
    
    intDerecha = intMultiplo - intResultado
    
    calculateVerif = IIf(intIzquierda < intDerecha, intIzquierda, intDerecha)

End Function

