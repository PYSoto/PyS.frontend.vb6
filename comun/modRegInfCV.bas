Attribute VB_Name = "modRegInfCV"
Option Explicit

Public Sub makeFile(cuit As String, path As String, desde As Date, hasta As Date, pDBs As Collection, pStatus As StatusBar, pProgreso As ProgressBar)

    cabecera cuit, path, desde, hasta, pStatus
    detalleVentas path, desde, hasta, pDBs, pStatus, pProgreso
    
End Sub

Private Sub cabecera(cuit As String, path As String, desde As Date, hasta As Date, pStatus As StatusBar)
Dim strLinea As String

    pStatus.SimpleText = "Generando CABECERA . . ."
    strLinea = Replace(cuit, "-", "") ' Campo 1
    strLinea = strLinea & Format(Year(desde), "0000") & Format(Month(desde), "00") ' Campo 2
    strLinea = strLinea & "00" ' Campo 3
    strLinea = strLinea & "N" ' Campo 4
    strLinea = strLinea & "N" ' Campo 5
    strLinea = strLinea & "1" ' Campo 6
    strLinea = strLinea & String(15, "0") ' Campo 7
    strLinea = strLinea & String(15, "0") ' Campo 8
    strLinea = strLinea & String(15, "0") ' Campo 9
    strLinea = strLinea & String(15, "0") ' Campo 10
    strLinea = strLinea & String(15, "0") ' Campo 11
    strLinea = strLinea & String(15, "0") ' Campo 12
    
    Open path & "REGINFO_CV_CABECERA.txt" For Output As #1
    Print #1, strLinea
    Close #1
    pStatus.SimpleText = ". . . CABECERA Generada"
    
End Sub

Private Sub detalleVentas(path As String, desde As Date, hasta As Date, pDBs As Collection, pStatus As StatusBar, pProgreso As ProgressBar)
Dim colVen As New Collection

Dim objDB As clsDB
Dim objEmp As New clsDAOEmpresa

    Open path & "REGINFO_CV_VENTAS_CBTE.txt" For Output As #1
    Open path & "REGINFO_CV_VENTAS_ALICUOTAS.txt" For Output As #2
    Open path & "ERRORES_CV_VENTAS.txt" For Output As #3
    
    For Each objDB In pDBs
        objEmp.findLast objDB
        pStatus.SimpleText = objEmp.ubicacion & " - Generando VENTAS . . ."

        detalleVentasNegocio objEmp.ubicacion, desde, hasta, objDB, pProgreso, colVen

        pStatus.SimpleText = ". . . " & objEmp.ubicacion & " - VENTAS Generada"
    Next
    
    Close #1
    Close #2
    Close #3
    
End Sub

Private Sub detalleVentasNegocio(ubicacion As String, desde As Date, hasta As Date, pDB As clsDB, pProgreso As ProgressBar, pCollection As Collection)
Dim objMCl As New clsDAOMovclie

Dim objTC As New clsDAOTiposComprob
Dim objCA As New clsDAOCompAfip
Dim objCli As New clsDAOClientes

Dim svcNDb As New clsSvcNDebito

Dim colMov As Collection

Dim strLinea As String
Dim strDocu As String
Dim strAlic As String

Dim intTipo As Integer
Dim intCantAlic As Integer

Dim lngNroComprob As Long

On Error Resume Next

    For Each objMCl In objMCl.collectionByLibroIVA(desde, hasta, pDB)
        svcNDb.autoCorrectNDebito objMCl.negID, objMCl.cgocomprob, objMCl.prefijo, objMCl.nroComprob, pDB
    Next

    Set colMov = objMCl.collectionByLibroIVA(desde, hasta, pDB)
    pProgreso.Min = 1
    pProgreso.Max = colMov.Count + 1
    pProgreso.Value = pProgreso.Min
    
    For Each objMCl In colMov
        pProgreso.Value = pProgreso.Value + 1
        
        DoEvents
        
        If Abs(objMCl.montoiva) > 0 Then
            intCantAlic = 0
            With objTC
                If .codigo <> objMCl.cgocomprob Then
                    .negID = objMCl.negID
                    .codigo = objMCl.cgocomprob
                    .findByPrimaryKey pDB
                    
                    objCA.comprobanteID = .tipoafip
                    objCA.findByPrimaryKey pDB
                    If objCA.comprobanteID = 0 Or Not objCA.exist(pDB) Then
                        Print #3, "ERROR: Ubicación " & ubicacion & " - Comprobante " & .descripcion & " (" & .codigo & ") SIN Tipo AFIP Asociado!!"
                    End If
                End If
            End With
            With objCli
                If .codigo <> objMCl.cgoclie Then .findByCodigo objMCl.cgoclie, pDB
                
                intTipo = 80
                strDocu = Trim(Replace(.cuit, "-", ""))
                
                If .posicion = 2 Or Val(strDocu) = 0 Then
                    intTipo = 96
                    strDocu = Trim(.nrodoc)
                End If
                
                If Val(strDocu) = 0 Then
                    intTipo = 99
                    strDocu = "0"
                End If
                
                If Abs(objMCl.importe) > 1000 And strDocu = "0" Then
                    intTipo = 96
                    strDocu = "99"
                End If
                
                If intTipo = 80 Then
                    If Not modCUIT.validaCUIT(strDocu) Then
                        Print #3, "ERROR: Ubicación " & ubicacion & " - Cliente " & .razon & " CUIT " & .cuit & " NO Válida"
                    End If
                End If
            End With
            
            With objMCl
                lngNroComprob = .nroComprob
                
                If Not modCollection.collectionExistElement(pCollection, "k." & objCA.comprobanteID & "." & .prefijo & "." & lngNroComprob) Then
                    pCollection.add objMCl, "k." & objCA.comprobanteID & "." & .prefijo & "." & lngNroComprob
                Else
                    Print #3, "ERROR: Ubicación " & ubicacion & " - Comprobante " & objTC.descripcion & " " & .prefijo & "/" & .nroComprob & " REPETIDO . . ."
                    Do While modCollection.collectionExistElement(pCollection, "k." & objCA.comprobanteID & "." & .prefijo & "." & lngNroComprob)
                        lngNroComprob = lngNroComprob + 10000000
                    Loop
                    pCollection.add objMCl, "k." & objCA.comprobanteID & "." & .prefijo & "." & lngNroComprob
                End If
                
                strLinea = Format(Year(.fechacomprob), "0000") & Format(Month(.fechacomprob), "00") & Format(Day(.fechacomprob), "00") ' Campo 1
                strLinea = strLinea & Format(objCA.comprobanteID, "000") ' Campo 2
                strLinea = strLinea & Format(.prefijo, "00000") ' Campo 3
                strLinea = strLinea & Format(lngNroComprob, String(20, "0")) ' Campo 4
                strLinea = strLinea & Format(lngNroComprob, String(20, "0")) ' Campo 5
                strLinea = strLinea & Format(intTipo, "00") ' Campo 6
                strLinea = strLinea & Format(strDocu, String(20, "0")) ' Campo 7
                strLinea = strLinea & Left(modConv.replaceSymbols(objCli.razon) & Space(30), 30) ' Campo 8
                strLinea = strLinea & Format(Abs(.importe) * 100, String(15, "0")) ' Campo 9
                strLinea = strLinea & String(15, "0") ' Campo 10
                strLinea = strLinea & String(15, "0") ' Campo 11
                strLinea = strLinea & Format(Abs(.montoexento) * 100, String(15, "0")) ' Campo 12
                strLinea = strLinea & String(15, "0") ' Campo 13
                strLinea = strLinea & String(15, "0") ' Campo 14
                strLinea = strLinea & String(15, "0") ' Campo 15
                strLinea = strLinea & String(15, "0") ' Campo 16
                strLinea = strLinea & "PES" ' Campo 17
                strLinea = strLinea & "0001000000" ' Campo 18
                If Abs(.montoiva) > 0 Then intCantAlic = intCantAlic + 1
                strLinea = strLinea & Trim(Str(intCantAlic)) ' Campo 19
                strLinea = strLinea & " " ' Campo 20
                strLinea = strLinea & String(15, "0") ' Campo 21
                strLinea = strLinea & Format(Year(.fechacomprob), "0000") & Format(Month(.fechacomprob), "00") & Format(Day(.fechacomprob), "00") ' Campo 22
                Print #1, strLinea
                
                If Abs(.montoiva) > 0 Then
                    If .iva = 21 Then
                        strAlic = Format(objCA.comprobanteID, "000")  ' Campo 1
                        strAlic = strAlic & Format(.prefijo, "00000") ' Campo 2
                        strAlic = strAlic & Format(lngNroComprob, String(20, "0")) ' Campo 3
                        strAlic = strAlic & Format(Abs(.montoiva / 0.21) * 100, String(15, "0")) ' Campo 4
                        strAlic = strAlic & "0005" ' Campo 5
                        strAlic = strAlic & Format(Abs(.montoiva) * 100, String(15, "0")) ' Campo 6
                        Print #2, strAlic
                    End If
                    If .iva = 10.5 Then
                        strAlic = Format(objCA.comprobanteID, "000")  ' Campo 1
                        strAlic = strAlic & Format(.prefijo, "00000") ' Campo 2
                        strAlic = strAlic & Format(lngNroComprob, String(20, "0")) ' Campo 3
                        strAlic = strAlic & Format(Abs(.montoiva / 0.105) * 100, String(15, "0")) ' Campo 4
                        strAlic = strAlic & "0004" ' Campo 5
                        strAlic = strAlic & Format(Abs(.montoiva) * 100, String(15, "0")) ' Campo 6
                        Print #2, strAlic
                    End If
                End If
            End With
        End If
    Next
    
End Sub

