VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOMovclie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vNegid As Integer
Private vEmpID As Integer
Private vCgoClie As Long
Private vCgocomprob As Integer
Private vFechaComprob As Date
Private vFechaVenc As Date
Private vPrefijo As Integer
Private vNroComprob As Long
Private vImporte As Currency
Private vCancelado As Currency
Private vNetoSinDescuento As Currency
Private vNeto As Currency
Private vNetoCancelado As Currency
Private vMontoIva As Currency
Private vMontoExento As Currency
Private vRecibo As Integer
Private vAnulada As Integer
Private vTipocompro As String
Private vLetras As String
Private vIva As Currency
Private vObservaciones As String
Private vCae As String
Private vCaeVenc As String
Private vBarras As String

Private clon As clsDAOMovclie

Public ddl As New clsDDLMovclie

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegid = gEmpresa.negID
    vEmpID = 1
    vCgoClie = 0
    vCgocomprob = 0
    vFechaComprob = Date
    vFechaVenc = Date
    vPrefijo = 0
    vNroComprob = 0
    vImporte = 0
    vCancelado = 0
    vNetoSinDescuento = 0
    vNeto = 0
    vNetoCancelado = 0
    vMontoIva = 0
    vMontoExento = 0
    vRecibo = 0
    vAnulada = 0
    vTipocompro = ""
    vLetras = ""
    vIva = 0
    vObservaciones = ""
    vCae = ""
    vCaeVenc = ""
    vBarras = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get negID() As Variant

    negID = vNegid

End Property

Public Property Let negID(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get empID() As Variant

    empID = vEmpID

End Property

Public Property Let empID(ByVal vNewValue As Variant)

On Error Resume Next

    vEmpID = vNewValue

End Property

Public Property Get cgoclie() As Variant

    cgoclie = vCgoClie

End Property

Public Property Let cgoclie(ByVal vNewValue As Variant)

On Error Resume Next

    vCgoClie = vNewValue

End Property

Public Property Get cgocomprob() As Variant

    cgocomprob = vCgocomprob

End Property

Public Property Let cgocomprob(ByVal vNewValue As Variant)

On Error Resume Next

    vCgocomprob = vNewValue

End Property

Public Property Get fechacomprob() As Variant

    fechacomprob = vFechaComprob

End Property

Public Property Let fechacomprob(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaComprob = vNewValue

End Property

Public Property Get fechavenc() As Variant

    fechavenc = vFechaVenc

End Property

Public Property Let fechavenc(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaVenc = vNewValue

End Property

Public Property Get prefijo() As Variant

    prefijo = vPrefijo

End Property

Public Property Let prefijo(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijo = vNewValue

End Property

Public Property Get nroComprob() As Variant

    nroComprob = vNroComprob

End Property

Public Property Let nroComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get cancelado() As Variant

    cancelado = vCancelado

End Property

Public Property Let cancelado(ByVal vNewValue As Variant)

On Error Resume Next

    vCancelado = vNewValue

End Property

Public Property Get netosindescuento() As Variant

    netosindescuento = vNetoSinDescuento

End Property

Public Property Let netosindescuento(ByVal vNewValue As Variant)

On Error Resume Next

    vNetoSinDescuento = vNewValue

End Property

Public Property Get neto() As Variant

    neto = vNeto

End Property

Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

Public Property Get netocancelado() As Variant

    netocancelado = vNetoCancelado

End Property

Public Property Let netocancelado(ByVal vNewValue As Variant)

On Error Resume Next

    vNetoCancelado = vNewValue

End Property

Public Property Get montoiva() As Variant

    montoiva = vMontoIva

End Property

Public Property Let montoiva(ByVal vNewValue As Variant)

On Error Resume Next

    vMontoIva = vNewValue

End Property

Public Property Get montoexento() As Variant

    montoexento = vMontoExento

End Property

Public Property Let montoexento(ByVal vNewValue As Variant)

On Error Resume Next

    vMontoExento = vNewValue

End Property

Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

Public Property Get anulada() As Variant

    anulada = vAnulada

End Property

Public Property Let anulada(ByVal vNewValue As Variant)

On Error Resume Next

    vAnulada = vNewValue

End Property

Public Property Get tipocompro() As Variant

    tipocompro = vTipocompro

End Property

Public Property Let tipocompro(ByVal vNewValue As Variant)

On Error Resume Next

    vTipocompro = vNewValue

End Property

Public Property Get letras() As Variant

    letras = vLetras

End Property

Public Property Let letras(ByVal vNewValue As Variant)

On Error Resume Next

    vLetras = vNewValue

End Property

Public Property Get iva() As Variant

    iva = vIva

End Property

Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

Public Property Get observaciones() As Variant

    observaciones = vObservaciones

End Property

Public Property Let observaciones(ByVal vNewValue As Variant)

On Error Resume Next

    vObservaciones = vNewValue

End Property

Public Property Get cae() As Variant

    cae = vCae

End Property

Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

Public Property Get caevenc() As Variant

    caevenc = vCaeVenc

End Property

Public Property Let caevenc(ByVal vNewValue As Variant)

On Error Resume Next

    vCaeVenc = vNewValue

End Property

Public Property Get barras() As Variant

    barras = vBarras

End Property

Public Property Let barras(ByVal vNewValue As Variant)

On Error Resume Next

    vBarras = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vNegid

End Property

Public Property Get comboKey() As Variant

    comboKey = vClave

End Property

Public Property Get listText() As Variant

    listText = vNegid

End Property

Public Property Get listKey() As Variant

    listKey = vClave

End Property

Public Function clone() As clsDAOMovclie
Dim objObjeto As New clsDAOMovclie

    With objObjeto
        .clave = vClave
        .negID = vNegid
        .empID = vEmpID
        .cgoclie = vCgoClie
        .cgocomprob = vCgocomprob
        .fechacomprob = vFechaComprob
        .fechavenc = vFechaVenc
        .prefijo = vPrefijo
        .nroComprob = vNroComprob
        .importe = vImporte
        .cancelado = vCancelado
        .netosindescuento = vNetoSinDescuento
        .neto = vNeto
        .netocancelado = vNetoCancelado
        .montoiva = vMontoIva
        .montoexento = vMontoExento
        .recibo = vRecibo
        .anulada = vAnulada
        .tipocompro = vTipocompro
        .letras = vLetras
        .iva = vIva
        .observaciones = vObservaciones
        .cae = vCae
        .caevenc = vCaeVenc
        .barras = vBarras
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOMovclie, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .negID = pQuery.Fields(.ddl.negID)
        .empID = pQuery.Fields(.ddl.empID)
        .cgoclie = pQuery.Fields(.ddl.cgoclie)
        .cgocomprob = pQuery.Fields(.ddl.cgocomprob)
        .fechacomprob = pQuery.Fields(.ddl.fechacomprob)
        .fechavenc = pQuery.Fields(.ddl.fechavenc)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .importe = pQuery.Fields(.ddl.importe)
        .cancelado = pQuery.Fields(.ddl.cancelado)
        .netosindescuento = pQuery.Fields(.ddl.netosindescuento)
        .neto = pQuery.Fields(.ddl.neto)
        .netocancelado = pQuery.Fields(.ddl.netocancelado)
        .montoiva = pQuery.Fields(.ddl.montoiva)
        .montoexento = pQuery.Fields(.ddl.montoexento)
        .recibo = pQuery.Fields(.ddl.recibo)
        .anulada = pQuery.Fields(.ddl.anulada)
        .tipocompro = pQuery.Fields(.ddl.tipocompro)
        .letras = pQuery.Fields(.ddl.letras)
        .iva = pQuery.Fields(.ddl.iva)
        .observaciones = pQuery.Fields(.ddl.observaciones)
        .cae = pQuery.Fields(.ddl.cae)
        .caevenc = pQuery.Fields(.ddl.caevenc)
        .barras = pQuery.Fields(.ddl.barras)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.clave & " = " & vClave

    makeWherePrimaryKey = strSQL

End Function

Private Sub findSQL(pSQL As String, pDB As clsDB)
Dim rstQuery As ADODB.Recordset

    Set rstQuery = pDB.query(pSQL)

    If Not rstQuery.EOF Then
        fillObject Me, rstQuery

        Set clon = Me.clone
    End If

    rstQuery.Close

End Sub

Public Sub findByPrimaryKey(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByComprobante(negocioID As Integer, tipoComprobanteID As Integer, prefijo As Integer, nroComprob As Long, pDB As clsDB)
Dim strSQL As String

    clean
    
    vNegid = negocioID
    vCgocomprob = tipoComprobanteID
    vPrefijo = prefijo
    vNroComprob = nroComprob

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.negID & " = " & vNegid
    strSQL = strSQL & " AND " & ddl.cgocomprob & " = " & vCgocomprob
    strSQL = strSQL & " AND " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & " AND " & ddl.nroComprob & " = " & vNroComprob
    strSQL = strSQL & ";"

    findSQL strSQL, pDB
    
End Sub

Public Sub findLastID(pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    clean

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.clave
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vClave = rstQuery.Fields(ddl.clave)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

Public Sub findLast(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.clave & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Function delete(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete = pDB.execute(strSQL)

End Function

Public Function save(pDB As clsDB) As Boolean

    If exist(pDB) Then
        save = update(pDB)
    Else
        save = add(pDB)
    End If

End Function

Public Function exist(pDB As clsDB) As Boolean
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.negID & " = " & vNegid
    strSQL = strSQL & ", " & ddl.empID & " = " & vEmpID
    strSQL = strSQL & ", " & ddl.cgoclie & " = " & vCgoClie
    strSQL = strSQL & ", " & ddl.cgocomprob & " = " & vCgocomprob
    strSQL = strSQL & ", " & ddl.fechacomprob & " = " & fechaDB(vFechaComprob)
    strSQL = strSQL & ", " & ddl.fechavenc & " = " & fechaDB(vFechaVenc)
    strSQL = strSQL & ", " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & ", " & ddl.nroComprob & " = " & vNroComprob
    strSQL = strSQL & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.cancelado & " = '" & Replace(vCancelado, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.netosindescuento & " = '" & Replace(vNetoSinDescuento, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.netocancelado & " = '" & Replace(vNetoCancelado, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.montoiva & " = '" & Replace(vMontoIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.montoexento & " = '" & Replace(vMontoExento, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.recibo & " = " & vRecibo
    strSQL = strSQL & ", " & ddl.anulada & " = " & vAnulada
    strSQL = strSQL & ", " & ddl.tipocompro & " = '" & replaceSpecialSymbols(vTipocompro) & "'"
    strSQL = strSQL & ", " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
    strSQL = strSQL & ", " & ddl.iva & " = '" & Replace(vIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
    strSQL = strSQL & ", " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
    strSQL = strSQL & ", " & ddl.caevenc & " = '" & replaceSpecialSymbols(vCaeVenc) & "'"
    strSQL = strSQL & ", " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
    strSQL = strSQL & ";"

    add = pDB.execute(strSQL)

    findLastID pDB

End Function

Public Function update(pDB As clsDB) As Boolean
Dim strSQL As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vNegid <> clon.negID Or IsNull(vNegid) <> IsNull(clon.negID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negID & " = " & vNegid
        blnChange = True
    End If
    If vEmpID <> clon.empID Or IsNull(vEmpID) <> IsNull(clon.empID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.empID & " = " & vEmpID
        blnChange = True
    End If
    If vCgoClie <> clon.cgoclie Or IsNull(vCgoClie) <> IsNull(clon.cgoclie) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgoclie & " = " & vCgoClie
        blnChange = True
    End If
    If vCgocomprob <> clon.cgocomprob Or IsNull(vCgocomprob) <> IsNull(clon.cgocomprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgocomprob & " = " & vCgocomprob
        blnChange = True
    End If
    If vFechaComprob <> clon.fechacomprob Or IsNull(vFechaComprob) <> IsNull(clon.fechacomprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechacomprob & " = " & fechaDB(vFechaComprob)
        blnChange = True
    End If
    If vFechaVenc <> clon.fechavenc Or IsNull(vFechaVenc) <> IsNull(clon.fechavenc) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechavenc & " = " & fechaDB(vFechaVenc)
        blnChange = True
    End If
    If vPrefijo <> clon.prefijo Or IsNull(vPrefijo) <> IsNull(clon.prefijo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.prefijo & " = " & vPrefijo
        blnChange = True
    End If
    If vNroComprob <> clon.nroComprob Or IsNull(vNroComprob) <> IsNull(clon.nroComprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroComprob & " = " & vNroComprob
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vCancelado <> clon.cancelado Or IsNull(vCancelado) <> IsNull(clon.cancelado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cancelado & " = '" & Replace(vCancelado, ",", ".") & "'"
        blnChange = True
    End If
    If vNetoSinDescuento <> clon.netosindescuento Or IsNull(vNetoSinDescuento) <> IsNull(clon.netosindescuento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.netosindescuento & " = '" & Replace(vNetoSinDescuento, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto <> clon.neto Or IsNull(vNeto) <> IsNull(clon.neto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
        blnChange = True
    End If
    If vNetoCancelado <> clon.netocancelado Or IsNull(vNetoCancelado) <> IsNull(clon.netocancelado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.netocancelado & " = '" & Replace(vNetoCancelado, ",", ".") & "'"
        blnChange = True
    End If
    If vMontoIva <> clon.montoiva Or IsNull(vMontoIva) <> IsNull(clon.montoiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.montoiva & " = '" & Replace(vMontoIva, ",", ".") & "'"
        blnChange = True
    End If
    If vMontoExento <> clon.montoexento Or IsNull(vMontoExento) <> IsNull(clon.montoexento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.montoexento & " = '" & Replace(vMontoExento, ",", ".") & "'"
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vAnulada <> clon.anulada Or IsNull(vAnulada) <> IsNull(clon.anulada) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.anulada & " = " & vAnulada
        blnChange = True
    End If
    If vTipocompro <> clon.tipocompro Or IsNull(vTipocompro) <> IsNull(clon.tipocompro) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipocompro & " = '" & replaceSpecialSymbols(vTipocompro) & "'"
        blnChange = True
    End If
    If vLetras <> clon.letras Or IsNull(vLetras) <> IsNull(clon.letras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.letras & " = '" & replaceSpecialSymbols(vLetras) & "'"
        blnChange = True
    End If
    If vIva <> clon.iva Or IsNull(vIva) <> IsNull(clon.iva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva & " = '" & Replace(vIva, ",", ".") & "'"
        blnChange = True
    End If
    If vObservaciones <> clon.observaciones Or IsNull(vObservaciones) <> IsNull(clon.observaciones) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.observaciones & " = '" & replaceSpecialSymbols(vObservaciones) & "'"
        blnChange = True
    End If
    If vCae <> clon.cae Or IsNull(vCae) <> IsNull(clon.cae) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
        blnChange = True
    End If
    If vCaeVenc <> clon.caevenc Or IsNull(vCaeVenc) <> IsNull(clon.caevenc) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caevenc & " = '" & replaceSpecialSymbols(vCaeVenc) & "'"
        blnChange = True
    End If
    If vBarras <> clon.barras Or IsNull(vBarras) <> IsNull(clon.barras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
        blnChange = True
    End If

    strSQL = "UPDATE " & ddl.table & " SET" & strSET
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    If blnChange Then update = pDB.execute(strSQL)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    vClave = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.clave & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vClave = 1 + rstQuery.Fields(ddl.clave)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByCliID(pCliID As Long, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoclie & " = " & pCliID
    strSQL = strSQL & " AND " & ddl.anulada & " = 0"
    strSQL = strSQL & ";"

    Set collectionByCliID = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByLibroIVA(pDesde As Date, pHasta As Date, pDB As clsDB) As Collection
Dim strSQL As String

Dim tc As New clsDDLTiposComprob

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " INNER JOIN " & tc.table
    strSQL = strSQL & " ON " & tc.codigo_t & " = " & ddl.cgocomprob_t
    strSQL = strSQL & " AND " & tc.libroiva_t & " <> 0"
    strSQL = strSQL & " WHERE " & ddl.fechacomprob_t & " BETWEEN " & fechaDB(pDesde) & " AND " & fechaDB(pHasta)
    strSQL = strSQL & " AND " & ddl.anulada & " = 0"
    strSQL = strSQL & " ORDER BY " & ddl.fechacomprob_t & ", " & ddl.prefijo_t
    strSQL = strSQL & ";"

    Set collectionByLibroIVA = collectionSQL(strSQL, pDB)

End Function

Public Function collectionAnulables(pCliID As Long, pDesde As Date, pHasta As Date, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoclie & " = " & pCliID
    strSQL = strSQL & " AND " & ddl.anulada & " = 0"
    strSQL = strSQL & " AND " & ddl.fechacomprob & " BETWEEN " & fechaDB(pDesde) & " AND " & fechaDB(pHasta)
    strSQL = strSQL & " ORDER BY " & ddl.fechacomprob
    strSQL = strSQL & ";"
    
    Set collectionAnulables = collectionSQL(strSQL, pDB)
    
End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOMovclie

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOMovclie
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = colRegistros

End Function

Public Sub fillCombo(pCombo As ComboBox, pDB As clsDB)

    fillCombo_ pCombo, collectionAll(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOMovclie

    pCombo.Clear

    For Each objElemento In pCollection
        pCombo.AddItem objElemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = objElemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, pDB As clsDB)

    fillList_ pList, collectionAll(pDB)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim objElemento As clsDAOMovclie

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function buscaNumeroRecibo(pPrefijo As Integer, pTipoCompro As String, pCgoComprob As Integer, pDB As clsDB) As Long
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    buscaNumeroRecibo = 1
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    If pTipoCompro <> "Z" Then
        strSQL = strSQL & " WHERE " & ddl.recibo & " <> 0"
        strSQL = strSQL & " AND " & ddl.prefijo & " = " & pPrefijo
    Else
        strSQL = strSQL & " WHERE " & ddl.cgocomprob & " = " & pCgoComprob
        strSQL = strSQL & " AND " & ddl.prefijo & " = 0"
    End If
    strSQL = strSQL & " ORDER BY " & ddl.nroComprob & " DESC"
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then buscaNumeroRecibo = 1 + rstQuery.Fields(ddl.nroComprob)
    rstQuery.Close
    
End Function

Public Function buscaNumeroFactura(pPrefijo As Integer, pTipoCompro As String, pDB As clsDB) As Long
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    buscaNumeroFactura = 1
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.recibo & " = 0"
    strSQL = strSQL & " AND " & ddl.prefijo & " = " & pPrefijo
    strSQL = strSQL & " AND " & ddl.tipocompro & " = '" & pTipoCompro & "'"
    strSQL = strSQL & " ORDER BY " & ddl.nroComprob & " DESC"
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then buscaNumeroFactura = 1 + rstQuery.Fields(ddl.nroComprob)
    rstQuery.Close
    
End Function

Public Sub extremosFechaComprob(pDesde As Date, pHasta As Date, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    pDesde = Date
    pHasta = Date

    strSQL = "SELECT MIN(fechacomprob) AS minimo, MAX(fechacomprob) AS maximo FROM " & ddl.table
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then
        pDesde = rstQuery!minimo
        pHasta = rstQuery!maximo
    End If
    rstQuery.Close
    
End Sub

Public Sub extremosNroComprob(pMinimo As Long, pMaximo As Long, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    pMinimo = 0
    pMaximo = 0

    strSQL = "SELECT MIN(nrocomprob) AS minimo, MAX(nrocomprob) AS maximo FROM " & ddl.table
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then
        pMinimo = rstQuery!minimo
        pMaximo = rstQuery!maximo
    End If
    rstQuery.Close
    
End Sub

Public Function saldoImporteHasta(pCliID As Integer, pFechaTope As Date, pDB As clsDB) As Currency
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    saldoImporteHasta = 0

    strSQL = "SELECT SUM(" & ddl.importe & ") AS debitos FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoclie & " = " & pCliID
    strSQL = strSQL & " AND " & ddl.fechacomprob & " < " & fechaDB(pFechaTope)
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then If Not IsNull(rstQuery!debitos) Then saldoImporteHasta = rstQuery!debitos
    rstQuery.Close
    
End Function
