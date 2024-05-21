VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vCodigo As String
Private vDescripcion As String
Private vPrecioventaconiva As Currency
Private vPrecioventasiniva As Currency
Private vPreciolistaconiva As Currency
Private vPrecioListaSinIva As Currency
Private vPreciocomprasiniva As Currency
Private vPreciocomprasinivaanterior As Currency
Private vIva105 As Integer
Private vExento As Integer
Private vModelocamion As String
Private vFechaActualizacion As Date
Private vOrigen As String
Private vDescuento As String
Private vPrvID As Long
Private vUltimacompra As Date
Private vMarca As String
Private vCatalogo As String
Private vPreciolistasinivaUSD As Currency
Private vCotizacionID As Long

Private clon As clsDAOArticulos

Public ddl As New clsDDLArticulos

Private Sub class_Initialize()

    vCodigo = ""

    initializeOther

End Sub

Private Sub initializeOther()

    vDescripcion = ""
    vPrecioventaconiva = 0
    vPrecioventasiniva = 0
    vPreciolistaconiva = 0
    vPrecioListaSinIva = 0
    vPreciocomprasiniva = 0
    vPreciocomprasinivaanterior = 0
    vIva105 = 0
    vExento = 0
    vModelocamion = ""
    vFechaActualizacion = Date
    vOrigen = ""
    vDescuento = ""
    vPrvID = 0
    vUltimacompra = Date
    vMarca = ""
    vCatalogo = ""
    vPreciolistasinivaUSD = 0
    vCotizacionID = 0
    vClave = 0

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

Public Property Get codigo() As Variant

    codigo = vCodigo

End Property

Public Property Let codigo(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigo = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get precioventaconiva() As Variant

    precioventaconiva = vPrecioventaconiva

End Property

Public Property Let precioventaconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioventaconiva = vNewValue

End Property

Public Property Get precioventasiniva() As Variant

    precioventasiniva = vPrecioventasiniva

End Property

Public Property Let precioventasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioventasiniva = vNewValue

End Property

Public Property Get preciolistaconiva() As Variant

    preciolistaconiva = vPreciolistaconiva

End Property

Public Property Let preciolistaconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciolistaconiva = vNewValue

End Property

Public Property Get preciolistasiniva() As Variant

    preciolistasiniva = vPrecioListaSinIva

End Property

Public Property Let preciolistasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioListaSinIva = vNewValue

End Property

Public Property Get preciocomprasiniva() As Variant

    preciocomprasiniva = vPreciocomprasiniva

End Property

Public Property Let preciocomprasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciocomprasiniva = vNewValue

End Property

Public Property Get preciocomprasinivaanterior() As Variant

    preciocomprasinivaanterior = vPreciocomprasinivaanterior

End Property

Public Property Let preciocomprasinivaanterior(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciocomprasinivaanterior = vNewValue

End Property

Public Property Get iva105() As Variant

    iva105 = vIva105

End Property

Public Property Let iva105(ByVal vNewValue As Variant)

On Error Resume Next

    vIva105 = vNewValue

End Property

Public Property Get exento() As Variant

    exento = vExento

End Property

Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

Public Property Get modelocamion() As Variant

    modelocamion = vModelocamion

End Property

Public Property Let modelocamion(ByVal vNewValue As Variant)

On Error Resume Next

    vModelocamion = vNewValue

End Property

Public Property Get fechaactualizacion() As Variant

    fechaactualizacion = vFechaActualizacion

End Property

Public Property Let fechaactualizacion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaActualizacion = vNewValue

End Property

Public Property Get origen() As Variant

    origen = vOrigen

End Property

Public Property Let origen(ByVal vNewValue As Variant)

On Error Resume Next

    vOrigen = vNewValue

End Property

Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

Public Property Get prvID() As Variant

    prvID = vPrvID

End Property

Public Property Let prvID(ByVal vNewValue As Variant)

On Error Resume Next

    vPrvID = vNewValue

End Property

Public Property Get ultimacompra() As Variant

    ultimacompra = vUltimacompra

End Property

Public Property Let ultimacompra(ByVal vNewValue As Variant)

On Error Resume Next

    vUltimacompra = vNewValue

End Property

Public Property Get marca() As Variant

    marca = vMarca

End Property

Public Property Let marca(ByVal vNewValue As Variant)

On Error Resume Next

    vMarca = vNewValue

End Property

Public Property Get catalogo() As Variant

    catalogo = vCatalogo

End Property

Public Property Let catalogo(ByVal vNewValue As Variant)

On Error Resume Next

    vCatalogo = vNewValue

End Property

Public Property Get preciolistasinivausd() As Variant

    preciolistasinivausd = vPreciolistasinivaUSD

End Property

Public Property Let preciolistasinivausd(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciolistasinivaUSD = vNewValue

End Property

Public Property Get cotizacionID() As Variant

    cotizacionID = vCotizacionID

End Property

Public Property Let cotizacionID(ByVal vNewValue As Variant)

On Error Resume Next

    vCotizacionID = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vDescripcion & " (" & vCodigo & ")"

End Property

Public Property Get comboKey() As Variant

    comboKey = vClave

End Property

Public Property Get listText() As Variant

    listText = vDescripcion & " (" & vCodigo & ")"

End Property

Public Property Get listKey() As Variant

    listKey = vClave

End Property

Public Function clone() As clsDAOArticulos
Dim objObjeto As New clsDAOArticulos

    With objObjeto
        .clave = vClave
        .codigo = vCodigo
        .descripcion = vDescripcion
        .precioventaconiva = vPrecioventaconiva
        .precioventasiniva = vPrecioventasiniva
        .preciolistaconiva = vPreciolistaconiva
        .preciolistasiniva = vPrecioListaSinIva
        .preciocomprasiniva = vPreciocomprasiniva
        .preciocomprasinivaanterior = vPreciocomprasinivaanterior
        .iva105 = vIva105
        .exento = vExento
        .modelocamion = vModelocamion
        .fechaactualizacion = vFechaActualizacion
        .origen = vOrigen
        .descuento = vDescuento
        .prvID = vPrvID
        .ultimacompra = vUltimacompra
        .marca = vMarca
        .catalogo = vCatalogo
        .preciolistasinivausd = vPreciolistasinivaUSD
        .cotizacionID = vCotizacionID
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOArticulos, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .codigo = pQuery.Fields(.ddl.codigo)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .precioventaconiva = pQuery.Fields(.ddl.precioventaconiva)
        .precioventasiniva = pQuery.Fields(.ddl.precioventasiniva)
        .preciolistaconiva = pQuery.Fields(.ddl.preciolistaconiva)
        .preciolistasiniva = pQuery.Fields(.ddl.preciolistasiniva)
        .preciocomprasiniva = pQuery.Fields(.ddl.preciocomprasiniva)
        .preciocomprasinivaanterior = pQuery.Fields(.ddl.preciocomprasinivaanterior)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .exento = pQuery.Fields(.ddl.exento)
        .modelocamion = pQuery.Fields(.ddl.modelocamion)
        .fechaactualizacion = pQuery.Fields(.ddl.fechaactualizacion)
        .origen = pQuery.Fields(.ddl.origen)
        .descuento = pQuery.Fields(.ddl.descuento)
        .prvID = pQuery.Fields(.ddl.prvID)
        .ultimacompra = pQuery.Fields(.ddl.ultimacompra)
        .marca = pQuery.Fields(.ddl.marca)
        .catalogo = pQuery.Fields(.ddl.catalogo)
        .preciolistasinivausd = pQuery.Fields(.ddl.preciolistasinivausd)
        .cotizacionID = pQuery.Fields(.ddl.cotizacionID)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.codigo & " = '" & vCodigo & "'"

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

Public Sub findByClave(pClave As Long, pDB As clsDB)
Dim strSQL As String

    clean
    
    vClave = pClave
    
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.clave & " = " & vClave
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findLastID(Optional pDB As Variant)

    findLastID_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findLastID_(pDB As clsDB)
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

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.clave & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Function delete(Optional pDB As Variant) As Boolean

    delete = delete_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function delete_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete_ = pDB.execute(strSQL)

End Function

Public Function save(Optional pDB As Variant) As Boolean

    save = save_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function save_(pDB As clsDB) As Boolean

    If exist_(pDB) Then
        save_ = update_(pDB)
    Else
        save_ = add_(pDB)
    End If

End Function

Public Function exist(Optional pDB As Variant) As Boolean

    exist = exist_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function exist_(pDB As clsDB) As Boolean
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)
    exist_ = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(Optional pDB As Variant) As Boolean

    add = add_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function add_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.codigo & " = '" & replaceSpecialSymbols(vCodigo) & "'"
    strSQL = strSQL & ", " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    strSQL = strSQL & ", " & ddl.precioventaconiva & " = '" & Replace(vPrecioventaconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.precioventasiniva & " = '" & Replace(vPrecioventasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciolistaconiva & " = '" & Replace(vPreciolistaconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciolistasiniva & " = '" & Replace(vPrecioListaSinIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciocomprasinivaanterior & " = '" & Replace(vPreciocomprasinivaanterior, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.iva105 & " = " & vIva105
    strSQL = strSQL & ", " & ddl.exento & " = " & vExento
    strSQL = strSQL & ", " & ddl.modelocamion & " = '" & replaceSpecialSymbols(vModelocamion) & "'"
    strSQL = strSQL & ", " & ddl.fechaactualizacion & " = " & modDB.fechaDB(vFechaActualizacion)
    strSQL = strSQL & ", " & ddl.origen & " = '" & replaceSpecialSymbols(vOrigen) & "'"
    strSQL = strSQL & ", " & ddl.descuento & " = '" & replaceSpecialSymbols(vDescuento) & "'"
    strSQL = strSQL & ", " & ddl.prvID & " = " & vPrvID
    strSQL = strSQL & ", " & ddl.ultimacompra & " = " & fechaDB(vUltimacompra)
    strSQL = strSQL & ", " & ddl.marca & " = '" & replaceSpecialSymbols(vMarca) & "'"
    strSQL = strSQL & ", " & ddl.catalogo & " = '" & replaceSpecialSymbols(vCatalogo) & "'"
    strSQL = strSQL & ", " & ddl.preciolistasinivausd & " = '" & Replace(vPreciolistasinivaUSD, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.cotizacionID & " = " & vCotizacionID
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findByPrimaryKey pDB

End Function

Public Function update(Optional pDB As Variant) As Boolean

    update = update_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function update_(pDB As clsDB) As Boolean
Dim strSQL As String
Dim strSET As String

Dim blnChange As Boolean

    update_ = True

    strSET = ""
    blnChange = False

    If clon Is Nothing Then
        Set clon = Me.clone
        clon.findByPrimaryKey pDB
    End If

    If vDescripcion <> clon.descripcion Or IsNull(vDescripcion) <> IsNull(clon.descripcion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
        blnChange = True
    End If
    If vPrecioventaconiva <> clon.precioventaconiva Or IsNull(vPrecioventaconiva) <> IsNull(clon.precioventaconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioventaconiva & " = '" & Replace(vPrecioventaconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vPrecioventasiniva <> clon.precioventasiniva Or IsNull(vPrecioventasiniva) <> IsNull(clon.precioventasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioventasiniva & " = '" & Replace(vPrecioventasiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciolistaconiva <> clon.preciolistaconiva Or IsNull(vPreciolistaconiva) <> IsNull(clon.preciolistaconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciolistaconiva & " = '" & Replace(vPreciolistaconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vPrecioListaSinIva <> clon.preciolistasiniva Or IsNull(vPrecioListaSinIva) <> IsNull(clon.preciolistasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciolistasiniva & " = '" & Replace(vPrecioListaSinIva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciocomprasiniva <> clon.preciocomprasiniva Or IsNull(vPreciocomprasiniva) <> IsNull(clon.preciocomprasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciocomprasinivaanterior <> clon.preciocomprasinivaanterior Or IsNull(vPreciocomprasinivaanterior) <> IsNull(clon.preciocomprasinivaanterior) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciocomprasinivaanterior & " = '" & Replace(vPreciocomprasinivaanterior, ",", ".") & "'"
        blnChange = True
    End If
    If vIva105 <> clon.iva105 Or IsNull(vIva105) <> IsNull(clon.iva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva105 & " = " & vIva105
        blnChange = True
    End If
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = " & vExento
        blnChange = True
    End If
    If vModelocamion <> clon.modelocamion Or IsNull(vModelocamion) <> IsNull(clon.modelocamion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.modelocamion & " = '" & replaceSpecialSymbols(vModelocamion) & "'"
        blnChange = True
    End If
    If vFechaActualizacion <> clon.fechaactualizacion Or IsNull(vFechaActualizacion) <> IsNull(clon.fechaactualizacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaactualizacion & " = " & fechaDB(vFechaActualizacion)
        blnChange = True
    End If
    If vOrigen <> clon.origen Or IsNull(vOrigen) <> IsNull(clon.origen) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.origen & " = '" & replaceSpecialSymbols(vOrigen) & "'"
        blnChange = True
    End If
    If vDescuento <> clon.descuento Or IsNull(vDescuento) <> IsNull(clon.descuento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descuento & " = '" & replaceSpecialSymbols(vDescuento) & "'"
        blnChange = True
    End If
    If vPrvID <> clon.prvID Or IsNull(vPrvID) <> IsNull(clon.prvID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.prvID & " = " & vPrvID
        blnChange = True
    End If
    If vUltimacompra <> clon.ultimacompra Or IsNull(vUltimacompra) <> IsNull(clon.ultimacompra) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ultimacompra & " = " & fechaDB(vUltimacompra)
        blnChange = True
    End If
    If vMarca <> clon.marca Or IsNull(vMarca) <> IsNull(clon.marca) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.marca & " = '" & replaceSpecialSymbols(vMarca) & "'"
        blnChange = True
    End If
    If vCatalogo <> clon.catalogo Or IsNull(vCatalogo) <> IsNull(clon.catalogo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.catalogo & " = '" & replaceSpecialSymbols(vCatalogo) & "'"
        blnChange = True
    End If
    If vPreciolistasinivaUSD <> clon.preciolistasinivausd Or IsNull(vPreciolistasinivaUSD) <> IsNull(clon.preciolistasinivausd) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciolistasinivausd & " = '" & Replace(vPreciolistasinivaUSD, ",", ".") & "'"
        blnChange = True
    End If
    If vCotizacionID <> clon.cotizacionID Or IsNull(vCotizacionID) <> IsNull(clon.cotizacionID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cotizacionID & " = " & vCotizacionID
        blnChange = True
    End If

    strSQL = "UPDATE " & ddl.table & " SET" & strSET
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    If blnChange Then update_ = pDB.execute(strSQL)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByDescripcion(pCadena As String, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.descripcion & " LIKE '" & pCadena & "%'"
    strSQL = strSQL & " LIMIT 100"
    strSQL = strSQL & ";"

    Set collectionByDescripcion = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByProveedor(pPrvID As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.prvID & " = " & pPrvID
    strSQL = strSQL & " ORDER BY " & ddl.descripcion
    strSQL = strSQL & ";"

    Set collectionByProveedor = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOArticulos

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOArticulos
        fillObject objElemento, rstQuery

        colRegistros.add objElemento, "k." & objElemento.codigo

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Sub fillCombo(pCombo As ComboBox, Optional pDB As Variant)

    fillCombo_ pCombo, collectionAll(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOArticulos

    pCombo.Clear

    For Each objElemento In pCollection
        pCombo.AddItem objElemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = objElemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, Optional pDB As Variant)

    fillList_ pList, collectionAll(pDB)

End Sub

Public Sub fillListDescripcion(pList As ListBox, pCadena As String, Optional pDB As Variant)

    fillList_ pList, collectionByDescripcion(pCadena, pDB)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim objElemento As clsDAOArticulos

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function aumentarPrecios(pPorcentaje As Currency, pFecha As Date, pDB As clsDB) As Boolean
Dim objClonado As clsDAOArticulos

Dim curPorcentaje As Currency

    If pPorcentaje = 0 Then Exit Function
    
    curPorcentaje = 1 + (pPorcentaje / 100)

    Set objClonado = Me.clone
    
    With objClonado
        vPreciocomprasiniva = .preciocomprasiniva * curPorcentaje
        vPreciocomprasinivaanterior = .preciocomprasiniva
        vPreciolistaconiva = .preciolistaconiva * curPorcentaje
        vPrecioListaSinIva = .preciolistasiniva * curPorcentaje
        vPrecioventaconiva = .precioventaconiva * curPorcentaje
        vPrecioventasiniva = .precioventasiniva * curPorcentaje
    End With
    
    vFechaActualizacion = pFecha
    
    aumentarPrecios = update(pDB)
    
End Function

Public Sub exportar(pArchivo As String, pDB As clsDB)
Dim strSQL As String
Dim strLinea As String

Dim rstQuery As ADODB.Recordset

    Open pArchivo For Append As #1

    strSQL = "SELECT CONCAT('REPLACE INTO " & ddl.table
    strSQL = strSQL & "( " & ddl.codigo
    strSQL = strSQL & ", " & ddl.descripcion
    strSQL = strSQL & ", " & ddl.precioventaconiva
    strSQL = strSQL & ", " & ddl.precioventasiniva
    strSQL = strSQL & ", " & ddl.preciolistaconiva
    strSQL = strSQL & ", " & ddl.preciolistasiniva
    strSQL = strSQL & ", " & ddl.preciocomprasiniva
    strSQL = strSQL & ", " & ddl.preciocomprasinivaanterior
    strSQL = strSQL & ", " & ddl.iva105
    strSQL = strSQL & ", " & ddl.exento
    strSQL = strSQL & ", " & ddl.modelocamion
    strSQL = strSQL & ", " & ddl.fechaactualizacion
    strSQL = strSQL & ", " & ddl.origen
    strSQL = strSQL & ", " & ddl.descuento
    strSQL = strSQL & ", " & ddl.prvID
    strSQL = strSQL & ", " & ddl.ultimacompra
    strSQL = strSQL & ", " & ddl.marca
    strSQL = strSQL & ", " & ddl.catalogo
    strSQL = strSQL & ", " & ddl.preciolistasinivausd
    strSQL = strSQL & ", " & ddl.cotizacionID
    strSQL = strSQL & ") VALUES "
    strSQL = strSQL & "( ''', " & ddl.codigo & ", '''"
    strSQL = strSQL & ", ''', " & ddl.descripcion & ", '''"
    strSQL = strSQL & ", ''', " & ddl.precioventaconiva & ", '''"
    strSQL = strSQL & ", ''', " & ddl.precioventasiniva & ", '''"
    strSQL = strSQL & ", ''', " & ddl.preciolistaconiva & ", '''"
    strSQL = strSQL & ", ''', " & ddl.preciolistasiniva & ", '''"
    strSQL = strSQL & ", ''', " & ddl.preciocomprasiniva & ", '''"
    strSQL = strSQL & ", ''', " & ddl.preciocomprasinivaanterior & ", '''"
    strSQL = strSQL & ", ', " & ddl.iva105 & ", '"
    strSQL = strSQL & ", ', " & ddl.exento & ", '"
    strSQL = strSQL & ", ''', " & ddl.modelocamion & ", '''"
    strSQL = strSQL & ", ''', " & ddl.fechaactualizacion & ", '''"
    strSQL = strSQL & ", ''', " & ddl.origen & ", '''"
    strSQL = strSQL & ", ''', " & ddl.descuento & ", '''"
    strSQL = strSQL & ", ', " & ddl.prvID & ", '"
    strSQL = strSQL & ", ''', " & ddl.ultimacompra & ", '''"
    strSQL = strSQL & ", ''', " & ddl.marca & ", '''"
    strSQL = strSQL & ", ''', " & ddl.catalogo & ", '''"
    strSQL = strSQL & ", ''', " & ddl.preciolistasinivausd & ", '''"
    strSQL = strSQL & ", ', " & ddl.cotizacionID & ", '"
    strSQL = strSQL & ");') AS cadena FROM " & ddl.table
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    
    Do While Not rstQuery.EOF
        Print #1, rstQuery!cadena
    
        rstQuery.MoveNext
    Loop
    
    rstQuery.Close
    
    Close #1
    
End Sub


