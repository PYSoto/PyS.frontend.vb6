VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOArticulosImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vFecha As Date
Private vArticuloID As String
Private vDescripcion As String
Private vPrecioListaSinIva As Currency
Private vOrigen As String
Private vDescuento As String
Private vFechaActualizacion As Date
Private vCotizacionID As Long
Private vValorUSD As Currency
Private vClave As Long

Private clon As clsDAOArticulosImport

Public ddl As New clsDDLArticulosImport

Private Sub Class_Initialize()

    vFecha = Date
    vArticuloID = ""

    initializeOther

End Sub

Private Sub initializeOther()

    vDescripcion = ""
    vPrecioListaSinIva = 0
    vOrigen = ""
    vDescuento = ""
    vFechaActualizacion = Date
    vCotizacionID = 0
    vValorUSD = 0
    vClave = 0

End Sub

Public Sub clean()

    Call Class_Initialize

End Sub

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get articuloID() As Variant

    articuloID = vArticuloID

End Property

Public Property Let articuloID(ByVal vNewValue As Variant)

On Error Resume Next

    vArticuloID = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get preciolistasiniva() As Variant

    preciolistasiniva = vPrecioListaSinIva

End Property

Public Property Let preciolistasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioListaSinIva = vNewValue

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

Public Property Get fechaactualizacion() As Variant

    fechaactualizacion = vFechaActualizacion

End Property

Public Property Let fechaactualizacion(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaActualizacion = vNewValue

End Property

Public Property Get cotizacionID() As Variant

    cotizacionID = vCotizacionID

End Property

Public Property Let cotizacionID(ByVal vNewValue As Variant)

On Error Resume Next

    vCotizacionID = vNewValue

End Property

Public Property Get valorUSD() As Variant

    valorUSD = vValorUSD

End Property

Public Property Let valorUSD(ByVal vNewValue As Variant)

On Error Resume Next

    vValorUSD = vNewValue

End Property

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vDescripcion

End Property

Public Property Get comboKey() As Variant

    comboKey = vArticuloID

End Property

Public Property Get listText() As Variant

    listText = vDescripcion

End Property

Public Property Get listKey() As Variant

    listKey = vArticuloID

End Property

Public Function clone() As clsDAOArticulosImport
Dim objObjeto As New clsDAOArticulosImport

    With objObjeto
        .fecha = vFecha
        .articuloID = vArticuloID
        .descripcion = vDescripcion
        .preciolistasiniva = vPrecioListaSinIva
        .origen = vOrigen
        .descuento = vDescuento
        .fechaactualizacion = vFechaActualizacion
        .cotizacionID = vCotizacionID
        .valorUSD = vValorUSD
        .clave = vClave
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOArticulosImport, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .fecha = pQuery.Fields(.ddl.fecha)
        .articuloID = pQuery.Fields(.ddl.articuloID)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .preciolistasiniva = pQuery.Fields(.ddl.preciolistasiniva)
        .origen = pQuery.Fields(.ddl.origen)
        .descuento = pQuery.Fields(.ddl.descuento)
        .fechaactualizacion = pQuery.Fields(.ddl.fechaactualizacion)
        .cotizacionID = pQuery.Fields(.ddl.cotizacionID)
        .valorUSD = pQuery.Fields(.ddl.valorUSD)
        .clave = pQuery.Fields(.ddl.clave)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " AND " & ddl.articuloID & " = '" & replaceSpecialSymbols(vArticuloID) & "'"

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

Public Sub findByPrimaryKey(Optional pDB As Variant)

    findByPrimaryKey_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findByPrimaryKey_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " ORDER BY " & ddl.articuloID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByClave(Optional pDB As Variant)

    findByClave_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findByClave_(pDB As clsDB)
Dim strSQL As String

Dim lngClave As Long

    lngClave = vClave

    clean

    vClave = lngClave

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.clave & " = " & vClave
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
    strSQL = strSQL & " SET " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    strSQL = strSQL & ", " & ddl.preciolistasiniva & " = '" & Replace(vPrecioListaSinIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.origen & " = '" & replaceSpecialSymbols(vOrigen) & "'"
    strSQL = strSQL & ", " & ddl.descuento & " = '" & replaceSpecialSymbols(vDescuento) & "'"
    strSQL = strSQL & ", " & ddl.fechaactualizacion & " = " & fechaDB(vFechaActualizacion)
    strSQL = strSQL & ", " & ddl.cotizacionID & " = " & vCotizacionID
    strSQL = strSQL & ", " & ddl.valorUSD & " = '" & Replace(vValorUSD, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & ", " & ddl.articuloID & " = '" & replaceSpecialSymbols(vArticuloID) & "'"
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findByPrimaryKey_ pDB

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
    If vPrecioListaSinIva <> clon.preciolistasiniva Or IsNull(vPrecioListaSinIva) <> IsNull(clon.preciolistasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciolistasiniva & " = '" & Replace(vPrecioListaSinIva, ",", ".") & "'"
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
    If vFechaActualizacion <> clon.fechaactualizacion Or IsNull(vFechaActualizacion) <> IsNull(clon.fechaactualizacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaactualizacion & " = " & fechaDB(vFechaActualizacion)
        blnChange = True
    End If
    If vCotizacionID <> clon.cotizacionID Or IsNull(vCotizacionID) <> IsNull(clon.cotizacionID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cotizacionID & " = " & vCotizacionID
        blnChange = True
    End If
    If vValorUSD <> clon.valorUSD Or IsNull(vValorUSD) <> IsNull(clon.valorUSD) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.valorUSD & " = '" & Replace(vValorUSD, ",", ".") & "'"
        blnChange = True
    End If

    strSQL = "UPDATE " & ddl.table & " SET" & strSET
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    If blnChange Then update_ = pDB.execute(strSQL)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, Optional pDB As Variant)

    newID_ pFind, IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub newID_(pFind As Boolean, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    vArticuloID = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " ORDER BY " & ddl.articuloID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vArticuloID = 1 + rstQuery.Fields(ddl.articuloID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionByFecha(pFecha As Date, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.fecha & " = " & fechaDB(pFecha)
    strSQL = strSQL & ";"

    Set collectionByFecha = collectionSQL(strSQL, pDB)

End Function

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOArticulosImport

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOArticulosImport
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Sub fillCombo(pCombo As ComboBox, Optional pDB As Variant)

    fillCombo_ pCombo, collectionAll(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOArticulosImport

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

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim objElemento As clsDAOArticulosImport

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

