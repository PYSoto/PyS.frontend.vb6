VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAODetartic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vClavemovclie As Long
Private vCgocomprob As Integer
Private vItem As Integer
Private vNegid As Integer
Private vCgoartic As String
Private vDescripcion As String
Private vCantidad As Double
Private vPrecioventasiniva As Currency
Private vPrecioventaconiva As Currency
Private vPreciodescuentosiniva As Currency
Private vPreciodescuentoconiva As Currency
Private vIva105 As Currency
Private vExento As Currency
Private vFecha As Variant
Private vFechafac As Variant
Private vPreciocomprasiniva As Currency
Private vDescuento As Currency

Private clon As clsDAODetartic

Public ddl As New clsDDLDetartic

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vClavemovclie = 0
    vCgocomprob = 0
    vItem = 0
    vNegid = gEmpresa.negid
    vCgoartic = ""
    vDescripcion = ""
    vCantidad = 0
    vPrecioventasiniva = 0
    vPrecioventaconiva = 0
    vPreciodescuentosiniva = 0
    vPreciodescuentoconiva = 0
    vIva105 = 0
    vExento = 0
    vFecha = Null
    vFechafac = Null
    vPreciocomprasiniva = 0
    vDescuento = 0

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

Public Property Get clavemovclie() As Variant

    clavemovclie = vClavemovclie

End Property

Public Property Let clavemovclie(ByVal vNewValue As Variant)

On Error Resume Next

    vClavemovclie = vNewValue

End Property

Public Property Get cgocomprob() As Variant

    cgocomprob = vCgocomprob

End Property

Public Property Let cgocomprob(ByVal vNewValue As Variant)

On Error Resume Next

    vCgocomprob = vNewValue

End Property

Public Property Get item() As Variant

    item = vItem

End Property

Public Property Let item(ByVal vNewValue As Variant)

On Error Resume Next

    vItem = vNewValue

End Property

Public Property Get negid() As Variant

    negid = vNegid

End Property

Public Property Let negid(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get cgoartic() As Variant

    cgoartic = vCgoartic

End Property

Public Property Let cgoartic(ByVal vNewValue As Variant)

On Error Resume Next

    vCgoartic = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get cantidad() As Variant

    cantidad = vCantidad

End Property

Public Property Let cantidad(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidad = vNewValue

End Property

Public Property Get precioventasiniva() As Variant

    precioventasiniva = vPrecioventasiniva

End Property

Public Property Let precioventasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioventasiniva = vNewValue

End Property

Public Property Get precioventaconiva() As Variant

    precioventaconiva = vPrecioventaconiva

End Property

Public Property Let precioventaconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioventaconiva = vNewValue

End Property

Public Property Get preciodescuentosiniva() As Variant

    preciodescuentosiniva = vPreciodescuentosiniva

End Property

Public Property Let preciodescuentosiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciodescuentosiniva = vNewValue

End Property

Public Property Get preciodescuentoconiva() As Variant

    preciodescuentoconiva = vPreciodescuentoconiva

End Property

Public Property Let preciodescuentoconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciodescuentoconiva = vNewValue

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

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get fechafac() As Variant

    fechafac = vFechafac

End Property

Public Property Let fechafac(ByVal vNewValue As Variant)

On Error Resume Next

    vFechafac = vNewValue

End Property

Public Property Get preciocomprasiniva() As Variant

    preciocomprasiniva = vPreciocomprasiniva

End Property

Public Property Let preciocomprasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciocomprasiniva = vNewValue

End Property

Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vClavemovclie

End Property

Public Property Get comboKey() As Variant

    comboKey = vClave

End Property

Public Property Get listText() As Variant

    listText = vClavemovclie

End Property

Public Property Get listKey() As Variant

    listKey = vClave

End Property

Public Function clone() As clsDAODetartic
Dim objObjeto As New clsDAODetartic

    With objObjeto
        .clave = vClave
        .clavemovclie = vClavemovclie
        .cgocomprob = vCgocomprob
        .item = vItem
        .negid = vNegid
        .cgoartic = vCgoartic
        .descripcion = vDescripcion
        .cantidad = vCantidad
        .precioventasiniva = vPrecioventasiniva
        .precioventaconiva = vPrecioventaconiva
        .preciodescuentosiniva = vPreciodescuentosiniva
        .preciodescuentoconiva = vPreciodescuentoconiva
        .iva105 = vIva105
        .exento = vExento
        .fecha = vFecha
        .fechafac = vFechafac
        .preciocomprasiniva = vPreciocomprasiniva
        .descuento = vDescuento
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAODetartic, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .clavemovclie = pQuery.Fields(.ddl.clavemovclie)
        .cgocomprob = pQuery.Fields(.ddl.cgocomprob)
        .item = pQuery.Fields(.ddl.item)
        .negid = pQuery.Fields(.ddl.negid)
        .cgoartic = pQuery.Fields(.ddl.cgoartic)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .precioventasiniva = pQuery.Fields(.ddl.precioventasiniva)
        .precioventaconiva = pQuery.Fields(.ddl.precioventaconiva)
        .preciodescuentosiniva = pQuery.Fields(.ddl.preciodescuentosiniva)
        .preciodescuentoconiva = pQuery.Fields(.ddl.preciodescuentoconiva)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .exento = pQuery.Fields(.ddl.exento)
        .fecha = pQuery.Fields(.ddl.fecha)
        .fechafac = pQuery.Fields(.ddl.fechafac)
        .preciocomprasiniva = pQuery.Fields(.ddl.preciocomprasiniva)
        .descuento = pQuery.Fields(.ddl.descuento)
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

    findByPrimaryKey_ pDB

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
    strSQL = strSQL & " SET " & ddl.clavemovclie & " = " & vClavemovclie
    strSQL = strSQL & ", " & ddl.cgocomprob & " = " & vCgocomprob
    strSQL = strSQL & ", " & ddl.item & " = " & vItem
    strSQL = strSQL & ", " & ddl.negid & " = " & vNegid
    strSQL = strSQL & ", " & ddl.cgoartic & " = '" & replaceSpecialSymbols(vCgoartic) & "'"
    strSQL = strSQL & ", " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    strSQL = strSQL & ", " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.precioventasiniva & " = '" & Replace(vPrecioventasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.precioventaconiva & " = '" & Replace(vPrecioventaconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciodescuentosiniva & " = '" & Replace(vPreciodescuentosiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciodescuentoconiva & " = '" & Replace(vPreciodescuentoconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & ", " & ddl.fechafac & " = " & fechaDB(vFechafac)
    strSQL = strSQL & ", " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findLastID_ pDB

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

    If vClavemovclie <> clon.clavemovclie Or IsNull(vClavemovclie) <> IsNull(clon.clavemovclie) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clavemovclie & " = " & vClavemovclie
        blnChange = True
    End If
    If vCgocomprob <> clon.cgocomprob Or IsNull(vCgocomprob) <> IsNull(clon.cgocomprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgocomprob & " = " & vCgocomprob
        blnChange = True
    End If
    If vItem <> clon.item Or IsNull(vItem) <> IsNull(clon.item) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.item & " = " & vItem
        blnChange = True
    End If
    If vNegid <> clon.negid Or IsNull(vNegid) <> IsNull(clon.negid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negid & " = " & vNegid
        blnChange = True
    End If
    If vCgoartic <> clon.cgoartic Or IsNull(vCgoartic) <> IsNull(clon.cgoartic) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgoartic & " = '" & replaceSpecialSymbols(vCgoartic) & "'"
        blnChange = True
    End If
    If vDescripcion <> clon.descripcion Or IsNull(vDescripcion) <> IsNull(clon.descripcion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
        blnChange = True
    End If
    If vCantidad <> clon.cantidad Or IsNull(vCantidad) <> IsNull(clon.cantidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
        blnChange = True
    End If
    If vPrecioventasiniva <> clon.precioventasiniva Or IsNull(vPrecioventasiniva) <> IsNull(clon.precioventasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioventasiniva & " = '" & Replace(vPrecioventasiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vPrecioventaconiva <> clon.precioventaconiva Or IsNull(vPrecioventaconiva) <> IsNull(clon.precioventaconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioventaconiva & " = '" & Replace(vPrecioventaconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciodescuentosiniva <> clon.preciodescuentosiniva Or IsNull(vPreciodescuentosiniva) <> IsNull(clon.preciodescuentosiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciodescuentosiniva & " = '" & Replace(vPreciodescuentosiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciodescuentoconiva <> clon.preciodescuentoconiva Or IsNull(vPreciodescuentoconiva) <> IsNull(clon.preciodescuentoconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciodescuentoconiva & " = '" & Replace(vPreciodescuentoconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vIva105 <> clon.iva105 Or IsNull(vIva105) <> IsNull(clon.iva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
        blnChange = True
    End If
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
        blnChange = True
    End If
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vFechafac <> clon.fechafac Or IsNull(vFechafac) <> IsNull(clon.fechafac) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechafac & " = " & fechaDB(vFechafac)
        blnChange = True
    End If
    If vPreciocomprasiniva <> clon.preciocomprasiniva Or IsNull(vPreciocomprasiniva) <> IsNull(clon.preciocomprasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vDescuento <> clon.descuento Or IsNull(vDescuento) <> IsNull(clon.descuento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
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

    vClave = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.clave & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vClave = 1 + rstQuery.Fields(ddl.clave)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByClaveMovClie(pClaveMovClie As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.clavemovclie & " = " & pClaveMovClie
    strSQL = strSQL & ";"

    Set collectionByClaveMovClie = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAODetartic

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAODetartic
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
Dim objElemento As clsDAODetartic

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
Dim objElemento As clsDAODetartic

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function cantidadMovs(pArtID As String, pDB As clsDB) As Long
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    cantidadMovs = 0
    
    strSQL = "SELECT COUNT(*) AS cantidad FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoartic & " = '" & pArtID & "'"
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    
    If Not rstQuery.EOF Then If Not IsNull(rstQuery!cantidad) Then cantidadMovs = rstQuery!cantidad
    
    rstQuery.Close
    
End Function

