VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOTMDetartic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vHwnd As Long
Private vItem As Long
Private vCgoarticulo As String
Private vDescripcion As String
Private vCantidad As Double
Private vTotalconiva As Currency
Private vTotalsiniva As Currency
Private vPrecioventaconiva As Currency
Private vPrecioventasiniva As Currency
Private vPreciodescuentoconiva As Currency
Private vPreciodescuentosiniva As Currency
Private vPreciocomprasiniva As Currency
Private vItem2 As Integer
Private vIva105 As Currency
Private vExento As Currency
Private vNrofactura As Long
Private vFechafactura As Variant
Private vDescuento As Currency

Private clon As clsDAOTMDetartic

Public ddl As New clsDDLTMDetartic

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vHwnd = 0
    vItem = 0
    vCgoarticulo = ""
    vDescripcion = ""
    vCantidad = 0
    vTotalconiva = 0
    vTotalsiniva = 0
    vPrecioventaconiva = 0
    vPrecioventasiniva = 0
    vPreciodescuentoconiva = 0
    vPreciodescuentosiniva = 0
    vPreciocomprasiniva = 0
    vItem2 = 0
    vIva105 = 0
    vExento = 0
    vNrofactura = 0
    vFechafactura = Null
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

Public Property Get hwnd() As Variant

    hwnd = vHwnd

End Property

Public Property Let hwnd(ByVal vNewValue As Variant)

On Error Resume Next

    vHwnd = vNewValue

End Property

Public Property Get item() As Variant

    item = vItem

End Property

Public Property Let item(ByVal vNewValue As Variant)

On Error Resume Next

    vItem = vNewValue

End Property

Public Property Get cgoarticulo() As Variant

    cgoarticulo = vCgoarticulo

End Property

Public Property Let cgoarticulo(ByVal vNewValue As Variant)

On Error Resume Next

    vCgoarticulo = vNewValue

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

Public Property Get totalconiva() As Variant

    totalconiva = vTotalconiva

End Property

Public Property Let totalconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalconiva = vNewValue

End Property

Public Property Get totalsiniva() As Variant

    totalsiniva = vTotalsiniva

End Property

Public Property Let totalsiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vTotalsiniva = vNewValue

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

Public Property Get preciodescuentoconiva() As Variant

    preciodescuentoconiva = vPreciodescuentoconiva

End Property

Public Property Let preciodescuentoconiva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciodescuentoconiva = vNewValue

End Property

Public Property Get preciodescuentosiniva() As Variant

    preciodescuentosiniva = vPreciodescuentosiniva

End Property

Public Property Let preciodescuentosiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciodescuentosiniva = vNewValue

End Property

Public Property Get preciocomprasiniva() As Variant

    preciocomprasiniva = vPreciocomprasiniva

End Property

Public Property Let preciocomprasiniva(ByVal vNewValue As Variant)

On Error Resume Next

    vPreciocomprasiniva = vNewValue

End Property

Public Property Get item2() As Variant

    item2 = vItem2

End Property

Public Property Let item2(ByVal vNewValue As Variant)

On Error Resume Next

    vItem2 = vNewValue

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

Public Property Get nrofactura() As Variant

    nrofactura = vNrofactura

End Property

Public Property Let nrofactura(ByVal vNewValue As Variant)

On Error Resume Next

    vNrofactura = vNewValue

End Property

Public Property Get fechafactura() As Variant

    fechafactura = vFechafactura

End Property

Public Property Let fechafactura(ByVal vNewValue As Variant)

On Error Resume Next

    vFechafactura = vNewValue

End Property

Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vHwnd

End Property

Public Property Get comboKey() As Variant

    comboKey = vClave

End Property

Public Property Get listText() As Variant

    listText = vHwnd

End Property

Public Property Get listKey() As Variant

    listKey = vClave

End Property

Public Function clone() As clsDAOTMDetartic
Dim objObjeto As New clsDAOTMDetartic

    With objObjeto
        .clave = vClave
        .hwnd = vHwnd
        .item = vItem
        .cgoarticulo = vCgoarticulo
        .descripcion = vDescripcion
        .cantidad = vCantidad
        .totalconiva = vTotalconiva
        .totalsiniva = vTotalsiniva
        .precioventaconiva = vPrecioventaconiva
        .precioventasiniva = vPrecioventasiniva
        .preciodescuentoconiva = vPreciodescuentoconiva
        .preciodescuentosiniva = vPreciodescuentosiniva
        .preciocomprasiniva = vPreciocomprasiniva
        .item2 = vItem2
        .iva105 = vIva105
        .exento = vExento
        .nrofactura = vNrofactura
        .fechafactura = vFechafactura
        .descuento = vDescuento
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOTMDetartic, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .hwnd = pQuery.Fields(.ddl.hwnd)
        .item = pQuery.Fields(.ddl.item)
        .cgoarticulo = pQuery.Fields(.ddl.cgoarticulo)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .totalconiva = pQuery.Fields(.ddl.totalconiva)
        .totalsiniva = pQuery.Fields(.ddl.totalsiniva)
        .precioventaconiva = pQuery.Fields(.ddl.precioventaconiva)
        .precioventasiniva = pQuery.Fields(.ddl.precioventasiniva)
        .preciodescuentoconiva = pQuery.Fields(.ddl.preciodescuentoconiva)
        .preciodescuentosiniva = pQuery.Fields(.ddl.preciodescuentosiniva)
        .preciocomprasiniva = pQuery.Fields(.ddl.preciocomprasiniva)
        .item2 = pQuery.Fields(.ddl.item2)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .exento = pQuery.Fields(.ddl.exento)
        .nrofactura = pQuery.Fields(.ddl.nrofactura)
        .fechafactura = pQuery.Fields(.ddl.fechafactura)
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

    findByPrimaryKey_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findByPrimaryKey_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByItem(pDB As clsDB)
Dim strSQL As String

Dim lngHWnd As Long

Dim intItem As Integer

    lngHWnd = vHwnd
    intItem = vItem
    
    clean
    
    vHwnd = lngHWnd
    vItem = intItem
    
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & " AND " & ddl.item & " = " & vItem
    strSQL = strSQL & ";"
    
    findSQL strSQL, pDB
    
End Sub

Public Sub findLastID(Optional pDB As Variant)

    findLastID_ IIf(IsMissing(pDB), DB, pDB)

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

    findLast_ IIf(IsMissing(pDB), DB, pDB)

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

    delete = delete_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function delete_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete_ = pDB.execute(strSQL)

End Function

Public Function deleteAll(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ";"
    
    deleteAll = pDB.execute(strSQL)
    
End Function

Public Function save(Optional pDB As Variant) As Boolean

    save = save_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function save_(pDB As clsDB) As Boolean

    If exist_(pDB) Then
        save_ = update_(pDB)
    Else
        save_ = add_(pDB)
    End If

End Function

Public Function exist(Optional pDB As Variant) As Boolean

    exist = exist_(IIf(IsMissing(pDB), DB, pDB))

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

    add = add_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function add_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ", " & ddl.item & " = " & vItem
    strSQL = strSQL & ", " & ddl.cgoarticulo & " = '" & replaceSpecialSymbols(vCgoarticulo) & "'"
    strSQL = strSQL & ", " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    strSQL = strSQL & ", " & ddl.cantidad & " = '" & Replace(vCantidad, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.totalconiva & " = '" & Replace(vTotalconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.totalsiniva & " = '" & Replace(vTotalsiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.precioventaconiva & " = '" & Replace(vPrecioventaconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.precioventasiniva & " = '" & Replace(vPrecioventasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciodescuentoconiva & " = '" & Replace(vPreciodescuentoconiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciodescuentosiniva & " = '" & Replace(vPreciodescuentosiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.item2 & " = " & vItem2
    strSQL = strSQL & ", " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.nrofactura & " = " & vNrofactura
    strSQL = strSQL & ", " & ddl.fechafactura & " = " & fechaDB(vFechafactura)
    strSQL = strSQL & ", " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findLastID_ pDB

End Function

Public Function update(Optional pDB As Variant) As Boolean

    update = update_(IIf(IsMissing(pDB), DB, pDB))

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

    If vHwnd <> clon.hwnd Or IsNull(vHwnd) <> IsNull(clon.hwnd) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.hwnd & " = " & vHwnd
        blnChange = True
    End If
    If vItem <> clon.item Or IsNull(vItem) <> IsNull(clon.item) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.item & " = " & vItem
        blnChange = True
    End If
    If vCgoarticulo <> clon.cgoarticulo Or IsNull(vCgoarticulo) <> IsNull(clon.cgoarticulo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgoarticulo & " = '" & replaceSpecialSymbols(vCgoarticulo) & "'"
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
    If vTotalconiva <> clon.totalconiva Or IsNull(vTotalconiva) <> IsNull(clon.totalconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.totalconiva & " = '" & Replace(vTotalconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vTotalsiniva <> clon.totalsiniva Or IsNull(vTotalsiniva) <> IsNull(clon.totalsiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.totalsiniva & " = '" & Replace(vTotalsiniva, ",", ".") & "'"
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
    If vPreciodescuentoconiva <> clon.preciodescuentoconiva Or IsNull(vPreciodescuentoconiva) <> IsNull(clon.preciodescuentoconiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciodescuentoconiva & " = '" & Replace(vPreciodescuentoconiva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciodescuentosiniva <> clon.preciodescuentosiniva Or IsNull(vPreciodescuentosiniva) <> IsNull(clon.preciodescuentosiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciodescuentosiniva & " = '" & Replace(vPreciodescuentosiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vPreciocomprasiniva <> clon.preciocomprasiniva Or IsNull(vPreciocomprasiniva) <> IsNull(clon.preciocomprasiniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.preciocomprasiniva & " = '" & Replace(vPreciocomprasiniva, ",", ".") & "'"
        blnChange = True
    End If
    If vItem2 <> clon.item2 Or IsNull(vItem2) <> IsNull(clon.item2) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.item2 & " = " & vItem2
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
    If vNrofactura <> clon.nrofactura Or IsNull(vNrofactura) <> IsNull(clon.nrofactura) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nrofactura & " = " & vNrofactura
        blnChange = True
    End If
    If vFechafactura <> clon.fechafactura Or IsNull(vFechafactura) <> IsNull(clon.fechafactura) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechafactura & " = " & fechaDB(vFechafactura)
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

    newID_ pFind, IIf(IsMissing(pDB), DB, pDB)

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

Public Function collectionByhWnd(phWnd As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & phWnd
    strSQL = strSQL & " ORDER BY " & ddl.item & ", " & ddl.item
    strSQL = strSQL & ";"

    Set collectionByhWnd = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOTMDetartic

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOTMDetartic
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
Dim objElemento As clsDAOTMDetartic

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
Dim objElemento As clsDAOTMDetartic

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function eliminarItem(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.item & " = " & vItem
    strSQL = strSQL & " AND " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ";"
    
    eliminarItem = pDB.execute(strSQL)

End Function

Public Function renumerarItem(pDB As clsDB) As Boolean
Dim intItem As Integer

Dim objTM As New clsDAOTMDetartic

    intItem = 1
    
    For Each objTM In objTM.collectionByhWnd(vHwnd, pDB)
        With objTM
            .item = intItem
            intItem = intItem + 1
            renumerarItem = .update(pDB)
        End With
    Next

End Function


