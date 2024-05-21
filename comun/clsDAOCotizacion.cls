VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCotizacionID As Long
Private vFecha As Date
Private vUSDCompra As Currency
Private vUSDVenta As Currency

Private clon As clsDAOCotizacion

Public ddl As New clsDDLCotizacion

Private Sub class_Initialize()

    vCotizacionID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vFecha = Date
    vUSDCompra = 0
    vUSDVenta = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get cotizacionID() As Variant

    cotizacionID = vCotizacionID

End Property

Public Property Let cotizacionID(ByVal vNewValue As Variant)

On Error Resume Next

    vCotizacionID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get usdcompra() As Variant

    usdcompra = vUSDCompra

End Property

Public Property Let usdcompra(ByVal vNewValue As Variant)

On Error Resume Next

    vUSDCompra = vNewValue

End Property

Public Property Get usdventa() As Variant

    usdventa = vUSDVenta

End Property

Public Property Let usdventa(ByVal vNewValue As Variant)

On Error Resume Next

    vUSDVenta = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vFecha

End Property

Public Property Get comboKey() As Variant

    comboKey = vCotizacionID

End Property

Public Property Get listText() As Variant

    listText = vFecha

End Property

Public Property Get listKey() As Variant

    listKey = vCotizacionID

End Property

Public Function clone() As clsDAOCotizacion
Dim objObjeto As New clsDAOCotizacion

    With objObjeto
        .cotizacionID = vCotizacionID
        .fecha = vFecha
        .usdcompra = vUSDCompra
        .usdventa = vUSDVenta
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOCotizacion, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .cotizacionID = pQuery.Fields(.ddl.cotizacionID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .usdcompra = pQuery.Fields(.ddl.usdcompra)
        .usdventa = pQuery.Fields(.ddl.usdventa)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.cotizacionID & " = " & vCotizacionID

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

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.cotizacionID
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCotizacionID = rstQuery.Fields(ddl.cotizacionID)

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
    strSQL = strSQL & " ORDER BY " & ddl.cotizacionID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByFecha(Optional pDB As Variant)

    findByFecha_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findByFecha_(pDB As clsDB)
Dim strSQL As String

Dim datFecha As Date

    datFecha = vFecha

    clean

    vFecha = datFecha

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.fecha & " = " & fechaDB(vFecha)
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
    strSQL = strSQL & " SET " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & ", " & ddl.usdcompra & " = '" & Replace(vUSDCompra, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.usdventa & " = '" & Replace(vUSDVenta, ",", ".") & "'"
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

    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vUSDCompra <> clon.usdcompra Or IsNull(vUSDCompra) <> IsNull(clon.usdcompra) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.usdcompra & " = '" & Replace(vUSDCompra, ",", ".") & "'"
        blnChange = True
    End If
    If vUSDVenta <> clon.usdventa Or IsNull(vUSDVenta) <> IsNull(clon.usdventa) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.usdventa & " = '" & Replace(vUSDVenta, ",", ".") & "'"
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

    vCotizacionID = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.cotizacionID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCotizacionID = 1 + rstQuery.Fields(ddl.cotizacionID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

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

Dim objElemento As clsDAOCotizacion

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOCotizacion
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
Dim objElemento As clsDAOCotizacion

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
Dim objElemento As clsDAOCotizacion

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

