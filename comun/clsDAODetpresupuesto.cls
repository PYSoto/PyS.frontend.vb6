VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAODetpresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPreid As Long
Private vOrden As Integer
Private vArtid As String
Private vCantidad As Integer
Private vUnitsiva As Currency
Private vUnitciva As Currency
Private vClave As Long

Private clon As clsDAODetpresupuesto

Public ddl As New clsDDLDetpresupuesto

Private Sub class_Initialize()

    vPreid = 0
    vOrden = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vArtid = ""
    vCantidad = 0
    vUnitsiva = 0
    vUnitciva = 0
    vClave = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get preID() As Variant

    preID = vPreid

End Property

Public Property Let preID(ByVal vNewValue As Variant)

On Error Resume Next

    vPreid = vNewValue

End Property

Public Property Get orden() As Variant

    orden = vOrden

End Property

Public Property Let orden(ByVal vNewValue As Variant)

On Error Resume Next

    vOrden = vNewValue

End Property

Public Property Get artid() As Variant

    artid = vArtid

End Property

Public Property Let artid(ByVal vNewValue As Variant)

On Error Resume Next

    vArtid = vNewValue

End Property

Public Property Get cantidad() As Variant

    cantidad = vCantidad

End Property

Public Property Let cantidad(ByVal vNewValue As Variant)

On Error Resume Next

    vCantidad = vNewValue

End Property

Public Property Get unitsiva() As Variant

    unitsiva = vUnitsiva

End Property

Public Property Let unitsiva(ByVal vNewValue As Variant)

On Error Resume Next

    vUnitsiva = vNewValue

End Property

Public Property Get unitciva() As Variant

    unitciva = vUnitciva

End Property

Public Property Let unitciva(ByVal vNewValue As Variant)

On Error Resume Next

    vUnitciva = vNewValue

End Property

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vArtid

End Property

Public Property Get comboKey() As Variant

    comboKey = vOrden

End Property

Public Property Get listText() As Variant

    listText = vArtid

End Property

Public Property Get listKey() As Variant

    listKey = vOrden

End Property

Public Function clone() As clsDAODetpresupuesto
Dim objObjeto As New clsDAODetpresupuesto

    With objObjeto
        .preID = vPreid
        .orden = vOrden
        .artid = vArtid
        .cantidad = vCantidad
        .unitsiva = vUnitsiva
        .unitciva = vUnitciva
        .clave = vClave
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAODetpresupuesto, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .preID = pQuery.Fields(.ddl.preID)
        .orden = pQuery.Fields(.ddl.orden)
        .artid = pQuery.Fields(.ddl.artid)
        .cantidad = pQuery.Fields(.ddl.cantidad)
        .unitsiva = pQuery.Fields(.ddl.unitsiva)
        .unitciva = pQuery.Fields(.ddl.unitciva)
        .clave = pQuery.Fields(.ddl.clave)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & vPreid
    strSQL = strSQL & " AND " & ddl.orden & " = " & vOrden

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

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & vPreid
    strSQL = strSQL & " ORDER BY " & ddl.orden & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByClave(Optional pDB As Variant)

    findByClave_ IIf(IsMissing(pDB), DB, pDB)

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

    delete = delete_(IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function delete_(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete_ = pDB.execute(strSQL)

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
    strSQL = strSQL & " SET " & ddl.artid & " = '" & replaceSpecialSymbols(vArtid) & "'"
    strSQL = strSQL & ", " & ddl.cantidad & " = " & vCantidad
    strSQL = strSQL & ", " & ddl.unitsiva & " = '" & Replace(vUnitsiva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.unitciva & " = '" & Replace(vUnitciva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.preID & " = " & vPreid
    strSQL = strSQL & ", " & ddl.orden & " = " & vOrden
    strSQL = strSQL & ";"

    add_ = pDB.execute(strSQL)

    findByPrimaryKey_ pDB

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

    If vArtid <> clon.artid Or IsNull(vArtid) <> IsNull(clon.artid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.artid & " = '" & replaceSpecialSymbols(vArtid) & "'"
        blnChange = True
    End If
    If vCantidad <> clon.cantidad Or IsNull(vCantidad) <> IsNull(clon.cantidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cantidad & " = " & vCantidad
        blnChange = True
    End If
    If vUnitsiva <> clon.unitsiva Or IsNull(vUnitsiva) <> IsNull(clon.unitsiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.unitsiva & " = '" & Replace(vUnitsiva, ",", ".") & "'"
        blnChange = True
    End If
    If vUnitciva <> clon.unitciva Or IsNull(vUnitciva) <> IsNull(clon.unitciva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.unitciva & " = '" & Replace(vUnitciva, ",", ".") & "'"
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

    vOrden = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & vPreid
    strSQL = strSQL & " ORDER BY " & ddl.orden & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vOrden = 1 + rstQuery.Fields(ddl.orden)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionByPreID(pPreID As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & pPreID
    strSQL = strSQL & " ORDER BY " & ddl.orden
    strSQL = strSQL & ";"

    Set collectionByPreID = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAODetpresupuesto

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAODetpresupuesto
        fillObject objElemento, rstQuery

        colRegistros.add objElemento, "k." & objElemento.orden

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collection_ = colRegistros

End Function

Public Function deleteRest(pPreID As Long, pOrden As Integer, pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & pPreID
    strSQL = strSQL & " AND " & ddl.orden & " >= " & pOrden
    strSQL = strSQL & ";"
    
    deleteRest = pDB.execute(strSQL)
    
End Function
