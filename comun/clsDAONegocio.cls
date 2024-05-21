VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAONegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNegid As Integer
Private vNombre As String
Private vDomicilio As String
Private vDsn As String
Private vIp As Variant
Private vDB As Variant
Private vUser As Variant

Private clon As clsDAONegocio

Public ddl As New clsDDLNegocio

Private Sub Class_Initialize()

    vNegid = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNombre = ""
    vDomicilio = ""
    vIp = Null
    vDB = Null
    vUser = Null

End Sub

Public Sub clean()

    Call Class_Initialize

End Sub

Public Property Get negID() As Variant

    negID = vNegid

End Property

Public Property Let negID(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get nombre() As Variant

    nombre = vNombre

End Property

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get domicilio() As Variant

    domicilio = vDomicilio

End Property

Public Property Let domicilio(ByVal vNewValue As Variant)

On Error Resume Next

    vDomicilio = vNewValue

End Property

Public Property Get dsn() As Variant

    dsn = vDsn

End Property

Public Property Let dsn(ByVal vNewValue As Variant)

On Error Resume Next

    vDsn = vNewValue

End Property

Public Property Get ip() As Variant

    ip = vIp

End Property

Public Property Let ip(ByVal vNewValue As Variant)

On Error Resume Next

    vIp = vNewValue

End Property

Public Property Get database() As Variant

    database = vDB

End Property

Public Property Let database(ByVal vNewValue As Variant)

On Error Resume Next

    vDB = vNewValue

End Property

Public Property Get user() As Variant

    user = vUser

End Property

Public Property Let user(ByVal vNewValue As Variant)

On Error Resume Next

    vUser = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vNombre

End Property

Public Property Get comboKey() As Variant

    comboKey = vNegid

End Property

Public Property Get listText() As Variant

    listText = vNombre

End Property

Public Property Get listKey() As Variant

    listKey = vNegid

End Property

Public Function clone() As clsDAONegocio
Dim objObjeto As New clsDAONegocio

    With objObjeto
        .negID = vNegid
        .nombre = vNombre
        .domicilio = vDomicilio
        .dsn = vDsn
        .ip = vIp
        .database = vDB
        .user = vUser
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAONegocio, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .negID = pQuery.Fields(.ddl.negID)
        .nombre = pQuery.Fields(.ddl.nombre)
        .domicilio = pQuery.Fields(.ddl.domicilio)
        .dsn = pQuery.Fields(.ddl.dsn)
        .ip = pQuery.Fields(.ddl.ip)
        .database = pQuery.Fields(.ddl.db)
        .user = pQuery.Fields(.ddl.user)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.negID & " = " & vNegid

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

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.negID
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vNegid = rstQuery.Fields(ddl.negID)

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
    strSQL = strSQL & " ORDER BY " & ddl.negID & " DESC"
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
    strSQL = strSQL & " SET " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    strSQL = strSQL & ", " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
    strSQL = strSQL & ", " & ddl.dsn & " = '" & replaceSpecialSymbols(vDsn) & "'"
    strSQL = strSQL & ", " & ddl.ip & " = " & validateNullField(vIp)
    strSQL = strSQL & ", " & ddl.db & " = " & validateNullField(vDB)
    strSQL = strSQL & ", " & ddl.user & " = " & validateNullField(vUser)
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

    If vNombre <> clon.nombre Or IsNull(vNombre) <> IsNull(clon.nombre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
        blnChange = True
    End If
    If vDomicilio <> clon.domicilio Or IsNull(vDomicilio) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
        blnChange = True
    End If
    If vDsn <> clon.domicilio Or IsNull(vDsn) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDsn) & "'"
        blnChange = True
    End If
    If vIp <> clon.ip Or IsNull(vIp) <> IsNull(clon.ip) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ip & " = " & validateNullField(vIp)
        blnChange = True
    End If
    If vDB <> clon.database Or IsNull(vDB) <> IsNull(clon.database) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.db & " = " & validateNullField(vDB)
        blnChange = True
    End If
    If vUser <> clon.user Or IsNull(vUser) <> IsNull(clon.user) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.user & " = " & validateNullField(vUser)
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

    vNegid = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.negID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vNegid = 1 + rstQuery.Fields(ddl.negID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.negID
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionExcluidoLocal(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.negID & " <> " & gEmpresa.negID
    strSQL = strSQL & ";"

    Set collectionExcluidoLocal = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAONegocio

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAONegocio
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

Public Sub fillComboExcluidoLocal(pCombo As ComboBox, Optional pDB As Variant)

    fillCombo_ pCombo, collectionExcluidoLocal(pDB)
    
End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAONegocio

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
Dim objElemento As clsDAONegocio

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

