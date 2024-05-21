VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDAOArticulosMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vDsn As String
Private vUid As String
Private vPwd As String
Private vIp As String
Private vDatabase As String
Private vId As Integer

Private clon As clsDAOArticulosMirror

Public ddl As New clsDDLArticulosmirror

Private Sub class_Initialize()

    vId = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDsn = ""
    vUid = ""
    vPwd = ""
    vIp = ""
    vDatabase = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get dsn() As Variant

    dsn = vDsn

End Property

Public Property Let dsn(ByVal vNewValue As Variant)

On Error Resume Next

    vDsn = vNewValue

End Property

Public Property Get uid() As Variant

    uid = vUid

End Property

Public Property Let uid(ByVal vNewValue As Variant)

On Error Resume Next

    vUid = vNewValue

End Property

Public Property Get pwd() As Variant

    pwd = vPwd

End Property

Public Property Let pwd(ByVal vNewValue As Variant)

On Error Resume Next

    vPwd = vNewValue

End Property

Public Property Get ip() As Variant

    ip = vIp

End Property

Public Property Let ip(ByVal vNewValue As Variant)

On Error Resume Next

    vIp = vNewValue

End Property

Public Property Get database() As Variant

    database = vDatabase

End Property

Public Property Let database(ByVal vNewValue As Variant)

On Error Resume Next

    vDatabase = vNewValue

End Property

Public Property Get id() As Variant

    id = vId

End Property

Public Property Let id(ByVal vNewValue As Variant)

On Error Resume Next

    vId = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vDsn

End Property

Public Property Get comboKey() As Variant

    comboKey = vId

End Property

Public Property Get listText() As Variant

    listText = vDsn

End Property

Public Property Get listKey() As Variant

    listKey = vId

End Property

Public Function clone() As clsDAOArticulosMirror
Dim objObjeto As New clsDAOArticulosMirror

    With objObjeto
        .dsn = vDsn
        .uid = vUid
        .pwd = vPwd
        .ip = vIp
        .database = vDatabase
        .id = vId
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOArticulosMirror, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .dsn = pQuery.fields(.ddl.dsn)
        .uid = pQuery.fields(.ddl.uid)
        .pwd = pQuery.fields(.ddl.pwd)
        .ip = pQuery.fields(.ddl.ip)
        .database = pQuery.fields(.ddl.database)
        .id = pQuery.fields(.ddl.id)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.id & " = " & vId

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

Public Sub findLastID(Optional pDB As Variant)

    findLastID_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findLastID_(pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    clean

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.id
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vId = rstQuery.Fields(ddl.id)

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
    strSQL = strSQL & " ORDER BY " & ddl.id & " DESC"
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
    strSQL = strSQL & " SET " & ddl.dsn & " = '" & ReplaceSpecialSymbols(vDsn) & "'"
    strSQL = strSQL & ", " & ddl.uid & " = '" & ReplaceSpecialSymbols(vUid) & "'"
    strSQL = strSQL & ", " & ddl.pwd & " = '" & ReplaceSpecialSymbols(vPwd) & "'"
    strSQL = strSQL & ", " & ddl.ip & " = '" & ReplaceSpecialSymbols(vIp) & "'"
    strSQL = strSQL & ", " & ddl.database & " = '" & ReplaceSpecialSymbols(vDatabase) & "'"
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
        Set clon = Me.Clone
        clon.findByPrimaryKey pDB
    End If

    If vDsn <> clon.Dsn Or IsNull(vDsn) <> IsNull(clon.Dsn) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.dsn & " = '" & ReplaceSpecialSymbols(vDsn) & "'"
        blnChange = True
    End If
    If vUid <> clon.Uid Or IsNull(vUid) <> IsNull(clon.Uid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.uid & " = '" & ReplaceSpecialSymbols(vUid) & "'"
        blnChange = True
    End If
    If vPwd <> clon.Pwd Or IsNull(vPwd) <> IsNull(clon.Pwd) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.pwd & " = '" & ReplaceSpecialSymbols(vPwd) & "'"
        blnChange = True
    End If
    If vIp <> clon.Ip Or IsNull(vIp) <> IsNull(clon.Ip) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ip & " = '" & ReplaceSpecialSymbols(vIp) & "'"
        blnChange = True
    End If
    If vDatabase <> clon.Database Or IsNull(vDatabase) <> IsNull(clon.Database) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.database & " = '" & ReplaceSpecialSymbols(vDatabase) & "'"
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

    vId = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.id & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vId = 1 + rstQuery.fields(ddl.id)

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

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOArticulosMirror

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOArticulosMirror
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
Dim objElemento As clsDAOArticulosMirror

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
Dim objElemento As clsDAOArticulosMirror

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

