VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vNegid As Integer
Private vIva1 As Currency
Private vIva2 As Currency
Private vFEProduccion As Integer
Private vTA As String

Private clon As clsDAOParametros

Public ddl As New clsDDLParametros

Private Sub Class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegid = 0
    vIva1 = 0
    vIva2 = 0
    vFEProduccion = 0
    vTA = ""

End Sub

Public Sub clean()

    Call Class_Initialize

End Sub

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get negid() As Variant

    negid = vNegid

End Property

Public Property Let negid(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get iva1() As Variant

    iva1 = vIva1

End Property

Public Property Let iva1(ByVal vNewValue As Variant)

On Error Resume Next

    vIva1 = vNewValue

End Property

Public Property Get iva2() As Variant

    iva2 = vIva2

End Property

Public Property Let iva2(ByVal vNewValue As Variant)

On Error Resume Next

    vIva2 = vNewValue

End Property

Public Property Get feproduccion() As Variant

    feproduccion = vFEProduccion

End Property

Public Property Let feproduccion(ByVal vNewValue As Variant)

On Error Resume Next

    vFEProduccion = vNewValue

End Property

Public Property Get ta() As Variant

    ta = vTA

End Property

Public Property Let ta(ByVal vNewValue As Variant)

On Error Resume Next

    vTA = vNewValue

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

Public Function clone() As clsDAOParametros
Dim objObjeto As New clsDAOParametros

    With objObjeto
        .clave = vClave
        .negid = vNegid
        .iva1 = vIva1
        .iva2 = vIva2
        .feproduccion = vFEProduccion
        .ta = vTA
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOParametros, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .negid = pQuery.Fields(.ddl.negid)
        .iva1 = pQuery.Fields(.ddl.iva1)
        .iva2 = pQuery.Fields(.ddl.iva2)
        .feproduccion = pQuery.Fields(.ddl.feproduccion)
        .ta = pQuery.Fields(.ddl.ta)
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
    strSQL = strSQL & " SET " & ddl.negid & " = " & vNegid
    strSQL = strSQL & ", " & ddl.iva1 & " = '" & Replace(vIva1, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.iva2 & " = '" & Replace(vIva2, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.feproduccion & " = " & vFEProduccion
    strSQL = strSQL & ", " & ddl.ta & " = '" & vTA & "'"
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

    If vNegid <> clon.negid Or IsNull(vNegid) <> IsNull(clon.negid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negid & " = " & vNegid
        blnChange = True
    End If
    If vIva1 <> clon.iva1 Or IsNull(vIva1) <> IsNull(clon.iva1) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva1 & " = '" & Replace(vIva1, ",", ".") & "'"
        blnChange = True
    End If
    If vIva2 <> clon.iva2 Or IsNull(vIva2) <> IsNull(clon.iva2) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva2 & " = '" & Replace(vIva2, ",", ".") & "'"
        blnChange = True
    End If
    If vFEProduccion <> clon.feproduccion Or IsNull(vFEProduccion) <> IsNull(clon.feproduccion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.feproduccion & " = " & vFEProduccion
        blnChange = True
    End If
    If vTA <> clon.ta Or IsNull(vTA) <> IsNull(clon.ta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ta & " = '" & vTA & "'"
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

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOParametros

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOParametros
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
Dim objElemento As clsDAOParametros

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
Dim objElemento As clsDAOParametros

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

