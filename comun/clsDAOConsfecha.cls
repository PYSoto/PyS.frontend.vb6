VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0   'vbNone
  MTSTransactionMode  = 0   'NotAnMTSObject
END
Attribute VB_Name = "clsDAOConsfecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNegid As Integer
Private vFecha As Date
Private vCuenta As Long
Private vDeudor As Currency
Private vAcreedor As Currency

Private clon As clsDAOConsfecha

Public ddl As New clsDDLConsfecha

Private Sub class_Initialize()

    vNegid = 0
    vFecha = Date
    vCuenta = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDeudor = 0
    vAcreedor = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get negid() As Variant

    negid = vNegid

End Property

Public Property Let negid(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get cuenta() As Variant

    cuenta = vCuenta

End Property

Public Property Let cuenta(ByVal vNewValue As Variant)

On Error Resume Next

    vCuenta = vNewValue

End Property

Public Property Get deudor() As Variant

    deudor = vDeudor

End Property

Public Property Let deudor(ByVal vNewValue As Variant)

On Error Resume Next

    vDeudor = vNewValue

End Property

Public Property Get acreedor() As Variant

    acreedor = vAcreedor

End Property

Public Property Let acreedor(ByVal vNewValue As Variant)

On Error Resume Next

    vAcreedor = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vDeudor

End Property

Public Property Get comboKey() As Variant

    comboKey = vCuenta

End Property

Public Property Get listText() As Variant

    listText = vDeudor

End Property

Public Property Get listKey() As Variant

    listKey = vCuenta

End Property

Public Function clone() As clsDAOConsfecha
Dim objObjeto As New clsDAOConsfecha

    With objObjeto
        .negid = vNegid
        .fecha = vFecha
        .cuenta = vCuenta
        .deudor = vDeudor
        .acreedor = vAcreedor
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOConsfecha, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .negid = pQuery.fields(.ddl.negid)
        .fecha = pQuery.fields(.ddl.fecha)
        .cuenta = pQuery.fields(.ddl.cuenta)
        .deudor = pQuery.fields(.ddl.deudor)
        .acreedor = pQuery.fields(.ddl.acreedor)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.negid & " = " & vNegid
    strSQL = strSQL & " AND " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " AND " & ddl.cuenta & " = " & vCuenta

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
    strSQL = strSQL & " WHERE " & ddl.negid & " = " & vNegid
    strSQL = strSQL & " AND " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " ORDER BY " & ddl.cuenta & " DESC"
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
    strSQL = strSQL & " SET " & ddl.deudor & " = '" & Replace(vDeudor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.acreedor & " = '" & Replace(vAcreedor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.negid & " = " & vNegid
    strSQL = strSQL & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & ", " & ddl.cuenta & " = " & vCuenta
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
        Set clon = Me.Clone
        clon.findByPrimaryKey pDB
    End If

    If vDeudor <> clon.Deudor Or IsNull(vDeudor) <> IsNull(clon.Deudor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.deudor & " = '" & Replace(vDeudor, ",", ".") & "'"
        blnChange = True
    End If
    If vAcreedor <> clon.Acreedor Or IsNull(vAcreedor) <> IsNull(clon.Acreedor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.acreedor & " = '" & Replace(vAcreedor, ",", ".") & "'"
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

    vCuenta = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.negid & " = " & vNegid
    strSQL = strSQL & " AND " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & " ORDER BY " & ddl.cuenta & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCuenta = 1 + rstQuery.fields(ddl.cuenta)

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

Dim objElemento As clsDAOConsfecha

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOConsfecha
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
Dim objElemento As clsDAOConsfecha

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
Dim objElemento As clsDAOConsfecha

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

