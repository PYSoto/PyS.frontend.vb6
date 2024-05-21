VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vPreid As Long
Private vCliID As Long
Private vFecha As Date
Private vFechaVto As Date
Private vObservac As String
Private vClavemovclie As Long
Private vCtacte As Integer

Private clon As clsDAOPresupuesto

Public ddl As New clsDDLPresupuesto

Private Sub class_Initialize()

    vPreid = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vCliID = 0
    vFecha = Date
    vFechaVto = Date
    vObservac = ""
    vClavemovclie = 0
    vCtacte = 0

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

Public Property Get cliID() As Variant

    cliID = vCliID

End Property

Public Property Let cliID(ByVal vNewValue As Variant)

On Error Resume Next

    vCliID = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get fechavto() As Variant

    fechavto = vFechaVto

End Property

Public Property Let fechavto(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaVto = vNewValue

End Property

Public Property Get observac() As Variant

    observac = vObservac

End Property

Public Property Let observac(ByVal vNewValue As Variant)

On Error Resume Next

    vObservac = vNewValue

End Property

Public Property Get clavemovclie() As Variant

    clavemovclie = vClavemovclie

End Property

Public Property Let clavemovclie(ByVal vNewValue As Variant)

On Error Resume Next

    vClavemovclie = vNewValue

End Property

Public Property Get ctacte() As Variant

    ctacte = vCtacte

End Property

Public Property Let ctacte(ByVal vNewValue As Variant)

On Error Resume Next

    vCtacte = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vCliID

End Property

Public Property Get comboKey() As Variant

    comboKey = vPreid

End Property

Public Property Get listText() As Variant

    listText = vCliID

End Property

Public Property Get listKey() As Variant

    listKey = vPreid

End Property

Public Function clone() As clsDAOPresupuesto
Dim objObjeto As New clsDAOPresupuesto

    With objObjeto
        .preID = vPreid
        .cliID = vCliID
        .fecha = vFecha
        .fechavto = vFechaVto
        .observac = vObservac
        .clavemovclie = vClavemovclie
        .ctacte = vCtacte
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOPresupuesto, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .preID = pQuery.Fields(.ddl.preID)
        .cliID = pQuery.Fields(.ddl.cliID)
        .fecha = pQuery.Fields(.ddl.fecha)
        .fechavto = pQuery.Fields(.ddl.fechavto)
        .observac = pQuery.Fields(.ddl.observac)
        .clavemovclie = pQuery.Fields(.ddl.clavemovclie)
        .ctacte = pQuery.Fields(.ddl.ctacte)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.preID & " = " & vPreid

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

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.preID
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vPreid = rstQuery.Fields(ddl.preID)

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
    strSQL = strSQL & " ORDER BY " & ddl.preID & " DESC"
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
    strSQL = strSQL & " SET " & ddl.cliID & " = " & vCliID
    strSQL = strSQL & ", " & ddl.fecha & " = " & fechaDB(vFecha)
    strSQL = strSQL & ", " & ddl.fechavto & " = " & fechaDB(vFechaVto)
    strSQL = strSQL & ", " & ddl.observac & " = '" & replaceSpecialSymbols(vObservac) & "'"
    strSQL = strSQL & ", " & ddl.clavemovclie & " = " & vClavemovclie
    strSQL = strSQL & ", " & ddl.ctacte & " = " & vCtacte
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

    If vCliID <> clon.cliID Or IsNull(vCliID) <> IsNull(clon.cliID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cliID & " = " & vCliID
        blnChange = True
    End If
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = " & fechaDB(vFecha)
        blnChange = True
    End If
    If vFechaVto <> clon.fechavto Or IsNull(vFechaVto) <> IsNull(clon.fechavto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechavto & " = " & fechaDB(vFechaVto)
        blnChange = True
    End If
    If vObservac <> clon.observac Or IsNull(vObservac) <> IsNull(clon.observac) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.observac & " = '" & replaceSpecialSymbols(vObservac) & "'"
        blnChange = True
    End If
    If vClavemovclie <> clon.clavemovclie Or IsNull(vClavemovclie) <> IsNull(clon.clavemovclie) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clavemovclie & " = " & vClavemovclie
        blnChange = True
    End If
    If vCtacte <> clon.ctacte Or IsNull(vCtacte) <> IsNull(clon.ctacte) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ctacte & " = " & vCtacte
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

    vPreid = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.preID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vPreid = 1 + rstQuery.Fields(ddl.preID)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionPendientes(pDias As Integer, pDB As clsDB, Optional pFacturados As Integer = 0) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE DATE_ADD(" & ddl.fecha & ", INTERVAL " & pDias & " DAY) >= NOW()"
    If pFacturados = 0 Then strSQL = strSQL & " AND " & ddl.clavemovclie & " = 0"
    strSQL = strSQL & " ORDER BY " & ddl.fechavto & " DESC, " & ddl.cliID
    strSQL = strSQL & ";"

    Set collectionPendientes = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOPresupuesto

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOPresupuesto
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
Dim objElemento As clsDAOPresupuesto

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
Dim objElemento As clsDAOPresupuesto

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

