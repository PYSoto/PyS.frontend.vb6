VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOArticulosAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vArtID As String
Private vAlias As String
Private vPrvID As Long
Private vPrecioCompra As Currency

Private clon As clsDAOArticulosAlias

Public ddl As New clsDDLArticulosAlias

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vArtID = ""
    vAlias = ""
    vPrvID = 0
    vPrecioCompra = 0

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

Public Property Get artID() As Variant

    artID = vArtID

End Property

Public Property Let artID(ByVal vNewValue As Variant)

On Error Resume Next

    vArtID = vNewValue

End Property

Public Property Get alias() As Variant

    alias = vAlias

End Property

Public Property Let alias(ByVal vNewValue As Variant)

On Error Resume Next

    vAlias = vNewValue

End Property

Public Property Get prvID() As Variant

    prvID = vPrvID

End Property

Public Property Let prvID(ByVal vNewValue As Variant)

On Error Resume Next

    vPrvID = vNewValue

End Property

Public Property Get precioCompra() As Variant

    precioCompra = vPrecioCompra

End Property

Public Property Let precioCompra(ByVal vNewValue As Variant)

On Error Resume Next

    vPrecioCompra = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vArtID

End Property

Public Property Get comboKey() As Variant

    comboKey = vClave

End Property

Public Property Get listText() As Variant

    listText = vArtID

End Property

Public Property Get listKey() As Variant

    listKey = vClave

End Property

Public Function clone() As clsDAOArticulosAlias
Dim objObjeto As New clsDAOArticulosAlias

    With objObjeto
        .clave = vClave
        .artID = vArtID
        .alias = vAlias
        .prvID = vPrvID
        .precioCompra = vPrecioCompra
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOArticulosAlias, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .artID = pQuery.Fields(.ddl.artID)
        .alias = pQuery.Fields(.ddl.alias)
        .prvID = pQuery.Fields(.ddl.prvID)
        .precioCompra = pQuery.Fields(.ddl.precioCompra)
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

Public Sub findByAlias(pAlias As String, pDB As clsDB)
Dim strSQL As String

    clean
    
    vAlias = pAlias
    
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.alias & " = '" & vAlias & "'"
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
    strSQL = strSQL & " SET " & ddl.artID & " = '" & replaceSpecialSymbols(vArtID) & "'"
    strSQL = strSQL & ", " & ddl.alias & " = '" & replaceSpecialSymbols(vAlias) & "'"
    strSQL = strSQL & ", " & ddl.prvID & " = " & vPrvID
    strSQL = strSQL & ", " & ddl.precioCompra & " = '" & Replace(vPrecioCompra, ",", ".") & "'"
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

    If vArtID <> clon.artID Or IsNull(vArtID) <> IsNull(clon.artID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.artID & " = '" & replaceSpecialSymbols(vArtID) & "'"
        blnChange = True
    End If
    If vAlias <> clon.alias Or IsNull(vAlias) <> IsNull(clon.alias) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.alias & " = '" & replaceSpecialSymbols(vAlias) & "'"
        blnChange = True
    End If
    If vPrvID <> clon.prvID Or IsNull(vPrvID) <> IsNull(clon.prvID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.prvID & " = " & vPrvID
        blnChange = True
    End If
    If vPrecioCompra <> clon.precioCompra Or IsNull(vPrecioCompra) <> IsNull(clon.precioCompra) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.precioCompra & " = '" & Replace(vPrecioCompra, ",", ".") & "'"
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

Public Function collectionByProveedor(pPrvID As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

Dim ar As New clsDDLArticulos

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " INNER JOIN " & ar.table
    strSQL = strSQL & " ON " & ddl.artid_t & " = " & ar.codigo_t
    strSQL = strSQL & " WHERE " & ddl.prvid_t & " = " & pPrvID
    strSQL = strSQL & " ORDER BY " & ar.descripcion
    strSQL = strSQL & ";"
    
    Set collectionByProveedor = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOArticulosAlias

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOArticulosAlias
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
Dim objElemento As clsDAOArticulosAlias

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
Dim objElemento As clsDAOArticulosAlias

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Function aumentarPrecios(pPorcentaje As Currency, pDB As clsDB) As Boolean
Dim objClonado As clsDAOArticulosAlias

Dim curPorcentaje As Currency

    If pPorcentaje = 0 Then Exit Function
    
    curPorcentaje = 1 + (pPorcentaje / 100)

    Set objClonado = Me.clone
    
    vPrecioCompra = objClonado.precioCompra * curPorcentaje
    
    aumentarPrecios = update(pDB)
    
End Function

Public Sub exportar(pArchivo As String, pDB As clsDB)
Dim strSQL As String
Dim strLinea As String

Dim rstQuery As ADODB.Recordset

    Open pArchivo For Append As #1

    strSQL = "SELECT CONCAT('INSERT INTO " & ddl.table
    strSQL = strSQL & "( " & ddl.artID
    strSQL = strSQL & ", " & ddl.alias
    strSQL = strSQL & ", " & ddl.prvID
    strSQL = strSQL & ", " & ddl.precioCompra
    strSQL = strSQL & ") VALUES "
    strSQL = strSQL & "( ''', " & ddl.artID & ", '''"
    strSQL = strSQL & ", ''', " & ddl.alias & ", '''"
    strSQL = strSQL & ", ', " & ddl.prvID & ", '"
    strSQL = strSQL & ", ''', " & ddl.precioCompra & ", '''"
    strSQL = strSQL & ");') AS cadena FROM " & ddl.table
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    
    Do While Not rstQuery.EOF
        Print #1, rstQuery!cadena
    
        rstQuery.MoveNext
    Loop
    
    rstQuery.Close
    
    Close #1
    
End Sub

Public Sub limpiar(pDB As clsDB)
Dim strSQL As String

    strSQL = "TRUNCATE " & ddl.table
    strSQL = strSQL & ";"
    
    pDB.execute strSQL
    
End Sub
