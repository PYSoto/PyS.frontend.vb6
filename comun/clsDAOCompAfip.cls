VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOCompAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vComprobanteID As Integer
Private vNombre As String
Private vEtiqueta As String

Private clon As clsDAOCompAfip

Public ddl As New clsDDLCompAfip

Private Sub class_Initialize()

    vComprobanteID = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNombre = ""
    vEtiqueta = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get comprobanteID() As Variant

    comprobanteID = vComprobanteID

End Property

Public Property Let comprobanteID(ByVal vNewValue As Variant)

On Error Resume Next

    vComprobanteID = vNewValue

End Property

Public Property Get nombre() As Variant

    nombre = vNombre

End Property

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get etiqueta() As Variant

    etiqueta = vEtiqueta

End Property

Public Property Let etiqueta(ByVal vNewValue As Variant)

On Error Resume Next

    vEtiqueta = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vNombre

End Property

Public Property Get comboKey() As Variant

    comboKey = vComprobanteID

End Property

Public Property Get listText() As Variant

    listText = vNombre

End Property

Public Property Get listKey() As Variant

    listKey = vComprobanteID

End Property

Public Function clone() As clsDAOCompAfip
Dim objObjeto As New clsDAOCompAfip

    With objObjeto
        .comprobanteID = vComprobanteID
        .nombre = vNombre
        .etiqueta = vEtiqueta
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOCompAfip, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .comprobanteID = pQuery.Fields(.ddl.comprobanteID)
        .nombre = pQuery.Fields(.ddl.nombre)
        .etiqueta = pQuery.Fields(.ddl.etiqueta)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.comprobanteID & " = " & vComprobanteID

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

Public Sub findByPrimaryKey(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findLast(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.comprobanteID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Function delete(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    delete = pDB.execute(strSQL)

End Function

Public Function save(pDB As clsDB) As Boolean

    If exist(pDB) Then
        save = update(pDB)
    Else
        save = add(pDB)
    End If

End Function

Public Function exist(pDB As clsDB) As Boolean
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)
    exist = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    strSQL = strSQL & ", " & ddl.etiqueta & " = '" & replaceSpecialSymbols(vEtiqueta) & "'"
    strSQL = strSQL & ", " & ddl.comprobanteID & " = " & vComprobanteID
    strSQL = strSQL & ";"

    add = pDB.execute(strSQL)

    findByPrimaryKey pDB

End Function

Public Function update(pDB As clsDB) As Boolean
Dim strSQL As String
Dim strSET As String

Dim blnChange As Boolean

    update = True

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
    If vEtiqueta <> clon.etiqueta Or IsNull(vEtiqueta) <> IsNull(clon.etiqueta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.etiqueta & " = '" & replaceSpecialSymbols(vEtiqueta) & "'"
        blnChange = True
    End If

    strSQL = "UPDATE " & ddl.table & " SET" & strSET
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    If blnChange Then update = pDB.execute(strSQL)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    vComprobanteID = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.comprobanteID & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vComprobanteID = 1 + rstQuery.Fields(ddl.comprobanteID)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionAny(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set collectionAny = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOCompAfip

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOCompAfip
        fillObject objElemento, rstQuery

        colRegistros.add objElemento

        rstQuery.MoveNext
    Loop

    rstQuery.Close

    Set collectionSQL = colRegistros

End Function

Public Sub fillCombo(pCombo As ComboBox, pDB As clsDB)

    fillCombo_ pCombo, collectionAll(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOCompAfip

    pCombo.Clear

    For Each objElemento In pCollection
        pCombo.AddItem objElemento.comboText
        pCombo.ItemData(pCombo.NewIndex) = objElemento.comboKey
    Next

    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0

End Sub

Public Sub fillList(pList As ListBox, pDB As clsDB)

    fillList_ pList, collectionAll(pDB)

End Sub

Private Sub fillList_(pList As ListBox, pCollection As Collection)
Dim objElemento As clsDAOCompAfip

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

