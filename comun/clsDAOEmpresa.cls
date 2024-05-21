VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vId As Integer
Private vNegID As Integer
Private vNombre As String
Private vRsocial As String
Private vDomicilio As String
Private vTelf As String
Private vCuit As String
Private vPuntovta As Integer
Private vIngbrutos As String
Private vNroestablecimiento As String
Private vSedetimbrado As String
Private vInicioactividades As String
Private vCondicioniva As String
Private vUbicacion As String

Private clon As clsDAOEmpresa

Public ddl As New clsDDLEmpresa

Private Sub class_Initialize()

    vId = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegID = 0
    vNombre = ""
    vRsocial = ""
    vDomicilio = ""
    vTelf = ""
    vCuit = ""
    vPuntovta = 0
    vIngbrutos = ""
    vNroestablecimiento = ""
    vSedetimbrado = ""
    vInicioactividades = ""
    vCondicioniva = ""
    vUbicacion = ""

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get id() As Variant

    id = vId

End Property

Public Property Let id(ByVal vNewValue As Variant)

On Error Resume Next

    vId = vNewValue

End Property

Public Property Get negID() As Variant

    negID = vNegID

End Property

Public Property Let negID(ByVal vNewValue As Variant)

On Error Resume Next

    vNegID = vNewValue

End Property

Public Property Get nombre() As Variant

    nombre = vNombre

End Property

Public Property Let nombre(ByVal vNewValue As Variant)

On Error Resume Next

    vNombre = vNewValue

End Property

Public Property Get rsocial() As Variant

    rsocial = vRsocial

End Property

Public Property Let rsocial(ByVal vNewValue As Variant)

On Error Resume Next

    vRsocial = vNewValue

End Property

Public Property Get domicilio() As Variant

    domicilio = vDomicilio

End Property

Public Property Let domicilio(ByVal vNewValue As Variant)

On Error Resume Next

    vDomicilio = vNewValue

End Property

Public Property Get telf() As Variant

    telf = vTelf

End Property

Public Property Let telf(ByVal vNewValue As Variant)

On Error Resume Next

    vTelf = vNewValue

End Property

Public Property Get cuit() As Variant

    cuit = vCuit

End Property

Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

Public Property Get puntovta() As Variant

    puntovta = vPuntovta

End Property

Public Property Let puntovta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntovta = vNewValue

End Property

Public Property Get ingbrutos() As Variant

    ingbrutos = vIngbrutos

End Property

Public Property Let ingbrutos(ByVal vNewValue As Variant)

On Error Resume Next

    vIngbrutos = vNewValue

End Property

Public Property Get nroestablecimiento() As Variant

    nroestablecimiento = vNroestablecimiento

End Property

Public Property Let nroestablecimiento(ByVal vNewValue As Variant)

On Error Resume Next

    vNroestablecimiento = vNewValue

End Property

Public Property Get sedetimbrado() As Variant

    sedetimbrado = vSedetimbrado

End Property

Public Property Let sedetimbrado(ByVal vNewValue As Variant)

On Error Resume Next

    vSedetimbrado = vNewValue

End Property

Public Property Get inicioactividades() As Variant

    inicioactividades = vInicioactividades

End Property

Public Property Let inicioactividades(ByVal vNewValue As Variant)

On Error Resume Next

    vInicioactividades = vNewValue

End Property

Public Property Get condicioniva() As Variant

    condicioniva = vCondicioniva

End Property

Public Property Let condicioniva(ByVal vNewValue As Variant)

On Error Resume Next

    vCondicioniva = vNewValue

End Property

Public Property Get ubicacion() As Variant

    ubicacion = vUbicacion

End Property

Public Property Let ubicacion(ByVal vNewValue As Variant)

On Error Resume Next

    vUbicacion = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vNegID

End Property

Public Property Get comboKey() As Variant

    comboKey = vId

End Property

Public Property Get listText() As Variant

    listText = vNegID

End Property

Public Property Get listKey() As Variant

    listKey = vId

End Property

Public Function clone() As clsDAOEmpresa
Dim objObjeto As New clsDAOEmpresa

    With objObjeto
        .id = vId
        .negID = vNegID
        .nombre = vNombre
        .rsocial = vRsocial
        .domicilio = vDomicilio
        .telf = vTelf
        .cuit = vCuit
        .puntovta = vPuntovta
        .ingbrutos = vIngbrutos
        .nroestablecimiento = vNroestablecimiento
        .sedetimbrado = vSedetimbrado
        .inicioactividades = vInicioactividades
        .condicioniva = vCondicioniva
        .ubicacion = vUbicacion
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOEmpresa, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .id = pQuery.Fields(.ddl.id)
        .negID = pQuery.Fields(.ddl.negID)
        .nombre = pQuery.Fields(.ddl.nombre)
        .rsocial = pQuery.Fields(.ddl.rsocial)
        .domicilio = pQuery.Fields(.ddl.domicilio)
        .telf = pQuery.Fields(.ddl.telf)
        .cuit = pQuery.Fields(.ddl.cuit)
        .puntovta = pQuery.Fields(.ddl.puntovta)
        .ingbrutos = pQuery.Fields(.ddl.ingbrutos)
        .nroestablecimiento = pQuery.Fields(.ddl.nroestablecimiento)
        .sedetimbrado = pQuery.Fields(.ddl.sedetimbrado)
        .inicioactividades = pQuery.Fields(.ddl.inicioactividades)
        .condicioniva = pQuery.Fields(.ddl.condicioniva)
        .ubicacion = pQuery.Fields(.ddl.ubicacion)
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

Public Sub findByPrimaryKey(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findLastID(pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    clean

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.id
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vId = rstQuery.Fields(ddl.id)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

Public Sub findLast(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.id & " DESC"
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
    strSQL = strSQL & " SET " & ddl.negID & " = " & vNegID
    strSQL = strSQL & ", " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
    strSQL = strSQL & ", " & ddl.rsocial & " = '" & replaceSpecialSymbols(vRsocial) & "'"
    strSQL = strSQL & ", " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
    strSQL = strSQL & ", " & ddl.telf & " = '" & replaceSpecialSymbols(vTelf) & "'"
    strSQL = strSQL & ", " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
    strSQL = strSQL & ", " & ddl.puntovta & " = " & vPuntovta
    strSQL = strSQL & ", " & ddl.ingbrutos & " = '" & replaceSpecialSymbols(vIngbrutos) & "'"
    strSQL = strSQL & ", " & ddl.nroestablecimiento & " = '" & replaceSpecialSymbols(vNroestablecimiento) & "'"
    strSQL = strSQL & ", " & ddl.sedetimbrado & " = '" & replaceSpecialSymbols(vSedetimbrado) & "'"
    strSQL = strSQL & ", " & ddl.inicioactividades & " = '" & replaceSpecialSymbols(vInicioactividades) & "'"
    strSQL = strSQL & ", " & ddl.condicioniva & " = '" & replaceSpecialSymbols(vCondicioniva) & "'"
    strSQL = strSQL & ", " & ddl.ubicacion & " = '" & replaceSpecialSymbols(vUbicacion) & "'"
    strSQL = strSQL & ";"

    add = pDB.execute(strSQL)

    findLastID pDB

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

    If vNegID <> clon.negID Or IsNull(vNegID) <> IsNull(clon.negID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negID & " = " & vNegID
        blnChange = True
    End If
    If vNombre <> clon.nombre Or IsNull(vNombre) <> IsNull(clon.nombre) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nombre & " = '" & replaceSpecialSymbols(vNombre) & "'"
        blnChange = True
    End If
    If vRsocial <> clon.rsocial Or IsNull(vRsocial) <> IsNull(clon.rsocial) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.rsocial & " = '" & replaceSpecialSymbols(vRsocial) & "'"
        blnChange = True
    End If
    If vDomicilio <> clon.domicilio Or IsNull(vDomicilio) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
        blnChange = True
    End If
    If vTelf <> clon.telf Or IsNull(vTelf) <> IsNull(clon.telf) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.telf & " = '" & replaceSpecialSymbols(vTelf) & "'"
        blnChange = True
    End If
    If vCuit <> clon.cuit Or IsNull(vCuit) <> IsNull(clon.cuit) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
        blnChange = True
    End If
    If vPuntovta <> clon.puntovta Or IsNull(vPuntovta) <> IsNull(clon.puntovta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntovta & " = " & vPuntovta
        blnChange = True
    End If
    If vIngbrutos <> clon.ingbrutos Or IsNull(vIngbrutos) <> IsNull(clon.ingbrutos) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ingbrutos & " = '" & replaceSpecialSymbols(vIngbrutos) & "'"
        blnChange = True
    End If
    If vNroestablecimiento <> clon.nroestablecimiento Or IsNull(vNroestablecimiento) <> IsNull(clon.nroestablecimiento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nroestablecimiento & " = '" & replaceSpecialSymbols(vNroestablecimiento) & "'"
        blnChange = True
    End If
    If vSedetimbrado <> clon.sedetimbrado Or IsNull(vSedetimbrado) <> IsNull(clon.sedetimbrado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.sedetimbrado & " = '" & replaceSpecialSymbols(vSedetimbrado) & "'"
        blnChange = True
    End If
    If vInicioactividades <> clon.inicioactividades Or IsNull(vInicioactividades) <> IsNull(clon.inicioactividades) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.inicioactividades & " = '" & replaceSpecialSymbols(vInicioactividades) & "'"
        blnChange = True
    End If
    If vCondicioniva <> clon.condicioniva Or IsNull(vCondicioniva) <> IsNull(clon.condicioniva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.condicioniva & " = '" & replaceSpecialSymbols(vCondicioniva) & "'"
        blnChange = True
    End If
    If vUbicacion <> clon.ubicacion Or IsNull(vUbicacion) <> IsNull(clon.ubicacion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ubicacion & " = '" & replaceSpecialSymbols(vUbicacion) & "'"
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

    vId = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.id & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vId = 1 + rstQuery.Fields(ddl.id)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOEmpresa

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOEmpresa
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
Dim objElemento As clsDAOEmpresa

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
Dim objElemento As clsDAOEmpresa

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

