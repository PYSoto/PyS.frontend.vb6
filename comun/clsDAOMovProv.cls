VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOMovProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vCgoProv As Long
Private vCgoComprob As Integer
Private vPrefijo As Integer
Private vNroComprob As Long
Private vEmpID As Integer
Private vNegid As Integer
Private vFechaComprob As Date
Private vImporte As Currency
Private vNeto As Currency
Private vMontoIva As Currency
Private vMontoIva27 As Currency
Private vMontoIva105 As Currency
Private vPercIva As Currency
Private vPercIngBrutos As Currency
Private vGng As Currency
Private vClave As Long

Private clon As clsDAOMovProv

Public ddl As New clsDDLMovProv

Private Sub class_Initialize()

    vCgoProv = 0
    vCgoComprob = 0
    vPrefijo = 0
    vNroComprob = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vEmpID = 1
    vNegid = gEmpresa.negid
    vFechaComprob = Date
    vImporte = 0
    vNeto = 0
    vMontoIva = 0
    vMontoIva27 = 0
    vMontoIva105 = 0
    vPercIva = 0
    vPercIngBrutos = 0
    vGng = 0
    vClave = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get cgoProv() As Variant

    cgoProv = vCgoProv

End Property

Public Property Let cgoProv(ByVal vNewValue As Variant)

On Error Resume Next

    vCgoProv = vNewValue

End Property

Public Property Get cgoComprob() As Variant

    cgoComprob = vCgoComprob

End Property

Public Property Let cgoComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vCgoComprob = vNewValue

End Property

Public Property Get prefijo() As Variant

    prefijo = vPrefijo

End Property

Public Property Let prefijo(ByVal vNewValue As Variant)

On Error Resume Next

    vPrefijo = vNewValue

End Property

Public Property Get nroComprob() As Variant

    nroComprob = vNroComprob

End Property

Public Property Let nroComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

End Property

Public Property Get empID() As Variant

    empID = vEmpID

End Property

Public Property Let empID(ByVal vNewValue As Variant)

On Error Resume Next

    vEmpID = vNewValue

End Property

Public Property Get negid() As Variant

    negid = vNegid

End Property

Public Property Let negid(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get fechaComprob() As Variant

    fechaComprob = vFechaComprob

End Property

Public Property Let fechaComprob(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaComprob = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get neto() As Variant

    neto = vNeto

End Property

Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

Public Property Get montoIva() As Variant

    montoIva = vMontoIva

End Property

Public Property Let montoIva(ByVal vNewValue As Variant)

On Error Resume Next

    vMontoIva = vNewValue

End Property

Public Property Get montoIva27() As Variant

    montoIva27 = vMontoIva27

End Property

Public Property Let montoIva27(ByVal vNewValue As Variant)

On Error Resume Next

    vMontoIva27 = vNewValue

End Property

Public Property Get montoIva105() As Variant

    montoIva105 = vMontoIva105

End Property

Public Property Let montoIva105(ByVal vNewValue As Variant)

On Error Resume Next

    vMontoIva105 = vNewValue

End Property

Public Property Get percIva() As Variant

    percIva = vPercIva

End Property

Public Property Let percIva(ByVal vNewValue As Variant)

On Error Resume Next

    vPercIva = vNewValue

End Property

Public Property Get percIngBrutos() As Variant

    percIngBrutos = vPercIngBrutos

End Property

Public Property Let percIngBrutos(ByVal vNewValue As Variant)

On Error Resume Next

    vPercIngBrutos = vNewValue

End Property

Public Property Get gng() As Variant

    gng = vGng

End Property

Public Property Let gng(ByVal vNewValue As Variant)

On Error Resume Next

    vGng = vNewValue

End Property

Public Property Get clave() As Variant

    clave = vClave

End Property

Public Property Let clave(ByVal vNewValue As Variant)

On Error Resume Next

    vClave = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vEmpID

End Property

Public Property Get comboKey() As Variant

    comboKey = vNroComprob

End Property

Public Property Get listText() As Variant

    listText = vEmpID

End Property

Public Property Get listKey() As Variant

    listKey = vNroComprob

End Property

Public Function clone() As clsDAOMovProv
Dim objObjeto As New clsDAOMovProv

    With objObjeto
        .cgoProv = vCgoProv
        .cgoComprob = vCgoComprob
        .prefijo = vPrefijo
        .nroComprob = vNroComprob
        .empID = vEmpID
        .negid = vNegid
        .fechaComprob = vFechaComprob
        .importe = vImporte
        .neto = vNeto
        .montoIva = vMontoIva
        .montoIva27 = vMontoIva27
        .montoIva105 = vMontoIva105
        .percIva = vPercIva
        .percIngBrutos = vPercIngBrutos
        .gng = vGng
        .clave = vClave
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOMovProv, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .cgoProv = pQuery.Fields(.ddl.cgoProv)
        .cgoComprob = pQuery.Fields(.ddl.cgoComprob)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .empID = pQuery.Fields(.ddl.empID)
        .negid = pQuery.Fields(.ddl.negid)
        .fechaComprob = pQuery.Fields(.ddl.fechaComprob)
        .importe = pQuery.Fields(.ddl.importe)
        .neto = pQuery.Fields(.ddl.neto)
        .montoIva = pQuery.Fields(.ddl.montoIva)
        .montoIva27 = pQuery.Fields(.ddl.montoIva27)
        .montoIva105 = pQuery.Fields(.ddl.montoIva105)
        .percIva = pQuery.Fields(.ddl.percIva)
        .percIngBrutos = pQuery.Fields(.ddl.percIngBrutos)
        .gng = pQuery.Fields(.ddl.gng)
        .clave = pQuery.Fields(.ddl.clave)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.cgoProv & " = " & vCgoProv
    strSQL = strSQL & " AND " & ddl.cgoComprob & " = " & vCgoComprob
    strSQL = strSQL & " AND " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & " AND " & ddl.nroComprob & " = " & vNroComprob

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

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoProv & " = " & vCgoProv
    strSQL = strSQL & " AND " & ddl.cgoComprob & " = " & vCgoComprob
    strSQL = strSQL & " AND " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & " ORDER BY " & ddl.nroComprob & " DESC"
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
    strSQL = strSQL & " SET " & ddl.empID & " = " & vEmpID
    strSQL = strSQL & ", " & ddl.negid & " = " & vNegid
    strSQL = strSQL & ", " & ddl.fechaComprob & " = " & fechaDB(vFechaComprob)
    strSQL = strSQL & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.montoIva & " = '" & Replace(vMontoIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.montoIva27 & " = '" & Replace(vMontoIva27, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.montoIva105 & " = '" & Replace(vMontoIva105, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.percIva & " = '" & Replace(vPercIva, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.percIngBrutos & " = '" & Replace(vPercIngBrutos, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.gng & " = '" & Replace(vGng, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.cgoProv & " = " & vCgoProv
    strSQL = strSQL & ", " & ddl.cgoComprob & " = " & vCgoComprob
    strSQL = strSQL & ", " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & ", " & ddl.nroComprob & " = " & vNroComprob
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

    If vEmpID <> clon.empID Or IsNull(vEmpID) <> IsNull(clon.empID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.empID & " = " & vEmpID
        blnChange = True
    End If
    If vNegid <> clon.negid Or IsNull(vNegid) <> IsNull(clon.negid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negid & " = " & vNegid
        blnChange = True
    End If
    If vFechaComprob <> clon.fechaComprob Or IsNull(vFechaComprob) <> IsNull(clon.fechaComprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaComprob & " = " & fechaDB(vFechaComprob)
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto <> clon.neto Or IsNull(vNeto) <> IsNull(clon.neto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
        blnChange = True
    End If
    If vMontoIva <> clon.montoIva Or IsNull(vMontoIva) <> IsNull(clon.montoIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.montoIva & " = '" & Replace(vMontoIva, ",", ".") & "'"
        blnChange = True
    End If
    If vMontoIva27 <> clon.montoIva27 Or IsNull(vMontoIva27) <> IsNull(clon.montoIva27) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.montoIva27 & " = '" & Replace(vMontoIva27, ",", ".") & "'"
        blnChange = True
    End If
    If vMontoIva105 <> clon.montoIva105 Or IsNull(vMontoIva105) <> IsNull(clon.montoIva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.montoIva105 & " = '" & Replace(vMontoIva105, ",", ".") & "'"
        blnChange = True
    End If
    If vPercIva <> clon.percIva Or IsNull(vPercIva) <> IsNull(clon.percIva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.percIva & " = '" & Replace(vPercIva, ",", ".") & "'"
        blnChange = True
    End If
    If vPercIngBrutos <> clon.percIngBrutos Or IsNull(vPercIngBrutos) <> IsNull(clon.percIngBrutos) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.percIngBrutos & " = '" & Replace(vPercIngBrutos, ",", ".") & "'"
        blnChange = True
    End If
    If vGng <> clon.gng Or IsNull(vGng) <> IsNull(clon.gng) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.gng & " = '" & Replace(vGng, ",", ".") & "'"
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

    vNroComprob = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoProv & " = " & vCgoProv
    strSQL = strSQL & " AND " & ddl.cgoComprob & " = " & vCgoComprob
    strSQL = strSQL & " AND " & ddl.prefijo & " = " & vPrefijo
    strSQL = strSQL & " ORDER BY " & ddl.nroComprob & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vNroComprob = 1 + rstQuery.Fields(ddl.nroComprob)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionByCgoProv(pCgoProv As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cgoProv & " = " & pCgoProv
    strSQL = strSQL & " ORDER BY " & ddl.fechaComprob
    strSQL = strSQL & ";"

    Set collectionByCgoProv = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOMovProv

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOMovProv
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
Dim objElemento As clsDAOMovProv

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
Dim objElemento As clsDAOMovProv

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

