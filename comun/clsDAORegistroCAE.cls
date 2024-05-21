VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAORegistroCAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTcoID As Integer
Private vPrefijo As Integer
Private vNroComprob As Long
Private vCliID As Long
Private vTotal As Currency
Private vExento As Currency
Private vNeto As Currency
Private vNeto105 As Currency
Private vIva As Currency
Private vIva105 As Currency
Private vCae As String
Private vFecha As String
Private vCaeVenc As String
Private vBarras As String
Private vTipodocumento As Integer
Private vNumerodocumento As Currency
Private vClientemovimientoIdasociado As Variant
Private vRecID As Long

Private clon As clsDAORegistroCAE

Public ddl As New clsDDLRegistroCAE

Private Sub class_Initialize()

    vTcoID = 0
    vPrefijo = 0
    vNroComprob = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vCliID = 0
    vTotal = 0
    vExento = 0
    vNeto = 0
    vNeto105 = 0
    vIva = 0
    vIva105 = 0
    vCae = ""
    vFecha = ""
    vCaeVenc = ""
    vBarras = ""
    vTipodocumento = 0
    vNumerodocumento = 0
    vClientemovimientoIdasociado = Null
    vRecID = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get tcoID() As Variant

    tcoID = vTcoID

End Property

Public Property Let tcoID(ByVal vNewValue As Variant)

On Error Resume Next

    vTcoID = vNewValue

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

Public Property Get cliID() As Variant

    cliID = vCliID

End Property

Public Property Let cliID(ByVal vNewValue As Variant)

On Error Resume Next

    vCliID = vNewValue

End Property

Public Property Get total() As Variant

    total = vTotal

End Property

Public Property Let total(ByVal vNewValue As Variant)

On Error Resume Next

    vTotal = vNewValue

End Property

Public Property Get exento() As Variant

    exento = vExento

End Property

Public Property Let exento(ByVal vNewValue As Variant)

On Error Resume Next

    vExento = vNewValue

End Property

Public Property Get neto() As Variant

    neto = vNeto

End Property

Public Property Let neto(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto = vNewValue

End Property

Public Property Get neto105() As Variant

    neto105 = vNeto105

End Property

Public Property Let neto105(ByVal vNewValue As Variant)

On Error Resume Next

    vNeto105 = vNewValue

End Property

Public Property Get iva() As Variant

    iva = vIva

End Property

Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

Public Property Get iva105() As Variant

    iva105 = vIva105

End Property

Public Property Let iva105(ByVal vNewValue As Variant)

On Error Resume Next

    vIva105 = vNewValue

End Property

Public Property Get cae() As Variant

    cae = vCae

End Property

Public Property Let cae(ByVal vNewValue As Variant)

On Error Resume Next

    vCae = vNewValue

End Property

Public Property Get fecha() As Variant

    fecha = vFecha

End Property

Public Property Let fecha(ByVal vNewValue As Variant)

On Error Resume Next

    vFecha = vNewValue

End Property

Public Property Get caevenc() As Variant

    caevenc = vCaeVenc

End Property

Public Property Let caevenc(ByVal vNewValue As Variant)

On Error Resume Next

    vCaeVenc = vNewValue

End Property

Public Property Get barras() As Variant

    barras = vBarras

End Property

Public Property Let barras(ByVal vNewValue As Variant)

On Error Resume Next

    vBarras = vNewValue

End Property

Public Property Get tipodocumento() As Variant

    tipodocumento = vTipodocumento

End Property

Public Property Let tipodocumento(ByVal vNewValue As Variant)

On Error Resume Next

    vTipodocumento = vNewValue

End Property

Public Property Get numerodocumento() As Variant

    numerodocumento = vNumerodocumento

End Property

Public Property Let numerodocumento(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerodocumento = vNewValue

End Property

Public Property Get clientemovimientoIdasociado() As Variant

    clientemovimientoIdasociado = vClientemovimientoIdasociado

End Property

Public Property Let clientemovimientoIdasociado(ByVal vNewValue As Variant)

On Error Resume Next

    vClientemovimientoIdasociado = vNewValue

End Property

Public Property Get recID() As Variant

    recID = vRecID

End Property

Public Property Let recID(ByVal vNewValue As Variant)

On Error Resume Next

    vRecID = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vCliID

End Property

Public Property Get comboKey() As Variant

    comboKey = vNroComprob

End Property

Public Property Get listText() As Variant

    listText = vCliID

End Property

Public Property Get listKey() As Variant

    listKey = vNroComprob

End Property

Public Function clone() As clsDAORegistroCAE
Dim objObjeto As New clsDAORegistroCAE

    With objObjeto
        .tcoID = vTcoID
        .prefijo = vPrefijo
        .nroComprob = vNroComprob
        .cliID = vCliID
        .total = vTotal
        .exento = vExento
        .neto = vNeto
        .neto105 = vNeto105
        .iva = vIva
        .iva105 = vIva105
        .cae = vCae
        .fecha = vFecha
        .caevenc = vCaeVenc
        .barras = vBarras
        .tipodocumento = vTipodocumento
        .numerodocumento = vNumerodocumento
        .clientemovimientoIdasociado = vClientemovimientoIdasociado
        .recID = vRecID
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAORegistroCAE, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .tcoID = pQuery.Fields(.ddl.tcoID)
        .prefijo = pQuery.Fields(.ddl.prefijo)
        .nroComprob = pQuery.Fields(.ddl.nroComprob)
        .cliID = pQuery.Fields(.ddl.cliID)
        .total = pQuery.Fields(.ddl.total)
        .exento = pQuery.Fields(.ddl.exento)
        .neto = pQuery.Fields(.ddl.neto)
        .neto105 = pQuery.Fields(.ddl.neto105)
        .iva = pQuery.Fields(.ddl.iva)
        .iva105 = pQuery.Fields(.ddl.iva105)
        .cae = pQuery.Fields(.ddl.cae)
        .fecha = pQuery.Fields(.ddl.fecha)
        .caevenc = pQuery.Fields(.ddl.caevenc)
        .barras = pQuery.Fields(.ddl.barras)
        .tipodocumento = pQuery.Fields(.ddl.tipodocumento)
        .numerodocumento = pQuery.Fields(.ddl.numerodocumento)
        .clientemovimientoIdasociado = pQuery.Fields(.ddl.clientemovimientoIdasociado)
        .recID = pQuery.Fields(.ddl.recID)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim sql As String

    sql = ""
    sql = sql & " WHERE " & ddl.tcoID & " = " & vTcoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " AND " & ddl.nroComprob & " = " & vNroComprob

    makeWherePrimaryKey = sql

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
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim sql As String

    initializeOther

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tcoID & " = " & vTcoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.nroComprob & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Sub findByRecID(Optional pDB As Variant)

    findByRecID_ IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub findByRecID_(pDB As clsDB)
Dim sql As String

Dim lngRecID As Long

    lngRecID = vRecID

    clean

    vRecID = lngRecID

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.recID & " = " & vRecID
    sql = sql & ";"

    findSQL sql, pDB

End Sub

Public Function delete(Optional pDB As Variant) As Boolean

    delete = delete_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function delete_(pDB As clsDB) As Boolean
Dim sql As String

    sql = "DELETE FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    delete_ = pDB.execute(sql)

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
Dim sql As String

Dim rstQuery As ADODB.Recordset

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)
    exist_ = Not rstQuery.EOF
    rstQuery.Close

End Function

Public Function add(Optional pDB As Variant) As Boolean

    add = add_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function add_(pDB As clsDB) As Boolean
Dim sql As String

    sql = "INSERT INTO " & ddl.table
    sql = sql & " SET " & ddl.cliID & " = " & vCliID
    sql = sql & ", " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
    sql = sql & ", " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
    sql = sql & ", " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
    sql = sql & ", " & ddl.neto105 & " = '" & Replace(vNeto105, ",", ".") & "'"
    sql = sql & ", " & ddl.iva & " = '" & Replace(vIva, ",", ".") & "'"
    sql = sql & ", " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
    sql = sql & ", " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
    sql = sql & ", " & ddl.fecha & " = '" & replaceSpecialSymbols(vFecha) & "'"
    sql = sql & ", " & ddl.caevenc & " = '" & replaceSpecialSymbols(vCaeVenc) & "'"
    sql = sql & ", " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
    sql = sql & ", " & ddl.tipodocumento & " = '" & tipodocumento & "'"
    sql = sql & ", " & ddl.numerodocumento & " = '" & numerodocumento & "'"
    sql = sql & ", " & ddl.clientemovimientoIdasociado & " = " & validateNullField(clientemovimientoIdasociado)
    sql = sql & ", " & ddl.tcoID & " = " & vTcoID
    sql = sql & ", " & ddl.prefijo & " = " & vPrefijo
    sql = sql & ", " & ddl.nroComprob & " = " & vNroComprob
    sql = sql & ";"

    add_ = pDB.execute(sql)

    findByPrimaryKey_ pDB

End Function

Public Function update(Optional pDB As Variant) As Boolean

    update = update_(IIf(IsMissing(pDB), db, pDB))

End Function

Private Function update_(pDB As clsDB) As Boolean
Dim sql As String
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
    If vTotal <> clon.total Or IsNull(vTotal) <> IsNull(clon.total) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.total & " = '" & Replace(vTotal, ",", ".") & "'"
        blnChange = True
    End If
    If vExento <> clon.exento Or IsNull(vExento) <> IsNull(clon.exento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.exento & " = '" & Replace(vExento, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto <> clon.neto Or IsNull(vNeto) <> IsNull(clon.neto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto & " = '" & Replace(vNeto, ",", ".") & "'"
        blnChange = True
    End If
    If vNeto105 <> clon.neto105 Or IsNull(vNeto105) <> IsNull(clon.neto105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.neto105 & " = '" & Replace(vNeto105, ",", ".") & "'"
        blnChange = True
    End If
    If vIva <> clon.iva Or IsNull(vIva) <> IsNull(clon.iva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva & " = '" & Replace(vIva, ",", ".") & "'"
        blnChange = True
    End If
    If vIva105 <> clon.iva105 Or IsNull(vIva105) <> IsNull(clon.iva105) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva105 & " = '" & Replace(vIva105, ",", ".") & "'"
        blnChange = True
    End If
    If vCae <> clon.cae Or IsNull(vCae) <> IsNull(clon.cae) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cae & " = '" & replaceSpecialSymbols(vCae) & "'"
        blnChange = True
    End If
    If vFecha <> clon.fecha Or IsNull(vFecha) <> IsNull(clon.fecha) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fecha & " = '" & replaceSpecialSymbols(vFecha) & "'"
        blnChange = True
    End If
    If vCaeVenc <> clon.caevenc Or IsNull(vCaeVenc) <> IsNull(clon.caevenc) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.caevenc & " = '" & replaceSpecialSymbols(vCaeVenc) & "'"
        blnChange = True
    End If
    If vBarras <> clon.barras Or IsNull(vBarras) <> IsNull(clon.barras) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.barras & " = '" & replaceSpecialSymbols(vBarras) & "'"
        blnChange = True
    End If
    If vTipodocumento <> clon.tipodocumento Or IsNull(vTipodocumento) <> IsNull(clon.tipodocumento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipodocumento & " = '" & vTipodocumento & "'"
        blnChange = True
    End If
    If vNumerodocumento <> clon.numerodocumento Or IsNull(vNumerodocumento) <> IsNull(clon.numerodocumento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.numerodocumento & " = '" & vNumerodocumento & "'"
        blnChange = True
    End If
    If vClientemovimientoIdasociado <> clon.clientemovimientoIdasociado Or IsNull(vClientemovimientoIdasociado) <> IsNull(clon.clientemovimientoIdasociado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clientemovimientoIdasociado & " = " & validateNullField(vClientemovimientoIdasociado)
        blnChange = True
    End If

    sql = "UPDATE " & ddl.table & " SET" & strSET
    sql = sql & makeWherePrimaryKey
    sql = sql & ";"

    If blnChange Then update_ = pDB.execute(sql)

    Set clon = Nothing
    Set clon = Me.clone

End Function

Public Sub newID(pFind As Boolean, Optional pDB As Variant)

    newID_ pFind, IIf(IsMissing(pDB), db, pDB)

End Sub

Private Sub newID_(pFind As Boolean, pDB As clsDB)
Dim sql As String

Dim rstQuery As ADODB.Recordset

    vNroComprob = 1

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & " WHERE " & ddl.tcoID & " = " & vTcoID
    sql = sql & " AND " & ddl.prefijo & " = " & vPrefijo
    sql = sql & " ORDER BY " & ddl.nroComprob & " DESC"
    sql = sql & " LIMIT 1"
    sql = sql & ";"

    Set rstQuery = pDB.query(sql)

    If Not rstQuery.EOF Then vNroComprob = 1 + rstQuery.Fields(ddl.nroComprob)

    rstQuery.Close

    If pFind Then findByPrimaryKey_ pDB

End Sub

Public Function collectionAll(Optional pDB As Variant) As Collection
Dim sql As String

    sql = "SELECT " & ddl.table & ".* FROM " & ddl.table
    sql = sql & ";"

    Set collectionAll = collectionSQL(sql, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), db, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAORegistroCAE

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAORegistroCAE
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
Dim objElemento As clsDAORegistroCAE

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
Dim objElemento As clsDAORegistroCAE

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

