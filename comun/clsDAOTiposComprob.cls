VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOTiposComprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vNegid As Integer
Private vCodigo As Long
Private vDescripcion As String
Private vModulo As Integer
Private vAplicapend As Integer
Private vCtacte As Integer
Private vDebita As Integer
Private vIva As Integer
Private vAplicable As Integer
Private vLibroiva As Integer
Private vTipocomprob As String
Private vRecibo As Integer
Private vContado As Integer
Private vPuntovta As Integer
Private vTipoAfip As Integer
Private vFactElect As Integer
Private vAsociado As Integer

Private clon As clsDAOTiposComprob

Public ddl As New clsDDLTiposComprob

Private Sub class_Initialize()

    vNegid = gEmpresa.negID
    vCodigo = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vDescripcion = ""
    vModulo = 0
    vAplicapend = 0
    vCtacte = 0
    vDebita = 0
    vIva = 0
    vAplicable = 0
    vLibroiva = 0
    vTipocomprob = ""
    vRecibo = 0
    vContado = 0
    vPuntovta = 0
    vTipoAfip = 0
    vFactElect = 0
    vAsociado = 0

End Sub

Public Sub clean()

    Call class_Initialize

End Sub

Public Property Get negID() As Variant

    negID = vNegid

End Property

Public Property Let negID(ByVal vNewValue As Variant)

On Error Resume Next

    vNegid = vNewValue

End Property

Public Property Get codigo() As Variant

    codigo = vCodigo

End Property

Public Property Let codigo(ByVal vNewValue As Variant)

On Error Resume Next

    vCodigo = vNewValue

End Property

Public Property Get descripcion() As Variant

    descripcion = vDescripcion

End Property

Public Property Let descripcion(ByVal vNewValue As Variant)

On Error Resume Next

    vDescripcion = vNewValue

End Property

Public Property Get modulo() As Variant

    modulo = vModulo

End Property

Public Property Let modulo(ByVal vNewValue As Variant)

On Error Resume Next

    vModulo = vNewValue

End Property

Public Property Get aplicapend() As Variant

    aplicapend = vAplicapend

End Property

Public Property Let aplicapend(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicapend = vNewValue

End Property

Public Property Get ctacte() As Variant

    ctacte = vCtacte

End Property

Public Property Let ctacte(ByVal vNewValue As Variant)

On Error Resume Next

    vCtacte = vNewValue

End Property

Public Property Get debita() As Variant

    debita = vDebita

End Property

Public Property Let debita(ByVal vNewValue As Variant)

On Error Resume Next

    vDebita = vNewValue

End Property

Public Property Get iva() As Variant

    iva = vIva

End Property

Public Property Let iva(ByVal vNewValue As Variant)

On Error Resume Next

    vIva = vNewValue

End Property

Public Property Get aplicable() As Variant

    aplicable = vAplicable

End Property

Public Property Let aplicable(ByVal vNewValue As Variant)

On Error Resume Next

    vAplicable = vNewValue

End Property

Public Property Get libroiva() As Variant

    libroiva = vLibroiva

End Property

Public Property Let libroiva(ByVal vNewValue As Variant)

On Error Resume Next

    vLibroiva = vNewValue

End Property

Public Property Get tipocomprob() As Variant

    tipocomprob = vTipocomprob

End Property

Public Property Let tipocomprob(ByVal vNewValue As Variant)

On Error Resume Next

    vTipocomprob = vNewValue

End Property

Public Property Get recibo() As Variant

    recibo = vRecibo

End Property

Public Property Let recibo(ByVal vNewValue As Variant)

On Error Resume Next

    vRecibo = vNewValue

End Property

Public Property Get contado() As Variant

    contado = vContado

End Property

Public Property Let contado(ByVal vNewValue As Variant)

On Error Resume Next

    vContado = vNewValue

End Property

Public Property Get puntovta() As Variant

    puntovta = vPuntovta

End Property

Public Property Let puntovta(ByVal vNewValue As Variant)

On Error Resume Next

    vPuntovta = vNewValue

End Property

Public Property Get tipoafip() As Variant

    tipoafip = vTipoAfip

End Property

Public Property Let tipoafip(ByVal vNewValue As Variant)

On Error Resume Next

    vTipoAfip = vNewValue

End Property

Public Property Get factelect() As Variant

    factelect = vFactElect

End Property

Public Property Let factelect(ByVal vNewValue As Variant)

On Error Resume Next

    vFactElect = vNewValue

End Property

Public Property Get asociado() As Variant

    asociado = vAsociado

End Property

Public Property Let asociado(ByVal vNewValue As Variant)

On Error Resume Next

    vAsociado = vNewValue

End Property

Public Property Get comboText() As Variant
Dim strPost As String
    
    strPost = IIf(vContado = 1, " Cdo", IIf(vCtacte = 1, " CtaCte", ""))
    comboText = vDescripcion & strPost

End Property

Public Property Get comboKey() As Variant

    comboKey = vCodigo

End Property

Public Property Get listText() As Variant

    listText = vDescripcion

End Property

Public Property Get listKey() As Variant

    listKey = vCodigo

End Property

Public Function clone() As clsDAOTiposComprob
Dim objObjeto As New clsDAOTiposComprob

    With objObjeto
        .negID = vNegid
        .codigo = vCodigo
        .descripcion = vDescripcion
        .modulo = vModulo
        .aplicapend = vAplicapend
        .ctacte = vCtacte
        .debita = vDebita
        .iva = vIva
        .aplicable = vAplicable
        .libroiva = vLibroiva
        .tipocomprob = vTipocomprob
        .recibo = vRecibo
        .contado = vContado
        .puntovta = vPuntovta
        .tipoafip = vTipoAfip
        .factelect = vFactElect
        .asociado = vAsociado
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOTiposComprob, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .negID = pQuery.Fields(.ddl.negID)
        .codigo = pQuery.Fields(.ddl.codigo)
        .descripcion = pQuery.Fields(.ddl.descripcion)
        .modulo = pQuery.Fields(.ddl.modulo)
        .aplicapend = pQuery.Fields(.ddl.aplicapend)
        .ctacte = pQuery.Fields(.ddl.ctacte)
        .debita = pQuery.Fields(.ddl.debita)
        .iva = pQuery.Fields(.ddl.iva)
        .aplicable = pQuery.Fields(.ddl.aplicable)
        .libroiva = pQuery.Fields(.ddl.libroiva)
        .tipocomprob = pQuery.Fields(.ddl.tipocomprob)
        .recibo = pQuery.Fields(.ddl.recibo)
        .contado = pQuery.Fields(.ddl.contado)
        .puntovta = pQuery.Fields(.ddl.puntovta)
        .tipoafip = pQuery.Fields(.ddl.tipoafip)
        .factelect = pQuery.Fields(.ddl.factelect)
        .asociado = pQuery.Fields(.ddl.asociado)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.negID & " = " & vNegid
    strSQL = strSQL & " AND " & ddl.codigo & " = " & vCodigo

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
    strSQL = strSQL & " WHERE " & ddl.negID & " = " & vNegid
    strSQL = strSQL & " ORDER BY " & ddl.codigo & " DESC"
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
    strSQL = strSQL & " SET " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
    strSQL = strSQL & ", " & ddl.modulo & " = " & vModulo
    strSQL = strSQL & ", " & ddl.aplicapend & " = " & vAplicapend
    strSQL = strSQL & ", " & ddl.ctacte & " = " & vCtacte
    strSQL = strSQL & ", " & ddl.debita & " = " & vDebita
    strSQL = strSQL & ", " & ddl.iva & " = " & vIva
    strSQL = strSQL & ", " & ddl.aplicable & " = " & vAplicable
    strSQL = strSQL & ", " & ddl.libroiva & " = " & vLibroiva
    strSQL = strSQL & ", " & ddl.tipocomprob & " = '" & replaceSpecialSymbols(vTipocomprob) & "'"
    strSQL = strSQL & ", " & ddl.recibo & " = " & vRecibo
    strSQL = strSQL & ", " & ddl.contado & " = " & vContado
    strSQL = strSQL & ", " & ddl.puntovta & " = " & vPuntovta
    strSQL = strSQL & ", " & ddl.tipoafip & " = " & vTipoAfip
    strSQL = strSQL & ", " & ddl.factelect & " = " & vFactElect
    strSQL = strSQL & ", " & ddl.asociado & " = " & vAsociado
    strSQL = strSQL & ", " & ddl.negID & " = " & vNegid
    strSQL = strSQL & ", " & ddl.codigo & " = " & vCodigo
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

    If vDescripcion <> clon.descripcion Or IsNull(vDescripcion) <> IsNull(clon.descripcion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descripcion & " = '" & replaceSpecialSymbols(vDescripcion) & "'"
        blnChange = True
    End If
    If vModulo <> clon.modulo Or IsNull(vModulo) <> IsNull(clon.modulo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.modulo & " = " & vModulo
        blnChange = True
    End If
    If vAplicapend <> clon.aplicapend Or IsNull(vAplicapend) <> IsNull(clon.aplicapend) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicapend & " = " & vAplicapend
        blnChange = True
    End If
    If vCtacte <> clon.ctacte Or IsNull(vCtacte) <> IsNull(clon.ctacte) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.ctacte & " = " & vCtacte
        blnChange = True
    End If
    If vDebita <> clon.debita Or IsNull(vDebita) <> IsNull(clon.debita) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.debita & " = " & vDebita
        blnChange = True
    End If
    If vIva <> clon.iva Or IsNull(vIva) <> IsNull(clon.iva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.iva & " = " & vIva
        blnChange = True
    End If
    If vAplicable <> clon.aplicable Or IsNull(vAplicable) <> IsNull(clon.aplicable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.aplicable & " = " & vAplicable
        blnChange = True
    End If
    If vLibroiva <> clon.libroiva Or IsNull(vLibroiva) <> IsNull(clon.libroiva) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.libroiva & " = " & vLibroiva
        blnChange = True
    End If
    If vTipocomprob <> clon.tipocomprob Or IsNull(vTipocomprob) <> IsNull(clon.tipocomprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipocomprob & " = '" & replaceSpecialSymbols(vTipocomprob) & "'"
        blnChange = True
    End If
    If vRecibo <> clon.recibo Or IsNull(vRecibo) <> IsNull(clon.recibo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.recibo & " = " & vRecibo
        blnChange = True
    End If
    If vContado <> clon.contado Or IsNull(vContado) <> IsNull(clon.contado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.contado & " = " & vContado
        blnChange = True
    End If
    If vPuntovta <> clon.puntovta Or IsNull(vPuntovta) <> IsNull(clon.puntovta) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.puntovta & " = " & vPuntovta
        blnChange = True
    End If
    If vTipoAfip <> clon.tipoafip Or IsNull(vTipoAfip) <> IsNull(clon.tipoafip) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipoafip & " = " & vTipoAfip
        blnChange = True
    End If
    If vFactElect <> clon.factelect Or IsNull(vFactElect) <> IsNull(clon.factelect) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.factelect & " = " & vFactElect
        blnChange = True
    End If
    If vAsociado <> clon.asociado Or IsNull(vAsociado) <> IsNull(clon.asociado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.asociado & " = " & vAsociado
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

    vCodigo = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.negID & " = " & vNegid
    strSQL = strSQL & " ORDER BY " & ddl.codigo & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCodigo = 1 + rstQuery.Fields(ddl.codigo)

    rstQuery.Close

    If pFind Then findByPrimaryKey pDB

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionVentas(pTipo As String, pExcluir As Integer, pNegID As Integer, pDB As clsDB) As Collection
Dim strSQL As String

    If pTipo = " " Then Exit Function

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.modulo & " = 3"
    If pExcluir <> 0 Then strSQL = strSQL & " AND " & ddl.codigo & " <> " & pExcluir
    strSQL = strSQL & " AND " & ddl.aplicapend & " = 0"
    strSQL = strSQL & " AND " & ddl.recibo & " = 0"
    strSQL = strSQL & " AND " & ddl.tipocomprob & " = '" & pTipo & "'"
    strSQL = strSQL & " AND " & ddl.negID & " = " & pNegID
    strSQL = strSQL & " AND " & ddl.factelect & " = 0"
    strSQL = strSQL & " ORDER BY " & ddl.descripcion
    strSQL = strSQL & ";"

    Set collectionVentas = collectionSQL(strSQL, pDB)

End Function

Public Function collectionVentasFE(pTipo As String, pExcluir As Integer, pNegID As Integer, pDB As clsDB) As Collection
Dim strSQL As String

    If pTipo = " " Then Exit Function

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.modulo & " = 3"
    If pExcluir <> 0 Then strSQL = strSQL & " AND " & ddl.codigo & " <> " & pExcluir
    strSQL = strSQL & " AND " & ddl.aplicapend & " = 0"
    strSQL = strSQL & " AND " & ddl.recibo & " = 0"
    strSQL = strSQL & " AND " & ddl.tipocomprob & " = '" & pTipo & "'"
    strSQL = strSQL & " AND " & ddl.negID & " = " & pNegID
    strSQL = strSQL & " AND " & ddl.factelect & " = 1"
    strSQL = strSQL & " ORDER BY " & ddl.descripcion
    strSQL = strSQL & ";"

    Set collectionVentasFE = collectionSQL(strSQL, pDB)

End Function

Public Function collectionRecibos(pTipoCompro As String, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.modulo & " = 3"
    strSQL = strSQL & " AND " & ddl.recibo & " <> 0"
    strSQL = strSQL & " AND " & ddl.tipocomprob & " IN ('" & pTipoCompro & "', 'X', 'Z')"
    strSQL = strSQL & " ORDER BY " & ddl.descripcion
    strSQL = strSQL & ";"

    Set collectionRecibos = collectionSQL(strSQL, pDB)

End Function

Public Function collectionModuloTipo(pModulo As Integer, pTipo As String, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.modulo & " = " & pModulo
    strSQL = strSQL & " AND (" & ddl.tipocomprob & " = '" & pTipo & "'"
    strSQL = strSQL & " OR " & ddl.tipocomprob & " IS Null"
    strSQL = strSQL & " OR " & ddl.tipocomprob & " = '')"
    strSQL = strSQL & ";"

    Set collectionModuloTipo = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOTiposComprob

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOTiposComprob
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

Public Sub fillComboVentas(pCombo As ComboBox, pTipo As String, pExcluir As Integer, pNegID As Integer, pDB As clsDB)
    
    fillCombo_ pCombo, collectionVentas(pTipo, pExcluir, pNegID, pDB)
    
End Sub

Public Sub fillComboVentasFE(pCombo As ComboBox, pTipo As String, pExcluir As Integer, pNegID As Integer, pDB As clsDB)
    
    fillCombo_ pCombo, collectionVentasFE(pTipo, pExcluir, pNegID, pDB)
    
End Sub

Public Sub fillComboRecibos(pCombo As ComboBox, pTipoCompro As String, pDB As clsDB)
    
    fillCombo_ pCombo, collectionRecibos(pTipoCompro, pDB)
    
End Sub

Public Sub fillComboModuloTipo(pCombo As ComboBox, pModulo As Integer, pTipo As String, pDB As clsDB)

    fillCombo_ pCombo, collectionModuloTipo(pModulo, pTipo, pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOTiposComprob

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
Dim objElemento As clsDAOTiposComprob

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

