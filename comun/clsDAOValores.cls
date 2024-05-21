VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vNegid As Integer
Private vCodigo As Integer
Private vCgocli As Long
Private vFechaemi As Variant
Private vFechavto As Variant
Private vTcoID As Integer
Private vNroComprob As Long
Private vImporte As Currency
Private vFechaReg As Date
Private vClavemovv As Long
Private vTitular As String
Private vBanco As String

Private clon As clsDAOValores

Public ddl As New clsDDLValores

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegid = gEmpresa.negid
    vCodigo = 0
    vCgocli = 0
    vFechaemi = Null
    vFechavto = Null
    vTcoID = 0
    vNroComprob = 0
    vImporte = 0
    vFechaReg = Date
    vClavemovv = 0
    vTitular = ""
    vBanco = ""

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

Public Property Get negid() As Variant

    negid = vNegid

End Property

Public Property Let negid(ByVal vNewValue As Variant)

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

Public Property Get cgocli() As Variant

    cgocli = vCgocli

End Property

Public Property Let cgocli(ByVal vNewValue As Variant)

On Error Resume Next

    vCgocli = vNewValue

End Property

Public Property Get fechaemi() As Variant

    fechaemi = vFechaemi

End Property

Public Property Let fechaemi(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaemi = vNewValue

End Property

Public Property Get fechavto() As Variant

    fechavto = vFechavto

End Property

Public Property Let fechavto(ByVal vNewValue As Variant)

On Error Resume Next

    vFechavto = vNewValue

End Property

Public Property Get tcoid() As Variant

    tcoid = vTcoID

End Property

Public Property Let tcoid(ByVal vNewValue As Variant)

On Error Resume Next

    vTcoID = vNewValue

End Property

Public Property Get nrocomprob() As Variant

    nrocomprob = vNroComprob

End Property

Public Property Let nrocomprob(ByVal vNewValue As Variant)

On Error Resume Next

    vNroComprob = vNewValue

End Property

Public Property Get importe() As Variant

    importe = vImporte

End Property

Public Property Let importe(ByVal vNewValue As Variant)

On Error Resume Next

    vImporte = vNewValue

End Property

Public Property Get fechaReg() As Variant

    fechaReg = vFechaReg

End Property

Public Property Let fechaReg(ByVal vNewValue As Variant)

On Error Resume Next

    vFechaReg = vNewValue

End Property

Public Property Get clavemovv() As Variant

    clavemovv = vClavemovv

End Property

Public Property Let clavemovv(ByVal vNewValue As Variant)

On Error Resume Next

    vClavemovv = vNewValue

End Property

Public Property Get titular() As Variant

    titular = vTitular

End Property

Public Property Let titular(ByVal vNewValue As Variant)

On Error Resume Next

    vTitular = vNewValue

End Property

Public Property Get banco() As Variant

    banco = vBanco

End Property

Public Property Let banco(ByVal vNewValue As Variant)

On Error Resume Next

    vBanco = vNewValue

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

Public Function clone() As clsDAOValores
Dim objObjeto As New clsDAOValores

    With objObjeto
        .clave = vClave
        .negid = vNegid
        .codigo = vCodigo
        .cgocli = vCgocli
        .fechaemi = vFechaemi
        .fechavto = vFechavto
        .tcoid = vTcoID
        .nrocomprob = vNroComprob
        .importe = vImporte
        .fechaReg = vFechaReg
        .clavemovv = vClavemovv
        .titular = vTitular
        .banco = vBanco
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOValores, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .negid = pQuery.Fields(.ddl.negid)
        .codigo = pQuery.Fields(.ddl.codigo)
        .cgocli = pQuery.Fields(.ddl.cgocli)
        .fechaemi = pQuery.Fields(.ddl.fechaemi)
        .fechavto = pQuery.Fields(.ddl.fechavto)
        .tcoid = pQuery.Fields(.ddl.tcoid)
        .nrocomprob = pQuery.Fields(.ddl.nrocomprob)
        .importe = pQuery.Fields(.ddl.importe)
        .fechaReg = pQuery.Fields(.ddl.fechaReg)
        .clavemovv = pQuery.Fields(.ddl.clavemovv)
        .titular = pQuery.Fields(.ddl.titular)
        .banco = pQuery.Fields(.ddl.banco)
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
    strSQL = strSQL & ", " & ddl.codigo & " = " & vCodigo
    strSQL = strSQL & ", " & ddl.cgocli & " = " & vCgocli
    strSQL = strSQL & ", " & ddl.fechaemi & " = " & fechaDB(vFechaemi)
    strSQL = strSQL & ", " & ddl.fechavto & " = " & fechaDB(vFechavto)
    strSQL = strSQL & ", " & ddl.tcoid & " = " & vTcoID
    strSQL = strSQL & ", " & ddl.nrocomprob & " = " & vNroComprob
    strSQL = strSQL & ", " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.fechaReg & " = " & fechaDB(vFechaReg)
    strSQL = strSQL & ", " & ddl.clavemovv & " = " & vClavemovv
    strSQL = strSQL & ", " & ddl.titular & " = '" & replaceSpecialSymbols(vTitular) & "'"
    strSQL = strSQL & ", " & ddl.banco & " = '" & replaceSpecialSymbols(vBanco) & "'"
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
    If vCodigo <> clon.codigo Or IsNull(vCodigo) <> IsNull(clon.codigo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.codigo & " = " & vCodigo
        blnChange = True
    End If
    If vCgocli <> clon.cgocli Or IsNull(vCgocli) <> IsNull(clon.cgocli) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cgocli & " = " & vCgocli
        blnChange = True
    End If
    If vFechaemi <> clon.fechaemi Or IsNull(vFechaemi) <> IsNull(clon.fechaemi) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaemi & " = " & fechaDB(vFechaemi)
        blnChange = True
    End If
    If vFechavto <> clon.fechavto Or IsNull(vFechavto) <> IsNull(clon.fechavto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechavto & " = " & fechaDB(vFechavto)
        blnChange = True
    End If
    If vTcoID <> clon.tcoid Or IsNull(vTcoID) <> IsNull(clon.tcoid) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tcoid & " = " & vTcoID
        blnChange = True
    End If
    If vNroComprob <> clon.nrocomprob Or IsNull(vNroComprob) <> IsNull(clon.nrocomprob) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nrocomprob & " = " & vNroComprob
        blnChange = True
    End If
    If vImporte <> clon.importe Or IsNull(vImporte) <> IsNull(clon.importe) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.importe & " = '" & Replace(vImporte, ",", ".") & "'"
        blnChange = True
    End If
    If vFechaReg <> clon.fechaReg Or IsNull(vFechaReg) <> IsNull(clon.fechaReg) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaReg & " = " & fechaDB(vFechaReg)
        blnChange = True
    End If
    If vClavemovv <> clon.clavemovv Or IsNull(vClavemovv) <> IsNull(clon.clavemovv) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.clavemovv & " = " & vClavemovv
        blnChange = True
    End If
    If vTitular <> clon.titular Or IsNull(vTitular) <> IsNull(clon.titular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.titular & " = '" & replaceSpecialSymbols(vTitular) & "'"
        blnChange = True
    End If
    If vBanco <> clon.banco Or IsNull(vBanco) <> IsNull(clon.banco) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.banco & " = '" & replaceSpecialSymbols(vBanco) & "'"
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

Public Function collectionByClaveMovClie(pClaveMovClie As Long, Optional pDB As Variant) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.clavemovv & " = " & pClaveMovClie
    strSQL = strSQL & ";"

    Set collectionByClaveMovClie = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, Optional pDB As Variant) As Collection

    Set collectionSQL = collection_(pSQL, IIf(IsMissing(pDB), DB, pDB))

End Function

Private Function collection_(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOValores

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOValores
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
Dim objElemento As clsDAOValores

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
Dim objElemento As clsDAOValores

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

