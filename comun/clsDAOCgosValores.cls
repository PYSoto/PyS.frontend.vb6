VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOCgosValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vNegid As Integer
Private vCodigo As Integer
Private vConcepto As String
Private vNumerable As Integer
Private vDuplicados As Integer
Private vFechaemi As Integer
Private vFechavto As Integer
Private vTitular As Integer
Private vBanco As Integer
Private vChtercero As Integer
Private vCtacte As Integer

Private clon As clsDAOCgosValores

Public ddl As New clsDDLCgosValores

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegid = 0
    vCodigo = 0
    vConcepto = ""
    vNumerable = 0
    vDuplicados = 0
    vFechaemi = 0
    vFechavto = 0
    vTitular = 0
    vBanco = 0
    vChtercero = 0
    vCtacte = 0

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

Public Property Get concepto() As Variant

    concepto = vConcepto

End Property

Public Property Let concepto(ByVal vNewValue As Variant)

On Error Resume Next

    vConcepto = vNewValue

End Property

Public Property Get numerable() As Variant

    numerable = vNumerable

End Property

Public Property Let numerable(ByVal vNewValue As Variant)

On Error Resume Next

    vNumerable = vNewValue

End Property

Public Property Get duplicados() As Variant

    duplicados = vDuplicados

End Property

Public Property Let duplicados(ByVal vNewValue As Variant)

On Error Resume Next

    vDuplicados = vNewValue

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

Public Property Get chtercero() As Variant

    chtercero = vChtercero

End Property

Public Property Let chtercero(ByVal vNewValue As Variant)

On Error Resume Next

    vChtercero = vNewValue

End Property

Public Property Get ctacte() As Variant

    ctacte = vCtacte

End Property

Public Property Let ctacte(ByVal vNewValue As Variant)

On Error Resume Next

    vCtacte = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vConcepto

End Property

Public Property Get comboKey() As Variant

    comboKey = vCodigo

End Property

Public Property Get listText() As Variant

    listText = vConcepto

End Property

Public Property Get listKey() As Variant

    listKey = vCodigo

End Property

Public Function clone() As clsDAOCgosValores
Dim objObjeto As New clsDAOCgosValores

    With objObjeto
        .clave = vClave
        .negid = vNegid
        .codigo = vCodigo
        .concepto = vConcepto
        .numerable = vNumerable
        .duplicados = vDuplicados
        .fechaemi = vFechaemi
        .fechavto = vFechavto
        .titular = vTitular
        .banco = vBanco
        .chtercero = vChtercero
        .ctacte = vCtacte
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOCgosValores, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .negid = pQuery.Fields(.ddl.negid)
        .codigo = pQuery.Fields(.ddl.codigo)
        .concepto = pQuery.Fields(.ddl.concepto)
        .numerable = pQuery.Fields(.ddl.numerable)
        .duplicados = pQuery.Fields(.ddl.duplicados)
        .fechaemi = pQuery.Fields(.ddl.fechaemi)
        .fechavto = pQuery.Fields(.ddl.fechavto)
        .titular = pQuery.Fields(.ddl.titular)
        .banco = pQuery.Fields(.ddl.banco)
        .chtercero = pQuery.Fields(.ddl.chtercero)
        .ctacte = pQuery.Fields(.ddl.ctacte)
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
    strSQL = strSQL & ", " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
    strSQL = strSQL & ", " & ddl.numerable & " = " & vNumerable
    strSQL = strSQL & ", " & ddl.duplicados & " = " & vDuplicados
    strSQL = strSQL & ", " & ddl.fechaemi & " = " & vFechaemi
    strSQL = strSQL & ", " & ddl.fechavto & " = " & vFechavto
    strSQL = strSQL & ", " & ddl.titular & " = " & vTitular
    strSQL = strSQL & ", " & ddl.banco & " = " & vBanco
    strSQL = strSQL & ", " & ddl.chtercero & " = " & vChtercero
    strSQL = strSQL & ", " & ddl.ctacte & " = " & vCtacte
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
    If vConcepto <> clon.concepto Or IsNull(vConcepto) <> IsNull(clon.concepto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.concepto & " = '" & replaceSpecialSymbols(vConcepto) & "'"
        blnChange = True
    End If
    If vNumerable <> clon.numerable Or IsNull(vNumerable) <> IsNull(clon.numerable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.numerable & " = " & vNumerable
        blnChange = True
    End If
    If vDuplicados <> clon.duplicados Or IsNull(vDuplicados) <> IsNull(clon.duplicados) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.duplicados & " = " & vDuplicados
        blnChange = True
    End If
    If vFechaemi <> clon.fechaemi Or IsNull(vFechaemi) <> IsNull(clon.fechaemi) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechaemi & " = " & vFechaemi
        blnChange = True
    End If
    If vFechavto <> clon.fechavto Or IsNull(vFechavto) <> IsNull(clon.fechavto) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fechavto & " = " & vFechavto
        blnChange = True
    End If
    If vTitular <> clon.titular Or IsNull(vTitular) <> IsNull(clon.titular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.titular & " = " & vTitular
        blnChange = True
    End If
    If vBanco <> clon.banco Or IsNull(vBanco) <> IsNull(clon.banco) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.banco & " = " & vBanco
        blnChange = True
    End If
    If vChtercero <> clon.chtercero Or IsNull(vChtercero) <> IsNull(clon.chtercero) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.chtercero & " = " & vChtercero
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

Dim objElemento As clsDAOCgosValores

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOCgosValores
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
Dim objElemento As clsDAOCgosValores

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
Dim objElemento As clsDAOCgosValores

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

