VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vNegid As Integer
Private vCodigo As Long
Private vCuit As String
Private vRazon As String
Private vDomicilio As String
Private vLocalidad As String
Private vProvincia As String
Private vTel As String
Private vFax As String
Private vEmail As String
Private vCelular As String
Private vPosicion As Integer
Private vTipodoc As String
Private vNrodoc As Long
Private vLimitecredito As Currency
Private vNacionalidad As String
Private vDescuento As Currency
Private vFacturable As Integer

Private clon As clsDAOClientes

Public ddl As New clsDDLClientes

Private Sub class_Initialize()

    vClave = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vNegid = gEmpresa.negID
    vCodigo = 0
    vCuit = "00-00000000-0"
    vRazon = ""
    vDomicilio = ""
    vLocalidad = ""
    vProvincia = ""
    vTel = ""
    vFax = ""
    vEmail = ""
    vCelular = ""
    vPosicion = 0
    vTipodoc = ""
    vNrodoc = 0
    vLimitecredito = 0
    vNacionalidad = ""
    vDescuento = 0
    vFacturable = 1

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

Public Property Get codigoBuscar() As Long

    codigoBuscar = vCodigo
    
End Property

Public Property Get cuit() As Variant

    cuit = vCuit

End Property

Public Property Let cuit(ByVal vNewValue As Variant)

On Error Resume Next

    vCuit = vNewValue

End Property

Public Property Get razon() As Variant

    razon = vRazon

End Property

Public Property Let razon(ByVal vNewValue As Variant)

On Error Resume Next

    vRazon = vNewValue

End Property

Public Property Get domicilio() As Variant

    domicilio = vDomicilio

End Property

Public Property Let domicilio(ByVal vNewValue As Variant)

On Error Resume Next

    vDomicilio = vNewValue

End Property

Public Property Get localidad() As Variant

    localidad = vLocalidad

End Property

Public Property Let localidad(ByVal vNewValue As Variant)

On Error Resume Next

    vLocalidad = vNewValue

End Property

Public Property Get provincia() As Variant

    provincia = vProvincia

End Property

Public Property Let provincia(ByVal vNewValue As Variant)

On Error Resume Next

    vProvincia = vNewValue

End Property

Public Property Get tel() As Variant

    tel = vTel

End Property

Public Property Let tel(ByVal vNewValue As Variant)

On Error Resume Next

    vTel = vNewValue

End Property

Public Property Get fax() As Variant

    fax = vFax

End Property

Public Property Let fax(ByVal vNewValue As Variant)

On Error Resume Next

    vFax = vNewValue

End Property

Public Property Get email() As Variant

    email = vEmail

End Property

Public Property Let email(ByVal vNewValue As Variant)

On Error Resume Next

    vEmail = vNewValue

End Property

Public Property Get celular() As Variant

    celular = vCelular

End Property

Public Property Let celular(ByVal vNewValue As Variant)

On Error Resume Next

    vCelular = vNewValue

End Property

Public Property Get posicion() As Variant

    posicion = vPosicion

End Property

Public Property Let posicion(ByVal vNewValue As Variant)

On Error Resume Next

    vPosicion = vNewValue

End Property

Public Property Get tipodoc() As Variant

    tipodoc = vTipodoc

End Property

Public Property Let tipodoc(ByVal vNewValue As Variant)

On Error Resume Next

    vTipodoc = vNewValue

End Property

Public Property Get nrodoc() As Variant

    nrodoc = vNrodoc

End Property

Public Property Let nrodoc(ByVal vNewValue As Variant)

On Error Resume Next

    vNrodoc = vNewValue

End Property

Public Property Get limitecredito() As Variant

    limitecredito = vLimitecredito

End Property

Public Property Let limitecredito(ByVal vNewValue As Variant)

On Error Resume Next

    vLimitecredito = vNewValue

End Property

Public Property Get nacionalidad() As Variant

    nacionalidad = vNacionalidad

End Property

Public Property Let nacionalidad(ByVal vNewValue As Variant)

On Error Resume Next

    vNacionalidad = vNewValue

End Property

Public Property Get descuento() As Variant

    descuento = vDescuento

End Property

Public Property Let descuento(ByVal vNewValue As Variant)

On Error Resume Next

    vDescuento = vNewValue

End Property

Public Property Get facturable() As Variant

    facturable = vFacturable

End Property

Public Property Let facturable(ByVal vNewValue As Variant)

On Error Resume Next

    vFacturable = vNewValue

End Property

Public Property Get descripcionBuscar()

    descripcionBuscar = vRazon & " (" & vCuit & ")"
    
End Property

Public Property Get comboText() As Variant

    comboText = vRazon

End Property

Public Property Get comboKey() As Variant

    comboKey = vCodigo

End Property

Public Property Get listText() As Variant

    listText = vRazon & " (" & vCodigo & ")"

End Property

Public Property Get listKey() As Variant

    listKey = vCodigo

End Property

Public Function clone() As clsDAOClientes
Dim objObjeto As New clsDAOClientes

    With objObjeto
        .clave = vClave
        .negID = vNegid
        .codigo = vCodigo
        .cuit = vCuit
        .razon = vRazon
        .domicilio = vDomicilio
        .localidad = vLocalidad
        .provincia = vProvincia
        .tel = vTel
        .fax = vFax
        .email = vEmail
        .celular = vCelular
        .posicion = vPosicion
        .tipodoc = vTipodoc
        .nrodoc = vNrodoc
        .limitecredito = vLimitecredito
        .nacionalidad = vNacionalidad
        .descuento = vDescuento
        .facturable = vFacturable
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOClientes, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .negID = pQuery.Fields(.ddl.negID)
        .codigo = pQuery.Fields(.ddl.codigo)
        .cuit = pQuery.Fields(.ddl.cuit)
        .razon = pQuery.Fields(.ddl.razon)
        .domicilio = pQuery.Fields(.ddl.domicilio)
        .localidad = pQuery.Fields(.ddl.localidad)
        .provincia = pQuery.Fields(.ddl.provincia)
        .tel = pQuery.Fields(.ddl.tel)
        .fax = pQuery.Fields(.ddl.fax)
        .email = pQuery.Fields(.ddl.email)
        .celular = pQuery.Fields(.ddl.celular)
        .posicion = pQuery.Fields(.ddl.posicion)
        .tipodoc = pQuery.Fields(.ddl.tipodoc)
        .nrodoc = pQuery.Fields(.ddl.nrodoc)
        .limitecredito = pQuery.Fields(.ddl.limitecredito)
        .nacionalidad = pQuery.Fields(.ddl.nacionalidad)
        .descuento = pQuery.Fields(.ddl.descuento)
        .facturable = pQuery.Fields(.ddl.facturable)
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

Public Sub findByPrimaryKey(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & makeWherePrimaryKey
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByCodigo(ByVal pCodigo As Long, pDB As clsDB)
Dim strSQL As String

    clean
    
    vCodigo = pCodigo
    
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.codigo & " = " & vCodigo
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub findByCUIT(ByVal pCUIT As String, pDB As clsDB)
Dim strSQL As String

    clean
    
    vCuit = pCUIT
    
    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cuit & " = '" & vCuit & "'"
    strSQL = strSQL & ";"

    findSQL strSQL, pDB

End Sub

Public Sub cargarBuscar(ByVal pCodigo As Long, pDB As clsDB)
    
    findByCodigo pCodigo, pDB
    
End Sub

Public Sub findLastID(pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    clean

    strSQL = "SELECT LAST_INSERT_ID() AS " & ddl.clave
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vClave = rstQuery.Fields(ddl.clave)

    rstQuery.Close

    findByPrimaryKey pDB

End Sub

Public Sub findLast(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " ORDER BY " & ddl.clave & " DESC"
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

Public Function existByCUIT(pDB As clsDB, Optional pNombre As Variant) As Boolean
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    If vCuit = "00-00000000-0" Then
        existByCUIT = False
        Exit Function
    End If

    pNombre = ""

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.cuit & " = '" & vCuit & "'"
    strSQL = strSQL & " AND " & ddl.codigo & " <> '" & vCodigo & "'"
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    existByCUIT = Not rstQuery.EOF
    If Not rstQuery.EOF Then pNombre = rstQuery.Fields(ddl.razon)
    rstQuery.Close
    
End Function

Public Function add(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "INSERT INTO " & ddl.table
    strSQL = strSQL & " SET " & ddl.negID & " = " & vNegid
    strSQL = strSQL & ", " & ddl.codigo & " = " & vCodigo
    strSQL = strSQL & ", " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
    strSQL = strSQL & ", " & ddl.razon & " = '" & replaceSpecialSymbols(vRazon) & "'"
    strSQL = strSQL & ", " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
    strSQL = strSQL & ", " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
    strSQL = strSQL & ", " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
    strSQL = strSQL & ", " & ddl.tel & " = '" & replaceSpecialSymbols(vTel) & "'"
    strSQL = strSQL & ", " & ddl.fax & " = '" & replaceSpecialSymbols(vFax) & "'"
    strSQL = strSQL & ", " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
    strSQL = strSQL & ", " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
    strSQL = strSQL & ", " & ddl.posicion & " = " & vPosicion
    strSQL = strSQL & ", " & ddl.tipodoc & " = '" & replaceSpecialSymbols(vTipodoc) & "'"
    strSQL = strSQL & ", " & ddl.nrodoc & " = " & vNrodoc
    strSQL = strSQL & ", " & ddl.limitecredito & " = '" & Replace(vLimitecredito, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.nacionalidad & " = '" & replaceSpecialSymbols(vNacionalidad) & "'"
    strSQL = strSQL & ", " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.facturable & " = " & vFacturable
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

    If vNegid <> clon.negID Or IsNull(vNegid) <> IsNull(clon.negID) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.negID & " = " & vNegid
        blnChange = True
    End If
    If vCodigo <> clon.codigo Or IsNull(vCodigo) <> IsNull(clon.codigo) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.codigo & " = " & vCodigo
        blnChange = True
    End If
    If vCuit <> clon.cuit Or IsNull(vCuit) <> IsNull(clon.cuit) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.cuit & " = '" & replaceSpecialSymbols(vCuit) & "'"
        blnChange = True
    End If
    If vRazon <> clon.razon Or IsNull(vRazon) <> IsNull(clon.razon) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.razon & " = '" & replaceSpecialSymbols(vRazon) & "'"
        blnChange = True
    End If
    If vDomicilio <> clon.domicilio Or IsNull(vDomicilio) <> IsNull(clon.domicilio) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.domicilio & " = '" & replaceSpecialSymbols(vDomicilio) & "'"
        blnChange = True
    End If
    If vLocalidad <> clon.localidad Or IsNull(vLocalidad) <> IsNull(clon.localidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.localidad & " = '" & replaceSpecialSymbols(vLocalidad) & "'"
        blnChange = True
    End If
    If vProvincia <> clon.provincia Or IsNull(vProvincia) <> IsNull(clon.provincia) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.provincia & " = '" & replaceSpecialSymbols(vProvincia) & "'"
        blnChange = True
    End If
    If vTel <> clon.tel Or IsNull(vTel) <> IsNull(clon.tel) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tel & " = '" & replaceSpecialSymbols(vTel) & "'"
        blnChange = True
    End If
    If vFax <> clon.fax Or IsNull(vFax) <> IsNull(clon.fax) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.fax & " = '" & replaceSpecialSymbols(vFax) & "'"
        blnChange = True
    End If
    If vEmail <> clon.email Or IsNull(vEmail) <> IsNull(clon.email) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.email & " = '" & replaceSpecialSymbols(vEmail) & "'"
        blnChange = True
    End If
    If vCelular <> clon.celular Or IsNull(vCelular) <> IsNull(clon.celular) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.celular & " = '" & replaceSpecialSymbols(vCelular) & "'"
        blnChange = True
    End If
    If vPosicion <> clon.posicion Or IsNull(vPosicion) <> IsNull(clon.posicion) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.posicion & " = " & vPosicion
        blnChange = True
    End If
    If vTipodoc <> clon.tipodoc Or IsNull(vTipodoc) <> IsNull(clon.tipodoc) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.tipodoc & " = '" & replaceSpecialSymbols(vTipodoc) & "'"
        blnChange = True
    End If
    If vNrodoc <> clon.nrodoc Or IsNull(vNrodoc) <> IsNull(clon.nrodoc) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nrodoc & " = " & vNrodoc
        blnChange = True
    End If
    If vLimitecredito <> clon.limitecredito Or IsNull(vLimitecredito) <> IsNull(clon.limitecredito) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.limitecredito & " = '" & Replace(vLimitecredito, ",", ".") & "'"
        blnChange = True
    End If
    If vNacionalidad <> clon.nacionalidad Or IsNull(vNacionalidad) <> IsNull(clon.nacionalidad) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.nacionalidad & " = '" & replaceSpecialSymbols(vNacionalidad) & "'"
        blnChange = True
    End If
    If vDescuento <> clon.descuento Or IsNull(vDescuento) <> IsNull(clon.descuento) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.descuento & " = '" & Replace(vDescuento, ",", ".") & "'"
        blnChange = True
    End If
    If vFacturable <> clon.facturable Or IsNull(vFacturable) <> IsNull(clon.facturable) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.facturable & " = " & vFacturable
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
    strSQL = strSQL & " ORDER BY " & ddl.codigo & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCodigo = 1 + rstQuery.Fields(ddl.codigo)

    rstQuery.Close

    If pFind Then findByCodigo vCodigo, pDB

End Sub

Public Function collectionAll(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & ";"

    Set collectionAll = collectionSQL(strSQL, pDB)

End Function

Public Function collectionFacturable(pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.facturable & " = 1"
    strSQL = strSQL & ";"

    Set collectionFacturable = collectionSQL(strSQL, pDB)

End Function

Public Function collectionSQL(pSQL As String, pDB As clsDB) As Collection
Dim rstQuery As ADODB.Recordset

Dim objElemento As clsDAOClientes

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOClientes
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

Public Sub fillComboFacturable(pCombo As ComboBox, pDB As clsDB)

    fillCombo_ pCombo, collectionFacturable(pDB)

End Sub

Private Sub fillCombo_(pCombo As ComboBox, pCollection As Collection)
Dim objElemento As clsDAOClientes

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
Dim objElemento As clsDAOClientes

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Sub formBuscar(pForm As Form, pObjeto As Variant, pKeyAscii As Integer, pQueBusco As String)

    Load pForm
    
    With pForm
        .setObjeto pObjeto
        .txtCadena.Text = Chr(pKeyAscii)
        .Caption = "Búsqueda " & pQueBusco
        .Show vbModal
    End With
    
    Unload pForm
    
End Sub

Public Function coleccionBuscar(pBuscar As String, pDB As clsDB) As Collection
Dim strSQL As String

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.razon & " LIKE '%" & pBuscar & "%'"
    strSQL = strSQL & " AND " & ddl.facturable & " = 1"
    strSQL = strSQL & " ORDER BY " & ddl.razon
    If Len(pBuscar) < 4 Then strSQL = strSQL & " LIMIT 50"
    strSQL = strSQL & ";"
    
    Set coleccionBuscar = collectionSQL(strSQL, pDB)

End Function

Public Sub extremos(pMinimo As Long, pMaximo As Long, pDB As clsDB)
Dim strSQL As String

Dim rstQuery As ADODB.Recordset

    pMinimo = 0
    pMaximo = 0

    strSQL = "SELECT MIN(codigo) AS minimo, MAX(codigo) AS maximo FROM " & ddl.table
    strSQL = strSQL & ";"
    
    Set rstQuery = pDB.query(strSQL)
    If Not rstQuery.EOF Then
        pMinimo = rstQuery!minimo
        pMaximo = rstQuery!maximo
    End If
    rstQuery.Close
    
End Sub



