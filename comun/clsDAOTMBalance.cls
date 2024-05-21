VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOTMBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vClave As Long
Private vHwnd As Long
Private vCuenta As Long
Private vDeudor As Currency
Private vAcreedor As Currency
Private vSaldodeudor As Currency
Private vSaldoacreedor As Currency
Private vGrado As Currency

Private clon As clsDAOTMBalance

Public ddl As New clsDDLTMBalance

Private Sub class_Initialize()

    vHwnd = 0
    vCuenta = 0

    initializeOther

End Sub

Private Sub initializeOther()

    vClave = 0
    vDeudor = 0
    vAcreedor = 0
    vSaldodeudor = 0
    vSaldoacreedor = 0
    vGrado = 0

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

Public Property Get hwnd() As Variant

    hwnd = vHwnd

End Property

Public Property Let hwnd(ByVal vNewValue As Variant)

On Error Resume Next

    vHwnd = vNewValue

End Property

Public Property Get cuenta() As Variant

    cuenta = vCuenta

End Property

Public Property Let cuenta(ByVal vNewValue As Variant)

On Error Resume Next

    vCuenta = vNewValue

End Property

Public Property Get deudor() As Variant

    deudor = vDeudor

End Property

Public Property Let deudor(ByVal vNewValue As Variant)

On Error Resume Next

    vDeudor = vNewValue

End Property

Public Property Get acreedor() As Variant

    acreedor = vAcreedor

End Property

Public Property Let acreedor(ByVal vNewValue As Variant)

On Error Resume Next

    vAcreedor = vNewValue

End Property

Public Property Get saldodeudor() As Variant

    saldodeudor = vSaldodeudor

End Property

Public Property Let saldodeudor(ByVal vNewValue As Variant)

On Error Resume Next

    vSaldodeudor = vNewValue

End Property

Public Property Get saldoacreedor() As Variant

    saldoacreedor = vSaldoacreedor

End Property

Public Property Let saldoacreedor(ByVal vNewValue As Variant)

On Error Resume Next

    vSaldoacreedor = vNewValue

End Property

Public Property Get grado() As Variant

    grado = vGrado

End Property

Public Property Let grado(ByVal vNewValue As Variant)

On Error Resume Next

    vGrado = vNewValue

End Property

Public Property Get comboText() As Variant

    comboText = vClave

End Property

Public Property Get comboKey() As Variant

    comboKey = vCuenta

End Property

Public Property Get listText() As Variant

    listText = vClave

End Property

Public Property Get listKey() As Variant

    listKey = vCuenta

End Property

Public Function clone() As clsDAOTMBalance
Dim objObjeto As New clsDAOTMBalance

    With objObjeto
        .clave = vClave
        .hwnd = vHwnd
        .cuenta = vCuenta
        .deudor = vDeudor
        .acreedor = vAcreedor
        .saldodeudor = vSaldodeudor
        .saldoacreedor = vSaldoacreedor
        .grado = vGrado
    End With

    Set clone = objObjeto

End Function

Private Sub fillObject(pObjeto As clsDAOTMBalance, pQuery As ADODB.Recordset)

On Error Resume Next

    With pObjeto
        .clave = pQuery.Fields(.ddl.clave)
        .hwnd = pQuery.Fields(.ddl.hwnd)
        .cuenta = pQuery.Fields(.ddl.cuenta)
        .deudor = pQuery.Fields(.ddl.deudor)
        .acreedor = pQuery.Fields(.ddl.acreedor)
        .saldodeudor = pQuery.Fields(.ddl.saldodeudor)
        .saldoacreedor = pQuery.Fields(.ddl.saldoacreedor)
        .grado = pQuery.Fields(.ddl.grado)
    End With

End Sub

Private Function makeWherePrimaryKey() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & " AND " & ddl.cuenta & " = " & vCuenta

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

Public Sub findLast(Optional pDB As Variant)

    findLast_ IIf(IsMissing(pDB), DB, pDB)

End Sub

Private Sub findLast_(pDB As clsDB)
Dim strSQL As String

    initializeOther

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & " ORDER BY " & ddl.cuenta & " DESC"
    strSQL = strSQL & " LIMIT 1"
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
    strSQL = strSQL & " SET " & ddl.deudor & " = '" & Replace(vDeudor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.acreedor & " = '" & Replace(vAcreedor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.saldodeudor & " = '" & Replace(vSaldodeudor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.saldoacreedor & " = '" & Replace(vSaldoacreedor, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.grado & " = '" & Replace(vGrado, ",", ".") & "'"
    strSQL = strSQL & ", " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ", " & ddl.cuenta & " = " & vCuenta
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

    If vDeudor <> clon.deudor Or IsNull(vDeudor) <> IsNull(clon.deudor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.deudor & " = '" & Replace(vDeudor, ",", ".") & "'"
        blnChange = True
    End If
    If vAcreedor <> clon.acreedor Or IsNull(vAcreedor) <> IsNull(clon.acreedor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.acreedor & " = '" & Replace(vAcreedor, ",", ".") & "'"
        blnChange = True
    End If
    If vSaldodeudor <> clon.saldodeudor Or IsNull(vSaldodeudor) <> IsNull(clon.saldodeudor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.saldodeudor & " = '" & Replace(vSaldodeudor, ",", ".") & "'"
        blnChange = True
    End If
    If vSaldoacreedor <> clon.saldoacreedor Or IsNull(vSaldoacreedor) <> IsNull(clon.saldoacreedor) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.saldoacreedor & " = '" & Replace(vSaldoacreedor, ",", ".") & "'"
        blnChange = True
    End If
    If vGrado <> clon.grado Or IsNull(vGrado) <> IsNull(clon.grado) Then
        If strSET <> "" Then strSET = strSET & ", "
        strSET = strSET & " " & ddl.grado & " = '" & Replace(vGrado, ",", ".") & "'"
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

    vCuenta = 1

    strSQL = "SELECT " & ddl.table & ".* FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & " ORDER BY " & ddl.cuenta & " DESC"
    strSQL = strSQL & " LIMIT 1"
    strSQL = strSQL & ";"

    Set rstQuery = pDB.query(strSQL)

    If Not rstQuery.EOF Then vCuenta = 1 + rstQuery.Fields(ddl.cuenta)

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

Dim objElemento As clsDAOTMBalance

Dim colRegistros As New Collection

    Set rstQuery = pDB.query(pSQL)

    Do While Not rstQuery.EOF
        Set objElemento = New clsDAOTMBalance
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
Dim objElemento As clsDAOTMBalance

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
Dim objElemento As clsDAOTMBalance

    pList.Clear

    For Each objElemento In pCollection
        pList.AddItem objElemento.listText
        pList.ItemData(pList.NewIndex) = objElemento.listKey
    Next

End Sub

Public Sub saldosImportesConAnterior(phWnd As Long, pDesde As Date, pHasta As Date, pDB As clsDB)
Dim strSQL As String

    ' Carga todos los clientes
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldodeudor)"
    strSQL = strSQL & " SELECT " & phWnd - 1 & ", codigo, 0 AS debitos FROM clientes"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL
    
    ' Calcula Saldo Anterior
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldodeudor)"
    strSQL = strSQL & " SELECT " & phWnd - 1 & ", cgoclie, SUM(importe) FROM movclie"
    strSQL = strSQL & " WHERE fechaComprob < " & fechaDB(pDesde)
    strSQL = strSQL & " AND anulada = 0"
    strSQL = strSQL & " GROUP BY cgoclie"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL
    
    ' Calcula Saldo Período
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldoacreedor)"
    strSQL = strSQL & " SELECT " & phWnd - 1 & ", cgoclie, SUM(importe) FROM movclie"
    strSQL = strSQL & " WHERE fechaComprob BETWEEN " & fechaDB(pDesde) & " AND " & fechaDB(pHasta)
    strSQL = strSQL & " AND anulada = 0"
    strSQL = strSQL & " GROUP BY cgoclie"
    strSQL = strSQL & ";"
    
    DB.activa.execute strSQL
    
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldodeudor, bal_saldoacreedor)"
    strSQL = strSQL & " SELECT " & phWnd & ", bal_cuenta, SUM(bal_saldodeudor), SUM(bal_saldoacreedor) FROM tmbalance"
    strSQL = strSQL & " WHERE bal_hwnd = " & phWnd - 1
    strSQL = strSQL & " GROUP BY bal_cuenta"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL

    vHwnd = phWnd - 1
        
    deleteAll pDB
    
End Sub

Public Sub saldosImportesSinAnterior(phWnd As Long, pDesde As Date, pHasta As Date, pDB As clsDB)
Dim strSQL As String

    ' Carga todos los clientes
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldodeudor)"
    strSQL = strSQL & " SELECT " & phWnd - 1 & ", codigo, 0 AS debitos FROM clientes"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL
    
    ' Calcula Saldo Período
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldoacreedor)"
    strSQL = strSQL & " SELECT " & phWnd - 1 & ", cgoclie, SUM(importe) FROM movclie"
    strSQL = strSQL & " WHERE fechaComprob BETWEEN " & fechaDB(pDesde) & " AND " & fechaDB(pHasta)
    strSQL = strSQL & " AND anulada = 0"
    strSQL = strSQL & " GROUP BY cgoclie"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL
    
    strSQL = "INSERT INTO tmbalance (bal_hwnd, bal_cuenta, bal_saldodeudor, bal_saldoacreedor)"
    strSQL = strSQL & " SELECT " & phWnd & ", bal_cuenta, SUM(bal_saldodeudor), SUM(bal_saldoacreedor) FROM tmbalance"
    strSQL = strSQL & " WHERE bal_hwnd = " & phWnd - 1
    strSQL = strSQL & " GROUP BY bal_cuenta"
    strSQL = strSQL & ";"
    
    pDB.activa.execute strSQL

    vHwnd = phWnd - 1
        
    deleteAll pDB
    
End Sub

Public Function deleteAll(pDB As clsDB) As Boolean
Dim strSQL As String

    strSQL = "DELETE FROM " & ddl.table
    strSQL = strSQL & " WHERE " & ddl.hwnd & " = " & vHwnd
    strSQL = strSQL & ";"
    
    deleteAll = pDB.execute(strSQL)
    
End Function
